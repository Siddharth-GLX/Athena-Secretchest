import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path


# ============================================================
# USER INPUT SETTINGS
# ============================================================

# Excel input file
FILE = "EPS Charge Discharge Test.xlsx"
SHEET_NAME = "Discharging"

# Column names in your Excel file
TIMESTAMP_COL = "TIMESTAMP"
VOLTAGE_COL = "TOTAL_BATTERY_VOLTAGE"
CURRENT_COL = "DISCHARGE_CURRENT"

# Output folder
OUTPUT_DIR = Path("battery_voltage_window_outputs")


# ============================================================
# BATTERY PACK ASSUMPTIONS
# ============================================================

# Panasonic NCR18650B assumed values
DOD = 0.18
CELL_NOMINAL_VOLTAGE_V = 3.6
CELL_CAPACITY_AH = 3.0
CELL_CAPACITY_WH = CELL_NOMINAL_VOLTAGE_V * CELL_CAPACITY_AH
CELL_CAPACITY_DERATED_WH = CELL_CAPACITY_WH * DOD

# Pack configuration
SERIES_CELLS = 8
PARALLEL_CELLS = 15
CELLS_NEEDED = SERIES_CELLS * PARALLEL_CELLS

PACK_NOMINAL_VOLTAGE_V = SERIES_CELLS * CELL_NOMINAL_VOLTAGE_V
PACK_CAPACITY_AH = PARALLEL_CELLS * CELL_CAPACITY_AH
PACK_NOMINAL_ENERGY_WH = PACK_NOMINAL_VOLTAGE_V * PACK_CAPACITY_AH

# This should be 1296 Wh for 8S15P, 3.6 V, 3 Ah cells
# 8 * 3.6 V = 28.8 V
# 15 * 3 Ah = 45 Ah
# 28.8 V * 45 Ah = 1296 Wh


# ============================================================
# SOLAR / ORBIT ENERGY ASSUMPTIONS
# ============================================================

SOLAR_GENERATION_WH = 476.0
MPPT_LOSS_FRACTION = 0.10
SUNLIT_LOAD_WH = 224.0
PACK_RTE_LOSS_FRACTION = 0.10

SOLAR_AFTER_MPPT_WH = SOLAR_GENERATION_WH * (1.0 - MPPT_LOSS_FRACTION)
ENERGY_TO_BATTERY_BEFORE_RTE_WH = SOLAR_AFTER_MPPT_WH - SUNLIT_LOAD_WH
ENERGY_STORED_AFTER_RTE_WH = ENERGY_TO_BATTERY_BEFORE_RTE_WH * (
    1.0 - PACK_RTE_LOSS_FRACTION
)

# Your required voltage-window condition
ENERGY_BETWEEN_POINTS_WH = 204.0
ENERGY_FROM_LOWER_TO_DRAIN_WH = 504.0

# Therefore:
# lower voltage point has 504 Wh remaining to drain
# upper voltage point has 504 + 204 = 708 Wh remaining to drain
LOWER_REMAINING_ENERGY_WH = ENERGY_FROM_LOWER_TO_DRAIN_WH
UPPER_REMAINING_ENERGY_WH = (
    ENERGY_FROM_LOWER_TO_DRAIN_WH + ENERGY_BETWEEN_POINTS_WH
)


# ============================================================
# IMPORTANT OPTION
# ============================================================

# Keep this True if your measured test data does not represent the full 1296 Wh pack.
# Example: if your measured data only shows around 204 Wh, this scales the curve
# to the assumed 1296 Wh pack.
SCALE_MEASURED_CURVE_TO_PACK_ENERGY = True


# ============================================================
# FUNCTIONS
# ============================================================

def load_and_process_excel():
    """
    Loads Excel data and calculates cumulative discharged energy.
    """

    df = pd.read_excel(FILE, sheet_name=SHEET_NAME)

    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]

    required_cols = [TIMESTAMP_COL, VOLTAGE_COL, CURRENT_COL]

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        raise ValueError(
            f"Missing required columns: {missing_cols}\n"
            f"Available columns are: {list(df.columns)}"
        )

    df = df[[TIMESTAMP_COL, VOLTAGE_COL, CURRENT_COL]].copy()

    df[TIMESTAMP_COL] = pd.to_numeric(df[TIMESTAMP_COL], errors="coerce")
    df[VOLTAGE_COL] = pd.to_numeric(df[VOLTAGE_COL], errors="coerce")
    df[CURRENT_COL] = pd.to_numeric(df[CURRENT_COL], errors="coerce")

    df = df.dropna(subset=[TIMESTAMP_COL, VOLTAGE_COL, CURRENT_COL])
    df = df.sort_values(TIMESTAMP_COL).reset_index(drop=True)

    if len(df) < 2:
        raise ValueError("Not enough valid data points. Need at least 2 rows.")

    # Convert timestamp to readable UTC time
    df["TIMESTAMP_UTC_DT"] = pd.to_datetime(
        df[TIMESTAMP_COL],
        unit="s",
        utc=True
    ).dt.tz_localize(None)

    df["TIMESTAMP_UTC"] = df["TIMESTAMP_UTC_DT"].dt.strftime(
        "%Y-%m-%d %H:%M:%S UTC"
    )

    # Time step between samples
    df["dt_sec"] = df[TIMESTAMP_COL].diff().fillna(0)
    df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

    # Use absolute current because discharge current may be negative
    df["Current_abs_A"] = df[CURRENT_COL].abs()

    # Power = Voltage * Current
    df["Power_W"] = df[VOLTAGE_COL] * df["Current_abs_A"]

    # Incremental energy in Wh
    df["d_Wh"] = df["Power_W"] * df["dt_sec"] / 3600.0

    # Cumulative discharged energy in Wh
    df["Energy_Discharged_Wh"] = df["d_Wh"].cumsum()

    # Capacity in Ah and mAh
    df["d_Ah"] = df["Current_abs_A"] * df["dt_sec"] / 3600.0
    df["Capacity_Ah"] = df["d_Ah"].cumsum()
    df["Capacity_mAh"] = df["Capacity_Ah"] * 1000.0

    return df


def create_remaining_energy_curve(df):
    """
    Creates a voltage vs remaining-energy curve.

    The measured data gives discharged energy from the start of the test.
    We convert it to remaining energy:

        Remaining Energy = Pack Energy - Energy Discharged

    If scaling is enabled, the measured test curve is scaled to 1296 Wh.
    """

    curve = df[[VOLTAGE_COL, "Energy_Discharged_Wh"]].copy()
    curve = curve.dropna()

    measured_total_energy_wh = curve["Energy_Discharged_Wh"].max()

    if measured_total_energy_wh <= 0:
        raise ValueError("Measured discharge energy is zero or negative.")

    if SCALE_MEASURED_CURVE_TO_PACK_ENERGY:
        curve["Energy_From_Full_Wh"] = (
            curve["Energy_Discharged_Wh"] / measured_total_energy_wh
        ) * PACK_NOMINAL_ENERGY_WH
    else:
        curve["Energy_From_Full_Wh"] = curve["Energy_Discharged_Wh"]

    curve["Remaining_Energy_Wh"] = (
        PACK_NOMINAL_ENERGY_WH - curve["Energy_From_Full_Wh"]
    )

    curve["SOC_percent"] = (
        curve["Remaining_Energy_Wh"] / PACK_NOMINAL_ENERGY_WH
    ) * 100.0

    # Sort for interpolation
    curve = curve.sort_values("Remaining_Energy_Wh").reset_index(drop=True)

    return curve, measured_total_energy_wh


def find_voltage_at_remaining_energy(curve, target_remaining_energy_wh):
    """
    Interpolates the pack voltage at a target remaining energy.
    """

    min_energy = curve["Remaining_Energy_Wh"].min()
    max_energy = curve["Remaining_Energy_Wh"].max()

    if target_remaining_energy_wh < min_energy or target_remaining_energy_wh > max_energy:
        raise ValueError(
            f"Target remaining energy {target_remaining_energy_wh:.2f} Wh "
            f"is outside curve range {min_energy:.2f} Wh to {max_energy:.2f} Wh."
        )

    voltage = np.interp(
        target_remaining_energy_wh,
        curve["Remaining_Energy_Wh"],
        curve[VOLTAGE_COL],
    )

    return voltage


def calculate_voltage_window(curve):
    """
    Finds the two voltage points:

    Lower point:
        Remaining energy to drain = 504 Wh

    Upper point:
        Remaining energy to drain = 504 Wh + 204 Wh = 708 Wh
    """

    if UPPER_REMAINING_ENERGY_WH > PACK_NOMINAL_ENERGY_WH:
        raise ValueError(
            "Upper remaining energy is greater than total pack energy. "
            "Check your input energy values."
        )

    lower_voltage = find_voltage_at_remaining_energy(
        curve,
        LOWER_REMAINING_ENERGY_WH
    )

    upper_voltage = find_voltage_at_remaining_energy(
        curve,
        UPPER_REMAINING_ENERGY_WH
    )

    result = {
        "Pack nominal voltage V": PACK_NOMINAL_VOLTAGE_V,
        "Pack capacity Ah": PACK_CAPACITY_AH,
        "Pack nominal energy Wh": PACK_NOMINAL_ENERGY_WH,
        "Cell nominal voltage V": CELL_NOMINAL_VOLTAGE_V,
        "Cell capacity Ah": CELL_CAPACITY_AH,
        "Cell capacity Wh": CELL_CAPACITY_WH,
        "DoD": DOD,
        "Cell capacity derated Wh": CELL_CAPACITY_DERATED_WH,
        "Series cells": SERIES_CELLS,
        "Parallel cells": PARALLEL_CELLS,
        "Cells needed": CELLS_NEEDED,
        "Solar generation Wh": SOLAR_GENERATION_WH,
        "Solar after MPPT Wh": SOLAR_AFTER_MPPT_WH,
        "Sunlit load Wh": SUNLIT_LOAD_WH,
        "Energy to battery before RTE Wh": ENERGY_TO_BATTERY_BEFORE_RTE_WH,
        "Energy stored after RTE Wh": ENERGY_STORED_AFTER_RTE_WH,
        "Energy between voltage points Wh": ENERGY_BETWEEN_POINTS_WH,
        "Energy from lower point to drain Wh": ENERGY_FROM_LOWER_TO_DRAIN_WH,
        "Upper remaining energy Wh": UPPER_REMAINING_ENERGY_WH,
        "Lower remaining energy Wh": LOWER_REMAINING_ENERGY_WH,
        "Upper SOC percent": (
            UPPER_REMAINING_ENERGY_WH / PACK_NOMINAL_ENERGY_WH
        ) * 100.0,
        "Lower SOC percent": (
            LOWER_REMAINING_ENERGY_WH / PACK_NOMINAL_ENERGY_WH
        ) * 100.0,
        "Upper voltage point V": upper_voltage,
        "Lower voltage point V": lower_voltage,
    }

    return result


def save_results_to_excel(processed_df, curve_df, result):
    """
    Saves processed data, remaining-energy curve, and final result to Excel.
    """

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    result_df = pd.DataFrame(
        list(result.items()),
        columns=["Parameter", "Value"]
    )

    output_file = OUTPUT_DIR / "Battery_Voltage_Window_Result.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="Voltage_Window_Result", index=False)
        processed_df.to_excel(writer, sheet_name="Processed_Data", index=False)
        curve_df.to_excel(writer, sheet_name="Remaining_Energy_Curve", index=False)

    return output_file


def plot_voltage_vs_remaining_energy(curve_df, result):
    """
    Plots voltage vs remaining energy and marks the upper/lower voltage points.
    """

    output_file = OUTPUT_DIR / "Voltage_Window_vs_Remaining_Energy.png"

    fig, ax = plt.subplots(figsize=(11, 6))

    ax.plot(
        curve_df["Remaining_Energy_Wh"],
        curve_df[VOLTAGE_COL],
        linewidth=2,
        label="Voltage vs Remaining Energy"
    )

    ax.axvline(
        result["Upper remaining energy Wh"],
        linestyle="--",
        label=f"Upper remaining energy = {result['Upper remaining energy Wh']:.1f} Wh"
    )

    ax.axvline(
        result["Lower remaining energy Wh"],
        linestyle="--",
        label=f"Lower remaining energy = {result['Lower remaining energy Wh']:.1f} Wh"
    )

    ax.scatter(
        result["Upper remaining energy Wh"],
        result["Upper voltage point V"],
        s=80,
        zorder=3,
        label=f"Upper voltage = {result['Upper voltage point V']:.3f} V"
    )

    ax.scatter(
        result["Lower remaining energy Wh"],
        result["Lower voltage point V"],
        s=80,
        zorder=3,
        label=f"Lower voltage = {result['Lower voltage point V']:.3f} V"
    )

    ax.set_xlabel("Remaining Energy to Complete Drain (Wh)")
    ax.set_ylabel("Pack Voltage (V)")
    ax.set_title("Battery Voltage Window Based on Required Energy")
    ax.grid(True, alpha=0.3)
    ax.legend()

    fig.tight_layout()
    fig.savefig(output_file, dpi=300)
    plt.close(fig)

    return output_file


def plot_voltage_vs_measured_energy(processed_df):
    """
    Plots voltage vs measured discharged energy.
    """

    output_file = OUTPUT_DIR / "Voltage_vs_Measured_Discharged_Energy.png"

    fig, ax = plt.subplots(figsize=(11, 6))

    ax.plot(
        processed_df["Energy_Discharged_Wh"],
        processed_df[VOLTAGE_COL],
        linewidth=2,
        label="Voltage vs Measured Discharged Energy"
    )

    ax.set_xlabel("Measured Discharged Energy (Wh)")
    ax.set_ylabel("Pack Voltage (V)")
    ax.set_title("Voltage vs Measured Discharged Energy")
    ax.grid(True, alpha=0.3)
    ax.legend()

    fig.tight_layout()
    fig.savefig(output_file, dpi=300)
    plt.close(fig)

    return output_file


def plot_voltage_vs_time(processed_df):
    """
    Plots pack voltage vs time.
    """

    output_file = OUTPUT_DIR / "Voltage_vs_Time.png"

    fig, ax = plt.subplots(figsize=(11, 6))

    ax.plot(
        processed_df["TIMESTAMP_UTC_DT"],
        processed_df[VOLTAGE_COL],
        linewidth=2,
        label="Pack Voltage"
    )

    ax.set_xlabel("Timestamp UTC")
    ax.set_ylabel("Pack Voltage (V)")
    ax.set_title("Pack Voltage vs Time")
    ax.grid(True, alpha=0.3)
    ax.legend()

    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(output_file, dpi=300)
    plt.close(fig)

    return output_file


def print_result(result, measured_total_energy_wh):
    """
    Prints final result to terminal.
    """

    print()
    print("============================================================")
    print("BATTERY VOLTAGE WINDOW RESULT")
    print("============================================================")
    print()

    print("Pack configuration")
    print("------------------")
    print(f"Cell nominal voltage              : {CELL_NOMINAL_VOLTAGE_V:.3f} V")
    print(f"Cell capacity                     : {CELL_CAPACITY_AH:.3f} Ah")
    print(f"Cell capacity                     : {CELL_CAPACITY_WH:.3f} Wh")
    print(f"Cell derated capacity at DoD       : {CELL_CAPACITY_DERATED_WH:.3f} Wh")
    print(f"Series cells                      : {SERIES_CELLS}")
    print(f"Parallel cells                    : {PARALLEL_CELLS}")
    print(f"Cells needed                      : {CELLS_NEEDED}")
    print(f"Pack nominal voltage              : {PACK_NOMINAL_VOLTAGE_V:.3f} V")
    print(f"Pack capacity                     : {PACK_CAPACITY_AH:.3f} Ah")
    print(f"Pack nominal energy               : {PACK_NOMINAL_ENERGY_WH:.3f} Wh")
    print()

    print("Solar/orbit energy")
    print("------------------")
    print(f"Solar generation                  : {SOLAR_GENERATION_WH:.3f} Wh")
    print(f"MPPT loss                         : {MPPT_LOSS_FRACTION * 100:.1f} %")
    print(f"Solar after MPPT                  : {SOLAR_AFTER_MPPT_WH:.3f} Wh")
    print(f"Sunlit load                       : {SUNLIT_LOAD_WH:.3f} Wh")
    print(f"Energy to battery before RTE      : {ENERGY_TO_BATTERY_BEFORE_RTE_WH:.3f} Wh")
    print(f"Pack RTE/storage loss             : {PACK_RTE_LOSS_FRACTION * 100:.1f} %")
    print(f"Energy stored after RTE           : {ENERGY_STORED_AFTER_RTE_WH:.3f} Wh")
    print()

    print("Voltage window condition")
    print("------------------------")
    print(f"Energy between voltage points     : {ENERGY_BETWEEN_POINTS_WH:.3f} Wh")
    print(f"Energy below lower point to drain : {ENERGY_FROM_LOWER_TO_DRAIN_WH:.3f} Wh")
    print(f"Upper remaining energy            : {UPPER_REMAINING_ENERGY_WH:.3f} Wh")
    print(f"Lower remaining energy            : {LOWER_REMAINING_ENERGY_WH:.3f} Wh")
    print(f"Upper SOC                         : {result['Upper SOC percent']:.3f} %")
    print(f"Lower SOC                         : {result['Lower SOC percent']:.3f} %")
    print()

    print("Calculated voltage points")
    print("-------------------------")
    print(f"Upper voltage point               : {result['Upper voltage point V']:.3f} V")
    print(f"Lower voltage point               : {result['Lower voltage point V']:.3f} V")
    print()

    print("Dataset")
    print("-------")
    print(f"Measured total discharge energy   : {measured_total_energy_wh:.3f} Wh")
    print(f"Scaling enabled                   : {SCALE_MEASURED_CURVE_TO_PACK_ENERGY}")
    print()

    print("Check")
    print("-----")
    print(
        f"Energy between points check       : "
        f"{result['Upper remaining energy Wh'] - result['Lower remaining energy Wh']:.3f} Wh"
    )
    print(
        f"Energy below lower point check    : "
        f"{result['Lower remaining energy Wh']:.3f} Wh"
    )
    print()


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    processed_df = load_and_process_excel()

    curve_df, measured_total_energy_wh = create_remaining_energy_curve(
        processed_df
    )

    result = calculate_voltage_window(curve_df)

    excel_file = save_results_to_excel(
        processed_df,
        curve_df,
        result
    )

    plot_1 = plot_voltage_vs_remaining_energy(
        curve_df,
        result
    )

    plot_2 = plot_voltage_vs_measured_energy(
        processed_df
    )

    plot_3 = plot_voltage_vs_time(
        processed_df
    )

    print_result(result, measured_total_energy_wh)

    print("Files created")
    print("-------------")
    print(f"Excel result file : {excel_file}")
    print(f"Plot 1            : {plot_1}")
    print(f"Plot 2            : {plot_2}")
    print(f"Plot 3            : {plot_3}")
    print()
    print("Done.")


if __name__ == "__main__":
    main()
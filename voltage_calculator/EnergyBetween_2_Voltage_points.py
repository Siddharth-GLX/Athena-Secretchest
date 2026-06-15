import pandas as pd
from pathlib import Path
BASE_DIR = Path(__file__).parent

FILE = BASE_DIR / "EPS_Capacity_Test_Processed_1.2V.xlsx"

SHEET_NAME = None

ASSUMED_DISCHARGE_CURRENT_A = 1.5

ASSUMED_SAMPLE_PERIOD_SEC = 1.0

POINT_A_VOLTAGE = 29.0
POINT_B_VOLTAGE = 24.0

VOLTAGE_OFFSET_V = 0.0


# LOAD EXCEL DATA

def load_excel_data(file_path, sheet_name=None):
    """
    Loads the Excel file and automatically selects useful columns.

    Voltage column priority:
    1. CORRECTED_BATTERY_VOLTAGE
    2. VOLTAGE_USED_FOR_CALCULATION
    3. TOTAL_BATTERY_VOLTAGE
    4. MEASURED_BATTERY_VOLTAGE

    Time handling priority:
    1. TIMESTAMP
    2. ELAPSED_TIME_SEC
    3. ELAPSED_TIME_MIN
    4. dt_sec
    5. SAMPLE_INDEX
    6. fallback assumed sample period
    """

    if sheet_name is None:
        excel_file = pd.ExcelFile(file_path)
        sheet_name = excel_file.sheet_names[0]
        print(f"Using sheet: {sheet_name}")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    voltage_column_options = [
        "CORRECTED_BATTERY_VOLTAGE",
        "VOLTAGE_USED_FOR_CALCULATION",
        "TOTAL_BATTERY_VOLTAGE",
        "MEASURED_BATTERY_VOLTAGE",
    ]

    voltage_col = None

    for col in voltage_column_options:
        if col in df.columns:
            voltage_col = col
            break

    if voltage_col is None:
        raise ValueError(
            "No voltage column found. Expected one of these columns: "
            "CORRECTED_BATTERY_VOLTAGE, VOLTAGE_USED_FOR_CALCULATION, "
            "TOTAL_BATTERY_VOLTAGE, MEASURED_BATTERY_VOLTAGE"
        )

    df[voltage_col] = pd.to_numeric(df[voltage_col], errors="coerce")
    df = df.dropna(subset=[voltage_col]).reset_index(drop=True)

    df["VOLTAGE_USED_FOR_CALCULATION"] = df[voltage_col] + VOLTAGE_OFFSET_V

        # Time handling

    if "TIMESTAMP" in df.columns:
        df["TIMESTAMP"] = pd.to_numeric(df["TIMESTAMP"], errors="coerce")
        df = df.dropna(subset=["TIMESTAMP"]).reset_index(drop=True)
        df = df.sort_values("TIMESTAMP").reset_index(drop=True)

        df["dt_sec"] = df["TIMESTAMP"].diff().fillna(0)
        df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

        df["TIME_DISPLAY"] = pd.to_datetime(
            df["TIMESTAMP"],
            unit="s",
            utc=True
        ).dt.strftime("%Y-%m-%d %H:%M:%S UTC")

    elif "ELAPSED_TIME_SEC" in df.columns:
        df["ELAPSED_TIME_SEC"] = pd.to_numeric(
            df["ELAPSED_TIME_SEC"],
            errors="coerce"
        )
        df = df.dropna(subset=["ELAPSED_TIME_SEC"]).reset_index(drop=True)
        df = df.sort_values("ELAPSED_TIME_SEC").reset_index(drop=True)

        df["dt_sec"] = df["ELAPSED_TIME_SEC"].diff().fillna(0)
        df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

        df["TIME_DISPLAY"] = df["ELAPSED_TIME_SEC"].round(3).astype(str) + " sec"

    elif "ELAPSED_TIME_MIN" in df.columns:
        df["ELAPSED_TIME_MIN"] = pd.to_numeric(
            df["ELAPSED_TIME_MIN"],
            errors="coerce"
        )
        df = df.dropna(subset=["ELAPSED_TIME_MIN"]).reset_index(drop=True)
        df = df.sort_values("ELAPSED_TIME_MIN").reset_index(drop=True)

        df["ELAPSED_TIME_SEC"] = df["ELAPSED_TIME_MIN"] * 60.0
        df["dt_sec"] = df["ELAPSED_TIME_SEC"].diff().fillna(0)
        df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

        df["TIME_DISPLAY"] = df["ELAPSED_TIME_MIN"].round(3).astype(str) + " min"

    elif "dt_sec" in df.columns:
        df["dt_sec"] = pd.to_numeric(df["dt_sec"], errors="coerce")
        df["dt_sec"] = df["dt_sec"].fillna(ASSUMED_SAMPLE_PERIOD_SEC)
        df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

        df["ELAPSED_TIME_SEC"] = df["dt_sec"].cumsum()
        df.loc[0, "ELAPSED_TIME_SEC"] = 0.0

        df["TIME_DISPLAY"] = df["ELAPSED_TIME_SEC"].round(3).astype(str) + " sec"

    elif "SAMPLE_INDEX" in df.columns:
        df["SAMPLE_INDEX"] = pd.to_numeric(df["SAMPLE_INDEX"], errors="coerce")
        df = df.dropna(subset=["SAMPLE_INDEX"]).reset_index(drop=True)
        df = df.sort_values("SAMPLE_INDEX").reset_index(drop=True)

        df["dt_sec"] = ASSUMED_SAMPLE_PERIOD_SEC
        df.loc[0, "dt_sec"] = 0.0

        df["ELAPSED_TIME_SEC"] = df["SAMPLE_INDEX"] * ASSUMED_SAMPLE_PERIOD_SEC
        df["TIME_DISPLAY"] = df["ELAPSED_TIME_SEC"].round(3).astype(str) + " sec"

    else:
        print()
        print("Warning: No time column found.")
        print(f"Using assumed sample period: {ASSUMED_SAMPLE_PERIOD_SEC} sec per row")

        df["dt_sec"] = ASSUMED_SAMPLE_PERIOD_SEC
        df.loc[0, "dt_sec"] = 0.0

        df["ELAPSED_TIME_SEC"] = df.index * ASSUMED_SAMPLE_PERIOD_SEC
        df["TIME_DISPLAY"] = df["ELAPSED_TIME_SEC"].round(3).astype(str) + " sec"

    return df, voltage_col
    """
    Loads the Excel file and automatically selects useful columns.

    The code looks for voltage columns in this order:
    1. CORRECTED_BATTERY_VOLTAGE
    2. VOLTAGE_USED_FOR_CALCULATION
    3. TOTAL_BATTERY_VOLTAGE
    4. MEASURED_BATTERY_VOLTAGE

    It also uses timestamp or elapsed time if available.
    """

    if sheet_name is None:
        excel_file = pd.ExcelFile(file_path)
        sheet_name = excel_file.sheet_names[0]
        print(f"Using sheet: {sheet_name}")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    voltage_column_options = [
        "CORRECTED_BATTERY_VOLTAGE",
        "VOLTAGE_USED_FOR_CALCULATION",
        "TOTAL_BATTERY_VOLTAGE",
        "MEASURED_BATTERY_VOLTAGE",
    ]

    voltage_col = None

    for col in voltage_column_options:
        if col in df.columns:
            voltage_col = col
            break

    if voltage_col is None:
        raise ValueError(
            "No voltage column found. Expected one of these columns: "
            "CORRECTED_BATTERY_VOLTAGE, VOLTAGE_USED_FOR_CALCULATION, "
            "TOTAL_BATTERY_VOLTAGE, MEASURED_BATTERY_VOLTAGE"
        )

    df[voltage_col] = pd.to_numeric(df[voltage_col], errors="coerce")

    df = df.dropna(subset=[voltage_col]).reset_index(drop=True)

    df["VOLTAGE_USED_FOR_CALCULATION"] = df[voltage_col] + VOLTAGE_OFFSET_V

    # --------------------------------------------------------
    # Time handling
    # --------------------------------------------------------

    if "TIMESTAMP" in df.columns:
        df["TIMESTAMP"] = pd.to_numeric(df["TIMESTAMP"], errors="coerce")
        df = df.dropna(subset=["TIMESTAMP"]).reset_index(drop=True)

        df = df.sort_values("TIMESTAMP").reset_index(drop=True)

        df["dt_sec"] = df["TIMESTAMP"].diff().fillna(0)
        df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

        df["TIME_DISPLAY"] = pd.to_datetime(
            df["TIMESTAMP"],
            unit="s",
            utc=True
        ).dt.strftime("%Y-%m-%d %H:%M:%S UTC")

    elif "ELAPSED_TIME_SEC" in df.columns:
        df["ELAPSED_TIME_SEC"] = pd.to_numeric(
            df["ELAPSED_TIME_SEC"],
            errors="coerce"
        )
        df = df.dropna(subset=["ELAPSED_TIME_SEC"]).reset_index(drop=True)

        df = df.sort_values("ELAPSED_TIME_SEC").reset_index(drop=True)

        df["dt_sec"] = df["ELAPSED_TIME_SEC"].diff().fillna(0)
        df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

        df["TIME_DISPLAY"] = df["ELAPSED_TIME_SEC"].astype(str) + " sec"

    elif "ELAPSED_TIME_MIN" in df.columns:
        df["ELAPSED_TIME_MIN"] = pd.to_numeric(
            df["ELAPSED_TIME_MIN"],
            errors="coerce"
        )
        df = df.dropna(subset=["ELAPSED_TIME_MIN"]).reset_index(drop=True)

        df = df.sort_values("ELAPSED_TIME_MIN").reset_index(drop=True)

        df["ELAPSED_TIME_SEC"] = df["ELAPSED_TIME_MIN"] * 60.0
        df["dt_sec"] = df["ELAPSED_TIME_SEC"].diff().fillna(0)
        df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

        df["TIME_DISPLAY"] = df["ELAPSED_TIME_MIN"].astype(str) + " min"

    else:
        raise ValueError(
            "No time column found. Expected TIMESTAMP, ELAPSED_TIME_SEC, "
            "or ELAPSED_TIME_MIN."
        )

    return df, voltage_col


# FIND VOLTAGE POINTS DURING DISCHARGE

def find_voltage_index_for_discharge(df, target_voltage):
    """
    Finds the row index closest to the given voltage.

    This works for discharge data where voltage generally decreases
    from high voltage to low voltage.
    """

    difference = (
        df["VOLTAGE_USED_FOR_CALCULATION"] - target_voltage
    ).abs()

    closest_index = difference.idxmin()

    return closest_index


# ============================================================
# CALCULATE ENERGY BETWEEN VOLTAGE POINTS
# ============================================================

def calculate_energy_between_voltage_points(
    df,
    point_a_voltage,
    point_b_voltage,
    assumed_current_a
):
    """
    Calculates energy consumed between two voltage points.

    Preferred method:
        If Capacity_Ah exists, use:
            Energy Wh = Average Voltage x Delta Capacity Ah

    Fallback method:
        If Capacity_Ah does not exist, use:
            Energy Wh = Voltage x Current x Time
    """

    point_a_index = find_voltage_index_for_discharge(df, point_a_voltage)
    point_b_index = find_voltage_index_for_discharge(df, point_b_voltage)

    start_index = min(point_a_index, point_b_index)
    end_index = max(point_a_index, point_b_index)

    if start_index == end_index:
        raise ValueError(
            "Point A and Point B matched the same data row. "
            "Choose voltage values farther apart."
        )

    section = df.loc[start_index:end_index].copy()

    matched_point_a_voltage = df.loc[
        point_a_index,
        "VOLTAGE_USED_FOR_CALCULATION"
    ]

    matched_point_b_voltage = df.loc[
        point_b_index,
        "VOLTAGE_USED_FOR_CALCULATION"
    ]

    start_voltage = section["VOLTAGE_USED_FOR_CALCULATION"].iloc[0]
    end_voltage = section["VOLTAGE_USED_FOR_CALCULATION"].iloc[-1]
    average_voltage_v = section["VOLTAGE_USED_FOR_CALCULATION"].mean()

    # ========================================================
    # METHOD 1: Use Capacity_Ah if available
    # ========================================================
    if "Capacity_Ah" in section.columns:
        section["Capacity_Ah"] = pd.to_numeric(
            section["Capacity_Ah"],
            errors="coerce"
        )

        section = section.dropna(subset=["Capacity_Ah"])

        start_capacity_ah = section["Capacity_Ah"].iloc[0]
        end_capacity_ah = section["Capacity_Ah"].iloc[-1]

        capacity_consumed_ah = abs(end_capacity_ah - start_capacity_ah)

        energy_consumed_wh = average_voltage_v * capacity_consumed_ah

        calculation_method = "Capacity based: Energy Wh = Average Voltage x Delta Capacity Ah"

        total_time_hr = capacity_consumed_ah / assumed_current_a
        total_time_sec = total_time_hr * 3600.0

    # ========================================================
    # METHOD 2: Fallback to time-based method
    # ========================================================
    else:
        section["ASSUMED_DISCHARGE_CURRENT_A"] = assumed_current_a

        section["ASSUMED_POWER_W"] = (
            section["VOLTAGE_USED_FOR_CALCULATION"]
            * section["ASSUMED_DISCHARGE_CURRENT_A"]
        )

        section["ASSUMED_d_Wh"] = (
            section["ASSUMED_POWER_W"]
            * section["dt_sec"]
            / 3600.0
        )

        section.loc[section.index[0], "ASSUMED_d_Wh"] = 0.0

        section["ASSUMED_d_Ah"] = (
            section["ASSUMED_DISCHARGE_CURRENT_A"]
            * section["dt_sec"]
            / 3600.0
        )

        section.loc[section.index[0], "ASSUMED_d_Ah"] = 0.0

        energy_consumed_wh = section["ASSUMED_d_Wh"].sum()
        capacity_consumed_ah = section["ASSUMED_d_Ah"].sum()

        total_time_sec = section["dt_sec"].sum()
        total_time_hr = total_time_sec / 3600.0

        calculation_method = "Time based: Energy Wh = Voltage x Current x Time"

    result = {
        "calculation_method": calculation_method,
        "point_a_index": point_a_index,
        "point_b_index": point_b_index,
        "start_index": start_index,
        "end_index": end_index,
        "input_point_a_voltage": point_a_voltage,
        "input_point_b_voltage": point_b_voltage,
        "matched_point_a_voltage": matched_point_a_voltage,
        "matched_point_b_voltage": matched_point_b_voltage,
        "start_voltage": start_voltage,
        "end_voltage": end_voltage,
        "start_time": section["TIME_DISPLAY"].iloc[0] if "TIME_DISPLAY" in section.columns else "N/A",
        "end_time": section["TIME_DISPLAY"].iloc[-1] if "TIME_DISPLAY" in section.columns else "N/A",
        "total_time_sec": total_time_sec,
        "total_time_hr": total_time_hr,
        "assumed_current_a": assumed_current_a,
        "capacity_consumed_ah": capacity_consumed_ah,
        "energy_consumed_wh": energy_consumed_wh,
    }

    return result

# ============================================================
# MAIN
# ============================================================

def main():
    df, voltage_col = load_excel_data(FILE, SHEET_NAME)

    min_voltage = df["VOLTAGE_USED_FOR_CALCULATION"].min()
    max_voltage = df["VOLTAGE_USED_FOR_CALCULATION"].max()

    print()
    print("Loaded discharge data")
    print("---------------------")
    print(f"File used                  : {FILE}")
    print(f"Voltage column used        : {voltage_col}")
    print(f"Voltage offset applied     : {VOLTAGE_OFFSET_V:.3f} V")
    print(f"Maximum voltage in sheet   : {max_voltage:.3f} V")
    print(f"Minimum voltage in sheet   : {min_voltage:.3f} V")

    if POINT_A_VOLTAGE > max_voltage or POINT_A_VOLTAGE < min_voltage:
        print()
        print(f"Warning: Point A voltage {POINT_A_VOLTAGE:.3f} V is outside sheet range.")

    if POINT_B_VOLTAGE > max_voltage or POINT_B_VOLTAGE < min_voltage:
        print()
        print(f"Warning: Point B voltage {POINT_B_VOLTAGE:.3f} V is outside sheet range.")

    result = calculate_energy_between_voltage_points(
        df=df,
        point_a_voltage=POINT_A_VOLTAGE,
        point_b_voltage=POINT_B_VOLTAGE,
        assumed_current_a=ASSUMED_DISCHARGE_CURRENT_A
    )

    print()
    print("Energy consumed between voltage points")
    print("--------------------------------------")
    print(f"Calculation method          : {result['calculation_method']}")
    print(f"Input Point A voltage       : {result['input_point_a_voltage']:.3f} V")
    print(f"Input Point B voltage       : {result['input_point_b_voltage']:.3f} V")
    print(f"Matched Point A voltage     : {result['matched_point_a_voltage']:.3f} V")
    print(f"Matched Point B voltage     : {result['matched_point_b_voltage']:.3f} V")
    print(f"Start voltage used          : {result['start_voltage']:.3f} V")
    print(f"End voltage used            : {result['end_voltage']:.3f} V")
    print(f"Start row index             : {result['start_index']}")
    print(f"End row index               : {result['end_index']}")
    print(f"Start time                  : {result['start_time']}")
    print(f"End time                    : {result['end_time']}")
    print(f"Total time                  : {result['total_time_sec']:.2f} sec")
    print(f"Total time                  : {result['total_time_hr']:.6f} hr")
    print(f"Assumed discharge current   : {result['assumed_current_a']:.3f} A")
    print(f"Capacity consumed           : {result['capacity_consumed_ah']:.6f} Ah")
    print(f"Energy consumed             : {result['energy_consumed_wh']:.6f} Wh")
    print()


if __name__ == "__main__":
    main()
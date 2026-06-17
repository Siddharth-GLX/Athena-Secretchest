import pandas as pd
import numpy as np
from scipy.interpolate import interp1d
import matplotlib.pyplot as plt
#df = pd.read_csv("file.csv")

# 1) Read original Excel (update name if needed)
INPUT_FILE = "Battery-plot digitalized.xlsx"
OUTPUT_FILE = "Battery-Charge-Characteristics-Unified.xlsx"

df = pd.read_excel(INPUT_FILE)

# 2) Extract each curve with its own time column
voltage = df[["Time (Minutes)", "Voltage (V)"]].dropna()
capacity = df[["Time (Minutes).1", "Capacity (mAh)"]].dropna()
current = df[["Time (Minutes).2", "Current(mA)"]].dropna()

# Rename time columns for convenience
voltage.columns = ["Time", "Voltage"]
capacity.columns = ["Time", "Capacity"]
current.columns = ["Time", "Current"]

# 3) Sort and remove duplicate time stamps
voltage = voltage.sort_values("Time").drop_duplicates("Time")
capacity = capacity.sort_values("Time").drop_duplicates("Time")
current = current.sort_values("Time").drop_duplicates("Time")

# 4) Build a unified time axis (here: 0–max_time in 1‑minute steps)
t_min = 0
t_max = max(voltage["Time"].max(),
            capacity["Time"].max(),
            current["Time"].max())

# step size in minutes; change to 0.5 or 0.1 if you want finer resolution
dt = 1.0

time_grid = np.arange(t_min, t_max + dt, dt)

# 5) Create interpolation functions for each parameter
f_v = interp1d(voltage["Time"], voltage["Voltage"],
               kind="linear", fill_value="extrapolate")
f_c = interp1d(capacity["Time"], capacity["Capacity"],
               kind="linear", fill_value="extrapolate")
f_i = interp1d(current["Time"], current["Current"],
               kind="linear", fill_value="extrapolate")

# 6) Evaluate on the common time grid
V = f_v(time_grid)
C = f_c(time_grid)
I = f_i(time_grid)

# 7) Put everything into one DataFrame
out = pd.DataFrame({
    "Time (Minutes)": np.round(time_grid, 3),
    "Voltage (V)":    np.round(V, 4),
    "Capacity (mAh)": np.round(C, 2),
    "Current (mA)":   np.round(I, 2),
})

# =========================
# 8S8P PACK CONVERSION
# =========================

SERIES_CELLS = 8
PARALLEL_CELLS = 8

# Convert cell-level values to 8S8P pack-level values
out["Pack Voltage 8S (V)"] = out["Voltage (V)"] * SERIES_CELLS
out["Pack Capacity 8P (mAh)"] = out["Capacity (mAh)"] * PARALLEL_CELLS
out["Pack Capacity 8P (Ah)"] = out["Pack Capacity 8P (mAh)"] / 1000.0
out["Pack Current 8P (mA)"] = out["Current (mA)"] * PARALLEL_CELLS
out["Pack Current 8P (A)"] = out["Pack Current 8P (mA)"] / 1000.0


# 8) Write result to Excel
out.to_excel(OUTPUT_FILE, index=False)

print(f"Saved unified data to: {OUTPUT_FILE}")
print(f"Rows: {len(out)}, Time range: {time_grid[0]}–{time_grid[-1]} min, step {dt} min")

# 9) Convert capacity from mAh to Ah
out["Capacity (Ah)"] = out["Capacity (mAh)"] / 1000.0

# =========================
# PLOT: 8S8P DISCHARGE CURVE
# =========================

plt.figure(figsize=(10, 6))

plt.plot(
    out["Pack Capacity 8P (Ah)"],
    out["Pack Voltage 8S (V)"],
    linewidth=2,
    label="8S8P Discharge Curve"
)

plt.scatter(
    out["Pack Capacity 8P (Ah)"],
    out["Pack Voltage 8S (V)"],
    s=12
)

plt.xlabel("Pack Capacity 8P (Ah)")
plt.ylabel("Pack Voltage 8S (V)")
plt.title("8S8P Battery Pack Discharge Curve: Voltage vs Capacity")
plt.grid(True, alpha=0.3)
plt.legend()
plt.tight_layout()

plt.savefig("Battery_8S8P_Discharge_Curve_vs_Ah.png", dpi=300)
plt.show()

# 10) Plot discharge curve: Voltage vs Capacity in Ah
plt.figure(figsize=(10, 6))

plt.plot(
    out["Capacity (Ah)"],
    out["Voltage (V)"],
    linewidth=2,
    label="Discharge Curve"
)

plt.scatter(
    out["Capacity (Ah)"],
    out["Voltage (V)"],
    s=12
)

plt.xlabel("Capacity (Ah)")
plt.ylabel("Voltage (V)")
plt.title("Battery Discharge Curve: Voltage vs Capacity")
plt.grid(True, alpha=0.3)
plt.legend()
plt.tight_layout()

plt.savefig("Battery_Discharge_Curve_vs_Ah.png", dpi=300)
plt.show()
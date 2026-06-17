import re
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d

BASE = Path(__file__).resolve().parent
TXT_FILE = BASE / "EPS_Capacity_Test.txt"
XLS_FILE = BASE / "1C_Discharge_Digitized.xls"

OUT_EXCEL = BASE / "Combined_Discharge_Processed.xlsx"
OUT_PLOT = BASE / "Combined_Discharge_Curve_mAh_Ah.png"

SAMPLE_SEC = 1.0
S, P, DT_MIN = 8, 8, 1.0


def val(block, label):
    m = re.search(rf"{re.escape(label)}\s*\|\s*([-+]?\d*\.?\d+)", block)
    return float(m.group(1)) if m else None


def read_txt(file):
    text = file.read_text(encoding="utf-8", errors="ignore")
    blocks = re.split(r"=+\s*\n\s*Command:\s*BMS_Get_Batt_V_and_I_Readings\s*\n=+", text)

    rows = [
        [val(b, "Total Battery Voltage Reading"), val(b, "Battery Current Reading")]
        for b in blocks
        if "CRC Match" in b and "Total Battery Voltage Reading" in b
    ]

    rows = [r for r in rows if None not in r]
    if not rows:
        raise ValueError("No valid CRC Match battery blocks found in TXT file.")

    df = pd.DataFrame(rows, columns=["TXT Voltage (V)", "TXT Current (A)"])
    df["dt (s)"] = SAMPLE_SEC
    df.loc[0, "dt (s)"] = 0
    df["TXT Capacity (mAh)"] = (df["TXT Current (A)"].abs() * df["dt (s)"] / 3600 * 1000).cumsum()
    df["TXT Capacity (Ah)"] = df["TXT Capacity (mAh)"] / 1000
    return df


def load_xls(file):
    for eng in ["xlrd", "openpyxl"]:
        try:
            return pd.read_excel(file, engine=eng)
        except Exception:
            pass

    try:
        return pd.read_csv(file, sep=None, engine="python")
    except Exception:
        raw = pd.read_excel(file, header=None)
        for i in range(min(20, len(raw))):
            row = " ".join(raw.iloc[i].astype(str).str.lower())
            if "voltage" in row or "capacity" in row or "mah" in row:
                df = raw.iloc[i + 1:].copy()
                df.columns = raw.iloc[i].astype(str).str.strip()
                return df
        return raw


def clean(df):
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    print("\nDetected XLS columns:")
    print(df.columns.tolist())
    return df


def find_col(df, words):
    for c in df.columns:
        name = str(c).lower()
        if all(w in name for w in words):
            return c
    return None


def read_digitized_xls(file):
    df = clean(load_xls(file))

    v_col = find_col(df, ["voltage"])
    c_col = find_col(df, ["capacity"]) or find_col(df, ["mah"])
    time_cols = [c for c in df.columns if "time" in str(c).lower()]

    if v_col and c_col and len(time_cols) >= 2:
        tv, tc = time_cols[0], time_cols[1]

        v = df[[tv, v_col]].apply(pd.to_numeric, errors="coerce").dropna()
        c = df[[tc, c_col]].apply(pd.to_numeric, errors="coerce").dropna()

        v = v.sort_values(tv).drop_duplicates(tv)
        c = c.sort_values(tc).drop_duplicates(tc)

        t = np.arange(0, max(v[tv].max(), c[tc].max()) + DT_MIN, DT_MIN)
        voltage = interp1d(v[tv], v[v_col], fill_value="extrapolate")(t)
        capacity = interp1d(c[tc], c[c_col], fill_value="extrapolate")(t)

    else:
        num = df.apply(pd.to_numeric, errors="coerce").dropna(axis=1, how="all")
        num = num.loc[:, num.notna().sum() > 2]

        if num.shape[1] >= 4:
            tv, v_col, tc, c_col = num.columns[:4]
            v = num[[tv, v_col]].dropna().sort_values(tv).drop_duplicates(tv)
            c = num[[tc, c_col]].dropna().sort_values(tc).drop_duplicates(tc)

            t = np.arange(0, max(v[tv].max(), c[tc].max()) + DT_MIN, DT_MIN)
            voltage = interp1d(v[tv], v[v_col], fill_value="extrapolate")(t)
            capacity = interp1d(c[tc], c[c_col], fill_value="extrapolate")(t)

        elif num.shape[1] >= 2:
            capacity = num.iloc[:, 0].dropna().to_numpy()
            voltage = num.iloc[:len(capacity), 1].dropna().to_numpy()
            n = min(len(capacity), len(voltage))
            capacity, voltage = capacity[:n], voltage[:n]
            t = np.arange(n)

        else:
            raise ValueError("Could not identify voltage/capacity columns in Discharge_Digitized.xls.")

    return pd.DataFrame({
        "Digitized Voltage Cell (V)": voltage,
        "Digitized Capacity Cell (mAh)": capacity,
        "Digitized Pack Voltage 8S (V)": voltage * S,
        "Digitized Pack Capacity 8P (mAh)": capacity * P,
        "Digitized Pack Capacity 8P (Ah)": capacity * P / 1000
    })


txt = read_txt(TXT_FILE)
dig = read_digitized_xls(XLS_FILE)

with pd.ExcelWriter(OUT_EXCEL, engine="openpyxl") as w:
    txt.to_excel(w, sheet_name="EPS_TXT", index=False)
    dig.to_excel(w, sheet_name="Digitized_8S8P", index=False)

fig, ax = plt.subplots(figsize=(11, 6))

ax.plot(txt["TXT Capacity (mAh)"], txt["TXT Voltage (V)"], linewidth=2, label="EPS Capacity Test")
ax.plot(dig["Digitized Pack Capacity 8P (mAh)"], dig["Digitized Pack Voltage 8S (V)"], linewidth=2, label="Digitized 8S8P Curve")

ax.set_xlabel("Capacity (mAh)")
ax.set_ylabel("Voltage (V)")
ax.set_title("Discharge Curve: Voltage vs Capacity")
ax.grid(True, alpha=0.3)
ax.legend()

top = ax.secondary_xaxis("top", functions=(lambda x: x / 1000, lambda x: x * 1000))
top.set_xlabel("Capacity (Ah)")

plt.tight_layout()
plt.savefig(OUT_PLOT, dpi=300)
plt.show()

print(f"\nSaved Excel: {OUT_EXCEL}")
print(f"Saved Plot : {OUT_PLOT}")
print(f"TXT Final Capacity       : {txt['TXT Capacity (Ah)'].iloc[-1]:.3f} Ah")
print(f"Digitized Final Capacity : {dig['Digitized Pack Capacity 8P (Ah)'].iloc[-1]:.3f} Ah")
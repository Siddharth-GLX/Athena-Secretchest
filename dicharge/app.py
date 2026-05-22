import io
import zipfile
from pathlib import Path
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import streamlit as st


# -----------------------------
# Streamlit page configuration
# -----------------------------
st.set_page_config(
    page_title="EPS Charge-Discharge Analysis",
    page_icon="🔋",
    layout="wide",
)

st.title("EPS Charge-Discharge Analysis Dashboard")
st.write(
    "Upload the EPS charge/discharge Excel file, process the discharging sheet, "
    "view plots and pack summary values, and download the processed results."
)


# -----------------------------
# Battery pack configuration
# -----------------------------
SERIES_CELLS = 8
PARALLEL_CELLS = 8

# Panasonic NCR18650B cell data
CELL_NOMINAL_VOLTAGE_V = 3.6
CELL_RATED_CAPACITY_AH = 3.2
CELL_TYP_CAPACITY_AH = 3.35

# Derived pack nominal values
PACK_NOMINAL_VOLTAGE_V = SERIES_CELLS * CELL_NOMINAL_VOLTAGE_V
PACK_RATED_CAPACITY_AH = PARALLEL_CELLS * CELL_RATED_CAPACITY_AH
PACK_TYP_CAPACITY_AH = PARALLEL_CELLS * CELL_TYP_CAPACITY_AH
PACK_RATED_ENERGY_WH = PACK_NOMINAL_VOLTAGE_V * PACK_RATED_CAPACITY_AH
PACK_TYP_ENERGY_WH = PACK_NOMINAL_VOLTAGE_V * PACK_TYP_CAPACITY_AH


# -----------------------------
# Processing functions
# -----------------------------
def load_sheet(uploaded_file, sheet_name: str, header_mode: str = "normal") -> pd.DataFrame:
    """Load and process the selected Excel sheet."""

    if header_mode == "normal":
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    elif header_mode == "no_header":
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        df.columns = ["TIMESTAMP", "TOTAL_BATTERY_VOLTAGE", "DISCHARGE_CURRENT"]
    else:
        raise ValueError(f"Invalid header_mode: {header_mode}")

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    required_cols = ["TIMESTAMP", "TOTAL_BATTERY_VOLTAGE", "DISCHARGE_CURRENT"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(
            f"Missing required columns: {missing_cols}. "
            f"Available columns are: {list(df.columns)}"
        )

    df = df[required_cols].copy()

    df["TIMESTAMP"] = pd.to_numeric(df["TIMESTAMP"], errors="coerce")
    df["TOTAL_BATTERY_VOLTAGE"] = pd.to_numeric(df["TOTAL_BATTERY_VOLTAGE"], errors="coerce")
    df["DISCHARGE_CURRENT"] = pd.to_numeric(df["DISCHARGE_CURRENT"], errors="coerce")

    df = df.dropna(subset=required_cols)
    df = df.sort_values("TIMESTAMP").reset_index(drop=True)

    # Real datetime column for plotting
    df["TIMESTAMP_UTC_DT"] = pd.to_datetime(
        df["TIMESTAMP"], unit="s", utc=True
    ).dt.tz_localize(None)

    # Text column for Excel export
    df["TIMESTAMP_UTC"] = df["TIMESTAMP_UTC_DT"].dt.strftime("%Y-%m-%d %H:%M:%S UTC")

    # Time difference in seconds
    df["dt_sec"] = df["TIMESTAMP"].diff().fillna(0)
    df.loc[df["dt_sec"] < 0, "dt_sec"] = 0

    # Use absolute current so discharge capacity/energy remains positive
    df["CURRENT_ABS_A"] = df["DISCHARGE_CURRENT"].abs()

    # Capacity
    df["d_mAh"] = df["CURRENT_ABS_A"] * df["dt_sec"] / 3600.0 * 1000.0
    df["Capacity_mAh"] = df["d_mAh"].cumsum()
    df["Capacity_Ah"] = df["Capacity_mAh"] / 1000.0

    # Power and energy
    df["Power_W"] = df["TOTAL_BATTERY_VOLTAGE"] * df["CURRENT_ABS_A"]
    df["d_Wh"] = df["Power_W"] * df["dt_sec"] / 3600.0
    df["Energy_Wh"] = df["d_Wh"].cumsum()

    return df


def create_excel_bytes(dis_df: pd.DataFrame) -> bytes:
    """Create Excel file in memory for download and optional local storage."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dis_df.to_excel(writer, sheet_name="Discharging_Processed", index=False)

        summary_df = pd.DataFrame(
            {
                "Parameter": [
                    "Nominal pack voltage",
                    "Rated pack capacity",
                    "Typical pack capacity",
                    "Rated nominal energy",
                    "Typical nominal energy",
                    "Measured cumulative discharge energy",
                ],
                "Value": [
                    PACK_NOMINAL_VOLTAGE_V,
                    PACK_RATED_CAPACITY_AH,
                    PACK_TYP_CAPACITY_AH,
                    PACK_RATED_ENERGY_WH,
                    PACK_TYP_ENERGY_WH,
                    dis_df["Energy_Wh"].iloc[-1],
                ],
                "Unit": ["V", "Ah", "Ah", "Wh", "Wh", "Wh"],
            }
        )
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    output.seek(0)
    return output.getvalue()


def make_single_axis_plot(
    dis_df: pd.DataFrame,
    x_col: str,
    y_col: str,
    xlabel: str,
    ylabel: str,
    title: str,
    is_time_axis: bool = False,
) -> tuple[plt.Figure, bytes]:
    """Create a matplotlib plot and return both figure and PNG bytes."""
    fig, ax = plt.subplots(figsize=(10, 6))

    ax.plot(dis_df[x_col], dis_df[y_col], label="Discharging", linewidth=2)
    ax.scatter(dis_df[x_col], dis_df[y_col], s=12, zorder=3)

    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.grid(True, alpha=0.3)
    ax.legend()

    if is_time_axis:
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d %H:%M:%S"))
        plt.setp(ax.get_xticklabels(), rotation=45, ha="right")

    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=300, bbox_inches="tight")
    buffer.seek(0)

    return fig, buffer.getvalue()


def make_voltage_current_plot(dis_df: pd.DataFrame) -> tuple[plt.Figure, bytes]:
    """Create voltage and current vs time plot with two y-axes."""
    fig, ax1 = plt.subplots(figsize=(12, 6))

    line1 = ax1.plot(
        dis_df["TIMESTAMP_UTC_DT"],
        dis_df["TOTAL_BATTERY_VOLTAGE"],
        label="Voltage",
        linewidth=2,
    )
    ax1.scatter(
        dis_df["TIMESTAMP_UTC_DT"],
        dis_df["TOTAL_BATTERY_VOLTAGE"],
        s=12,
        zorder=3,
    )
    ax1.set_xlabel("Timestamp (UTC)")
    ax1.set_ylabel("Voltage (V)")
    ax1.grid(True, alpha=0.3)
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d %H:%M:%S"))

    ax2 = ax1.twinx()
    line2 = ax2.plot(
        dis_df["TIMESTAMP_UTC_DT"],
        dis_df["DISCHARGE_CURRENT"],
        label="Current",
        linewidth=2,
        linestyle="--",
    )
    ax2.scatter(
        dis_df["TIMESTAMP_UTC_DT"],
        dis_df["DISCHARGE_CURRENT"],
        s=12,
        zorder=3,
    )
    ax2.set_ylabel("Discharge Current (A)")

    lines = line1 + line2
    labels = [line.get_label() for line in lines]
    ax1.legend(lines, labels, loc="best")

    ax1.set_title("Voltage and Current vs Time")
    plt.setp(ax1.get_xticklabels(), rotation=45, ha="right")
    fig.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=300, bbox_inches="tight")
    buffer.seek(0)

    return fig, buffer.getvalue()


def create_all_plots(dis_df: pd.DataFrame) -> dict[str, tuple[plt.Figure, bytes]]:
    """Create all output plots."""
    plots = {}

    plots["voltage_vs_mAh.png"] = make_single_axis_plot(
        dis_df,
        "Capacity_mAh",
        "TOTAL_BATTERY_VOLTAGE",
        "Capacity (mAh)",
        "Voltage (V)",
        "Voltage vs mAh",
    )

    plots["voltage_vs_Ah.png"] = make_single_axis_plot(
        dis_df,
        "Capacity_Ah",
        "TOTAL_BATTERY_VOLTAGE",
        "Capacity (Ah)",
        "Voltage (V)",
        "Voltage vs Ah",
    )

    plots["voltage_vs_Time.png"] = make_single_axis_plot(
        dis_df,
        "TIMESTAMP_UTC_DT",
        "TOTAL_BATTERY_VOLTAGE",
        "Timestamp (UTC)",
        "Voltage (V)",
        "Voltage vs Time",
        is_time_axis=True,
    )

    plots["Current_vs_Time.png"] = make_single_axis_plot(
        dis_df,
        "TIMESTAMP_UTC_DT",
        "DISCHARGE_CURRENT",
        "Timestamp (UTC)",
        "Discharge Current (A)",
        "Current vs Time",
        is_time_axis=True,
    )

    plots["Capacity_vs_Time.png"] = make_single_axis_plot(
        dis_df,
        "TIMESTAMP_UTC_DT",
        "Capacity_Ah",
        "Timestamp (UTC)",
        "Capacity (Ah)",
        "Capacity vs Time",
        is_time_axis=True,
    )

    plots["Voltage_Current_vs_Time.png"] = make_voltage_current_plot(dis_df)

    plots["voltage_vs_Wh.png"] = make_single_axis_plot(
        dis_df,
        "Energy_Wh",
        "TOTAL_BATTERY_VOLTAGE",
        "Energy (Wh)",
        "Voltage (V)",
        "Voltage vs Wh",
    )

    return plots


def create_zip_bytes(excel_bytes: bytes, plots: dict[str, tuple[plt.Figure, bytes]]) -> bytes:
    """Create a ZIP file containing Excel output and all plot PNG files."""
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("EPS-Charge-Discharge-Processed.xlsx", excel_bytes)

        for file_name, (_fig, image_bytes) in plots.items():
            zip_file.writestr(file_name, image_bytes)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def save_outputs_locally(excel_bytes: bytes, plots: dict[str, tuple[plt.Figure, bytes]]) -> Path:
    """Save outputs to a local folder on the PC running Streamlit."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path("streamlit_outputs") / f"run_{timestamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    (output_dir / "EPS-Charge-Discharge-Processed.xlsx").write_bytes(excel_bytes)

    for file_name, (_fig, image_bytes) in plots.items():
        (output_dir / file_name).write_bytes(image_bytes)

    return output_dir


# -----------------------------
# Sidebar controls
# -----------------------------
with st.sidebar:
    st.header("Input Settings")

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        help="Upload your EPS Charge Discharge Test Excel file.",
    )

    sheet_name = st.text_input("Sheet name", value="Discharging")

    header_mode_label = st.selectbox(
        "Header mode",
        options=["Normal header", "No header"],
        index=0,
    )

    header_mode = "normal" if header_mode_label == "Normal header" else "no_header"

    save_to_local_folder = st.checkbox(
        "Also save outputs locally on this PC",
        value=True,
        help="Files will be saved in a streamlit_outputs folder near this app.py file.",
    )

    process_button = st.button("Process File", type="primary")


# -----------------------------
# Main app logic
# -----------------------------
if uploaded_file is None:
    st.info("Upload your Excel file from the sidebar to start.")
    st.stop()

if process_button:
    try:
        with st.spinner("Processing uploaded Excel file..."):
            dis_df = load_sheet(uploaded_file, sheet_name=sheet_name, header_mode=header_mode)
            excel_bytes = create_excel_bytes(dis_df)
            plots = create_all_plots(dis_df)
            zip_bytes = create_zip_bytes(excel_bytes, plots)

            saved_folder = None
            if save_to_local_folder:
                saved_folder = save_outputs_locally(excel_bytes, plots)

        st.success("Processing completed successfully.")

        if saved_folder is not None:
            st.success(f"Outputs saved locally at: {saved_folder.resolve()}")

        # -----------------------------
        # Summary metrics
        # -----------------------------
        st.subheader("Pack Nominal Values")

        col1, col2, col3 = st.columns(3)
        col1.metric("Nominal Pack Voltage", f"{PACK_NOMINAL_VOLTAGE_V:.2f} V")
        col2.metric("Rated Pack Capacity", f"{PACK_RATED_CAPACITY_AH:.2f} Ah")
        col3.metric("Typical Pack Capacity", f"{PACK_TYP_CAPACITY_AH:.2f} Ah")

        col4, col5, col6 = st.columns(3)
        col4.metric("Rated Nominal Energy", f"{PACK_RATED_ENERGY_WH:.2f} Wh")
        col5.metric("Typical Nominal Energy", f"{PACK_TYP_ENERGY_WH:.2f} Wh")
        col6.metric(
            "Measured Discharge Energy",
            f"{dis_df['Energy_Wh'].iloc[-1]:.3f} Wh",
        )

        # -----------------------------
        # Processed table
        # -----------------------------
        st.subheader("Processed Discharging Data")
        st.dataframe(dis_df, use_container_width=True)

        # -----------------------------
        # Downloads
        # -----------------------------
        st.subheader("Download Results")

        dcol1, dcol2 = st.columns(2)

        with dcol1:
            st.download_button(
                label="Download Processed Excel",
                data=excel_bytes,
                file_name="EPS-Charge-Discharge-Processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with dcol2:
            st.download_button(
                label="Download All Outputs as ZIP",
                data=zip_bytes,
                file_name="EPS-Charge-Discharge-Results.zip",
                mime="application/zip",
            )

        # -----------------------------
        # Plots
        # -----------------------------
        st.subheader("Plots")

        for file_name, (fig, image_bytes) in plots.items():
            st.markdown(f"### {file_name}")
            st.pyplot(fig)

            st.download_button(
                label=f"Download {file_name}",
                data=image_bytes,
                file_name=file_name,
                mime="image/png",
                key=f"download_{file_name}",
            )

            plt.close(fig)

    except Exception as error:
        st.error("Processing failed.")
        st.exception(error)

else:
    st.warning("Click **Process File** from the sidebar after uploading the Excel file.")

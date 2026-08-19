import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sales Report Builder", layout="wide")

OUTPUT_COLUMNS = [
    "shipping date",
    "Reference Number",
    "Order date",
    "Consignee City",
    "Consignee Phone",
    "Carrier WayBill",
    "Salesman",
    "Order Value",
    "date",
    "status",
    "Clarify",
    "Notes",
    "Orders (first-time)",
    "Orders (returning)",
]

UAE_OMAN_COUNTRIES = {"united arab emirates", "uae", "oman"}

EXCLUDED_SALES_CHANNEL = "point of sale"
REQUIRED_FULFILLMENT_STATUS = "fulfilled"


def read_any(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)


def excel_serial_to_date(series):
    numeric = pd.to_numeric(series, errors="coerce")
    converted = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    return converted


def parse_day_column(series):
    if pd.api.types.is_numeric_dtype(series):
        return excel_serial_to_date(series)
    parsed = pd.to_datetime(series, errors="coerce")
    if parsed.isna().mean() > 0.5:
        parsed = excel_serial_to_date(series)
    return parsed


def standardize(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    def col(name):
        return df[name] if name in df.columns else pd.Series([None] * len(df), index=df.index)

    out = pd.DataFrame(index=df.index)
    out["Order name"] = col("Order name")
    out["Day"] = parse_day_column(col("Day"))
    out["Shipping country"] = col("Shipping country")
    out["Sales channel"] = col("Sales channel")
    out["Staff member name"] = col("Staff member name")
    out["Order fulfillment status"] = col("Order fulfillment status")
    out["Orders"] = pd.to_numeric(col("Orders"), errors="coerce")
    out["Total sales"] = pd.to_numeric(col("Total sales"), errors="coerce")

    if "Orders (first-time)" in df.columns or "Orders (returning)" in df.columns:
        out["Orders (first-time)"] = pd.to_numeric(col("Orders (first-time)"), errors="coerce").fillna(0)
        out["Orders (returning)"] = pd.to_numeric(col("Orders (returning)"), errors="coerce").fillna(0)
    elif "New or returning customer" in df.columns:
        flag = col("New or returning customer").astype(str).str.strip().str.lower()
        out["Orders (first-time)"] = (flag == "new").astype(int)
        out["Orders (returning)"] = (flag == "returning").astype(int)
    else:
        out["Orders (first-time)"] = 0
        out["Orders (returning)"] = 0

    return out


def apply_filters(df):
    channel = df["Sales channel"].astype(str).str.strip().str.lower()
    status = df["Order fulfillment status"].astype(str).str.strip().str.lower()

    mask = (
        (channel != EXCLUDED_SALES_CHANNEL)
        & (status == REQUIRED_FULFILLMENT_STATUS)
        & (df["Orders"] == 1)
    )
    return df[mask].copy()


def build_report_rows(df):
    staff_raw = df["Staff member name"]
    staff = staff_raw.astype(str).str.strip()
    is_missing = staff_raw.isna() | staff.isin(["", "nan", "None", "<NA>"])
    staff = staff.where(~is_missing, "created by order")

    out = pd.DataFrame()
    out["shipping date"] = pd.NA
    out["Reference Number"] = df["Order name"]
    out["Order date"] = df["Day"]
    out["Consignee City"] = df["Shipping country"]
    out["Consignee Phone"] = pd.NA
    out["Carrier WayBill"] = pd.NA
    out["Salesman"] = staff
    out["Order Value"] = df["Total sales"]
    out["date"] = pd.NA
    out["status"] = pd.NA
    out["Clarify"] = pd.NA
    out["Notes"] = pd.NA
    out["Orders (first-time)"] = df["Orders (first-time)"]
    out["Orders (returning)"] = df["Orders (returning)"]
    return out[OUTPUT_COLUMNS]


def split_by_region(df):
    country_norm = df["Consignee City"].astype(str).str.strip().str.lower()
    uae_oman = df[country_norm.isin(UAE_OMAN_COUNTRIES)].copy()
    rest_of_gulf = df[~country_norm.isin(UAE_OMAN_COUNTRIES) & country_norm.ne("nan") & country_norm.ne("")].copy()
    return uae_oman, rest_of_gulf


def to_excel_bytes(uae_oman, rest_of_gulf):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        uae_oman.to_excel(writer, sheet_name="UAE & Oman", index=False)
        rest_of_gulf.to_excel(writer, sheet_name="Rest of Gulf", index=False)
    buffer.seek(0)
    return buffer


st.title("Sales Report Builder")
st.write("Upload one or more raw sales export files (CSV or XLSX). The app will clean, filter, "
         "and merge them into a single report split into two sheets: UAE & Oman, and Rest of Gulf.")

uploaded_files = st.file_uploader(
    "Upload files", type=["csv", "xlsx"], accept_multiple_files=True
)

if uploaded_files:
    standardized_frames = []
    for f in uploaded_files:
        raw = read_any(f)
        standardized_frames.append(standardize(raw))

    combined = pd.concat(standardized_frames, ignore_index=True)
    filtered = apply_filters(combined)
    report = build_report_rows(filtered)
    uae_oman, rest_of_gulf = split_by_region(report)

    st.subheader("UAE & Oman")
    st.dataframe(uae_oman, use_container_width=True)
    st.write(f"Rows: {len(uae_oman)}")

    st.subheader("Rest of Gulf")
    st.dataframe(rest_of_gulf, use_container_width=True)
    st.write(f"Rows: {len(rest_of_gulf)}")

    excel_bytes = to_excel_bytes(uae_oman, rest_of_gulf)
    st.download_button(
        label="Download report (xlsx)",
        data=excel_bytes,
        file_name="sales_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload at least one file to build the report.")

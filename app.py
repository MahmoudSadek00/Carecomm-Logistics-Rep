import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sales Report Builder", layout="wide")

OUTPUT_COLUMNS = [
    "Shipping Date",
    "Reference Number",
    "Order Date",
    "Consignee City",
    "Consignee Phone",
    "Carrier WayBill",
    "Salesman",
    "Order Value",
    "Date",
    "Status",
    "Clarify",
    "Notes",
    "New Customer Orders",
    "Returning Customer Orders",
]

UAE_OMAN_CODES = {"uae", "om"}

COUNTRY_CODE_MAP = {
    "united arab emirates": "UAE",
    "uae": "UAE",
    "oman": "OM",
    "om": "OM",
    "saudi arabia": "SA",
    "sa": "SA",
    "kuwait": "KW",
    "kw": "KW",
    "qatar": "QA",
    "qa": "QA",
}

EXCLUDED_SALES_CHANNEL = "point of sale"
REQUIRED_FULFILLMENT_STATUS = "fulfilled"


def to_country_code(series):
    normalized = series.astype(str).str.strip().str.lower()
    return normalized.map(COUNTRY_CODE_MAP).fillna(series)


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
    staff = staff.str.title()
    staff = staff.where(~is_missing, "Created by customer")

    order_date = df["Day"]
    if pd.api.types.is_datetime64_any_dtype(order_date):
        order_date = order_date.dt.date

    out = pd.DataFrame()
    out["Shipping Date"] = pd.NA
    out["Reference Number"] = df["Order name"]
    out["Order Date"] = order_date
    out["Consignee City"] = to_country_code(df["Shipping country"])
    out["Consignee Phone"] = pd.NA
    out["Carrier WayBill"] = pd.NA
    out["Salesman"] = staff
    out["Order Value"] = df["Total sales"]
    out["Date"] = pd.NA
    out["Status"] = pd.NA
    out["Clarify"] = pd.NA
    out["Notes"] = pd.NA
    out["New Customer Orders"] = df["Orders (first-time)"]
    out["Returning Customer Orders"] = df["Orders (returning)"]
    return out[OUTPUT_COLUMNS]


def split_by_region(df):
    country_norm = df["Consignee City"].astype(str).str.strip().str.lower()
    uae_oman = df[country_norm.isin(UAE_OMAN_CODES)].copy()
    rest_of_gulf = df[~country_norm.isin(UAE_OMAN_CODES) & country_norm.ne("nan") & country_norm.ne("")].copy()
    return uae_oman, rest_of_gulf


def style_header(worksheet):
    for cell in worksheet[1]:
        cell.font = cell.font.copy(bold=True)


def autofit_columns(worksheet, dataframe, min_width=8, max_width=45, padding=2):
    for i, column in enumerate(dataframe.columns, start=1):
        header_len = len(str(column))
        if len(dataframe) > 0:
            lengths = dataframe[column].apply(lambda v: 0 if pd.isna(v) else len(str(v)))
            value_len = lengths.max()
        else:
            value_len = 0
        width = max(header_len, value_len) + padding
        width = max(min_width, min(max_width, width))
        letter = worksheet.cell(row=1, column=i).column_letter
        worksheet.column_dimensions[letter].width = width


def to_excel_bytes(uae_oman, rest_of_gulf):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        uae_oman.to_excel(writer, sheet_name="UAE & Oman", index=False)
        rest_of_gulf.to_excel(writer, sheet_name="Rest of Gulf", index=False)
        for sheet_name, frame in (("UAE & Oman", uae_oman), ("Rest of Gulf", rest_of_gulf)):
            ws = writer.sheets[sheet_name]
            style_header(ws)
            autofit_columns(ws, frame)
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

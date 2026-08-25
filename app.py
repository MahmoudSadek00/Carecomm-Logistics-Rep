import datetime as dt

import streamlit as st

from engine import (
    TARGET_SHEETS, FIELD_LABELS, SHOPIFY_DEFAULTS, SHIPPING_DEFAULTS,
    read_any, default_mapping, merge_sources, workbook_to_bytes,
)

st.set_page_config(page_title="Orders Consolidation Tool", layout="wide")
st.title("Shopify + Shipping Company Consolidation Tool")
st.caption(
    "Upload this batch's Shopify report and the shipping company's export for ONE country "
    "group, confirm the column mapping, and download an Excel file ready to paste straight "
    "into that raw Google Sheet. Shipping date and delivery status (Analysis) are NOT "
    "included here -- those aren't known yet at this stage and still get filled in later, "
    "same as today."
)

# ---------------------------------------------------------------------------
# 1. Which country group -- upload only what you need, nothing else required
# ---------------------------------------------------------------------------
st.header("1. Country group")
target_key = st.selectbox(
    "Which raw sheet is this batch for?",
    options=list(TARGET_SHEETS.keys()),
    format_func=lambda k: TARGET_SHEETS[k]['label'],
)
fields = TARGET_SHEETS[target_key]['fields']
join_base = TARGET_SHEETS[target_key].get('join_base', 'shopify')
has_baked_shipping = any(f == 'shipping_fee' for h, f, s in fields)
st.caption("Output columns: " + ', '.join(h for h, f, s in fields))
if join_base == 'shipping':
    st.info(
        "This sheet is joined shipping-first: only orders that already have a row in the shipping "
        "company's export show up in the download. A Shopify order with no shipping row yet is left "
        "out (it'll appear once shipping company export catches up), rather than shown as Pending."
    )

# ---------------------------------------------------------------------------
# 2. Shopify file + column mapping + filters
# ---------------------------------------------------------------------------
st.header("2. Shopify report")
shopify_file = st.file_uploader("Shopify report (csv or xlsx)", type=['csv', 'xlsx', 'xls'], key='shopify')

shopify_df = None
shopify_map = {}
canceled_col = None
fulfillment_col = None
allowed_statuses = None
exclude_canceled = True
shopify_date_convention = 'month_first'
default_city = 'UAE'
include_shipping_fee = False

if shopify_file is not None:
    shopify_df = read_any(shopify_file)
    st.write(f"{len(shopify_df)} rows loaded.")
    guessed = default_mapping(shopify_df.columns.tolist(), SHOPIFY_DEFAULTS)

    # 'ref_number' is always needed on the Shopify side to group/aggregate an
    # order's rows, even on a sheet (like Iraq) whose OUTPUT Reference column
    # is sourced from the shipping file instead.
    shopify_fields = ['ref_number'] + [f for h, f, s in fields if s == 'shopify' and f != 'ref_number']
    cols = st.columns(3)
    for i, field in enumerate(shopify_fields):
        with cols[i % 3]:
            options = ['(none)'] + shopify_df.columns.tolist()
            default = guessed.get(field)
            idx = options.index(default) if default in options else 0
            choice = st.selectbox(FIELD_LABELS[field], options, index=idx, key=f'shopify_{target_key}_{field}')
            shopify_map[field] = None if choice == '(none)' else choice

    c1, c2 = st.columns(2)
    with c1:
        shopify_date_convention = st.radio(
            "'Order date' format in this file", ['month_first (e.g. Shopify default, MM/DD/YYYY)', 'day_first (DD/MM/YYYY)'],
            horizontal=True, key='shopify_date_conv',
        )
        shopify_date_convention = 'month_first' if shopify_date_convention.startswith('month_first') else 'day_first'
    with c2:
        if TARGET_SHEETS[target_key]['country_choices']:
            default_city = st.selectbox(
                "Default city for orders with no shipping address (POS / Draft / Mobile)",
                options=list(TARGET_SHEETS[target_key]['country_choices'].keys()),
                format_func=lambda k: f"{k} ({TARGET_SHEETS[target_key]['country_choices'][k]})",
            )

    shipping_fee_box_title = (
        "Shipping column (Total sales - Net sales, summed per order)" if has_baked_shipping
        else "Optional: add a Shipping column (Total sales - Net sales, summed per order)"
    )
    with st.expander(shipping_fee_box_title, expanded=has_baked_shipping):
        st.caption(
            "Every 'item added later' row seen in this format has Total sales == Net sales, so this "
            "stays correctly anchored to the original order's shipping charge even when an order has "
            "more than one row. See the README for the worked example."
        )
        options = ['(none)'] + shopify_df.columns.tolist()
        net_guess = ['Net sales'] if 'Net sales' in shopify_df.columns else []
        net_choice = st.selectbox(
            FIELD_LABELS['net_sales'], options, index=options.index(net_guess[0]) if net_guess else 0,
        )
        shopify_map['net_sales'] = None if net_choice == '(none)' else net_choice
        if has_baked_shipping:
            include_shipping_fee = True
            if not shopify_map['net_sales']:
                st.warning("This sheet's Shipping column needs a Net sales mapping, or it will be left blank.")
        else:
            include_shipping_fee = st.checkbox(
                "Add the Shipping column to the download", value=False, disabled=not shopify_map['net_sales'],
            )

    with st.expander("Filters (optional columns on the Shopify file)"):
        cancel_options = ['(none)'] + shopify_df.columns.tolist()
        canceled_guess = ['Is canceled order'] if 'Is canceled order' in shopify_df.columns else []
        canceled_col_choice = st.selectbox(
            "Cancelled-order column", cancel_options,
            index=cancel_options.index(canceled_guess[0]) if canceled_guess else 0,
        )
        canceled_col = None if canceled_col_choice == '(none)' else canceled_col_choice
        exclude_canceled = st.checkbox("Exclude cancelled orders", value=True, disabled=canceled_col is None)

        fulfill_options = ['(none)'] + shopify_df.columns.tolist()
        fulfill_guess = ['Order fulfillment status'] if 'Order fulfillment status' in shopify_df.columns else []
        fulfillment_col_choice = st.selectbox(
            "Fulfillment-status column", fulfill_options,
            index=fulfill_options.index(fulfill_guess[0]) if fulfill_guess else 0,
        )
        fulfillment_col = None if fulfillment_col_choice == '(none)' else fulfillment_col_choice
        if fulfillment_col:
            statuses_present = sorted(shopify_df[fulfillment_col].dropna().unique().tolist())
            allowed_statuses = st.multiselect(
                "Include only these statuses (leave all checked to include everything)",
                options=statuses_present, default=statuses_present,
            )

# ---------------------------------------------------------------------------
# 3. Shipping company file + column mapping
# ---------------------------------------------------------------------------
st.header("3. Shipping company export")
shipping_file = st.file_uploader("Shipping company export (csv or xlsx)", type=['csv', 'xlsx', 'xls'], key='shipping')

shipping_df = None
shipping_map = {}
if shipping_file is not None:
    shipping_df = read_any(shipping_file)
    st.write(f"{len(shipping_df)} rows loaded.")
    guessed = default_mapping(shipping_df.columns.tolist(), SHIPPING_DEFAULTS)

    # 'ref_number' is always prepended explicitly -- exclude it from the
    # list-comprehension pass too, or a sheet whose OUTPUT ref_number field
    # is itself sourced from the shipping file (Iraq: 'ReceiptNumber' comes
    # from shipping, not Shopify) ends up with it twice, producing two
    # selectboxes with the same widget key and crashing with
    # StreamlitDuplicateElementKey. Mirrors the same fix already applied to
    # shopify_fields above.
    shipping_fields = ['ref_number'] + [f for h, f, s in fields if s == 'shipping' and f != 'ref_number']
    cols = st.columns(3)
    for i, field in enumerate(shipping_fields):
        with cols[i % 3]:
            options = ['(none)'] + shipping_df.columns.tolist()
            default = guessed.get(field)
            idx = options.index(default) if default in options else 0
            choice = st.selectbox(FIELD_LABELS[field], options, index=idx, key=f'shipping_{target_key}_{field}')
            shipping_map[field] = None if choice == '(none)' else choice

    if shipping_map.get('ref_number') is None:
        st.warning("Pick the Reference/Order Number column on the shipping file -- otherwise nothing can be matched to Shopify.")

# ---------------------------------------------------------------------------
# 4. Merge + preview + download
# ---------------------------------------------------------------------------
st.header("4. Merge & download")

ready = (
    shopify_df is not None and shipping_df is not None
    and shopify_map.get('ref_number') and shipping_map.get('ref_number')
)
if not ready:
    st.info("Upload both files and confirm the Reference Number column on each to continue.")

if st.button("Merge files", disabled=not ready, type="primary"):
    with st.spinner("Matching orders..."):
        merged, warnings, stats = merge_sources(
            target_key,
            shopify_df, shopify_map,
            shipping_df, shipping_map,
            shopify_date_convention=shopify_date_convention,
            default_city_for_blank=default_city,
            exclude_canceled=exclude_canceled,
            canceled_col=canceled_col,
            allowed_fulfillment_statuses=allowed_statuses,
            fulfillment_col=fulfillment_col,
            include_shipping_fee=include_shipping_fee,
        )

    if join_base == 'shopify':
        st.success(
            f"{stats['output_rows']} orders in this batch ({stats['matched']} matched a shipping-company "
            f"row, {stats['output_rows'] - stats['matched']} did not yet). "
            f"{stats['canceled_excluded']} cancelled row(s) excluded, {stats['fulfillment_excluded']} excluded "
            f"by fulfillment-status filter, {stats['orders_with_multiple_rows']} order(s) had their value "
            f"summed across more than one row."
        )
    else:
        st.success(
            f"{stats['output_rows']} shipping-company row(s) in this batch ({stats['matched']} matched a "
            f"Shopify order, {stats['output_rows'] - stats['matched']} did not). "
            f"{stats['orders_total']} Shopify orders were considered in total this run "
            f"({stats['canceled_excluded']} cancelled excluded, {stats['orders_with_multiple_rows']} had "
            f"their value summed across more than one row)."
        )
    for w in warnings:
        st.warning(w)

    review = merged[~merged['_matched']]
    if len(review):
        with st.expander(f"Orders not yet matched to a shipping row ({len(review)})"):
            st.dataframe(review, use_container_width=True)

    st.subheader("Preview -- exactly what will be pasted into the sheet")
    preview_fields = list(fields)
    if include_shipping_fee and not has_baked_shipping:
        preview_fields = preview_fields + [('Shipping', 'shipping_fee', 'computed')]
    preview_cols = [h for h, f, s in preview_fields]
    st.dataframe(merged[[c for c in preview_cols if c in merged.columns]], use_container_width=True)

    xlsx_bytes = workbook_to_bytes(merged, target_key, include_shipping_fee=include_shipping_fee)
    today = dt.date.today().strftime('%Y%m%d')
    st.download_button(
        "Download Excel (ready to paste)",
        data=xlsx_bytes,
        file_name=f"{TARGET_SHEETS[target_key]['label'].split(' (')[0].replace(' ', '_')}_batch_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

# Orders Consolidation Tool

Merges a Shopify report + a shipping company export into one Excel file for
ONE country group (UAE & Oman, Gulf, or Iraq), formatted with the exact
header row of that raw Google Sheet, so the ops team can select-all/copy and
paste it in directly instead of retyping order details by hand.

Built from real sample files across all three country groups, not guessed.

**Shipping date and delivery status (Analysis) are intentionally left
blank.** Neither is known yet at the point an order is created and a
shipping label is generated -- they only exist once the order has actually
moved. The tool still writes those columns (blank) so the download matches
the raw sheet's exact column positions and can be pasted straight in --
filling them in stays part of the existing process (manual, or the separate
sync script), not this tool.

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

## How it works

1. Pick the country group -- decides the output columns. Each one has its
   own real-world quirks:
   - **UAE & Oman / Gulf**: waybill column is called "Carrier WayBill" for
     UAE & Oman but "AWB" for Gulf. Both use the "Golden Collection" carrier
     format on the shipping side (Reference, Carrier Waybill, Consignee
     Contact, Consignee City, ...). Shipping City comes from Shopify's own
     `Shipping country` field, normalized to the raw sheet's country codes.
     **Joined Shopify-first**: every Shopify order shows up, shipping data
     filled in where a match exists, left blank ("Pending") otherwise.
     Output includes the FULL raw sheet layout, confirmed against
     `sheet_templates.xlsx`'s own header row -- `shipping date` / `date` /
     the unnamed column / `Notes` (UAE & Oman only) / `Analysis` come out as
     blank placeholder columns, and `Shipping` is baked in right after
     `Order Value`, same as Iraq. **Gulf's real sheet has no `Notes` column
     but has TWO unnamed blank columns instead of one** -- that's the one
     structural difference from UAE & Oman; everything else (including
     which fields come from Shopify vs. the shipping file) is identical.
   - **Iraq**: uses a completely different carrier export (Order #,
     Recipient Name, Recipient Phone, City, Subtotal, Total, ...) with no
     waybill/AWB field at all, but it DOES carry a consignee name (`Name` on
     the raw sheet), which UAE/Gulf don't have. Iraq's Consignee City comes
     from the **shipping file's own City column** (a real city name, e.g.
     "anbar"), not from Shopify -- unlike UAE/Gulf where it's a country code.
     The output Reference column drops the leading "#" to match the raw
     sheet's own convention (`IQ-1059`, not `#IQ-1059`) -- matching itself
     ignores "#" on either side regardless. Output includes the FULL raw
     sheet layout (`date ship`, `status`, `Notes`, `Analysis` included as
     blank placeholder columns, plus a baked-in `Shipping` column right after
     `Value`) so it pastes in with the exact same column positions.
     **Joined shipping-first** (per Mahmoud, Aug 2026): only orders that
     already have a row in the shipping company's export are included --
     a Shopify order with no shipping row yet is left OUT of the download
     entirely (not shown as Pending), since for Iraq an order isn't logged
     in the tracking sheet until it's actually with the shipping company.
     This is the opposite direction from UAE/Gulf -- see "Join direction"
     below.
2. Upload the Shopify report. Column mapping is pre-filled from the real
   "Monthly POS Report" export format (Day, Sales channel, Shipping country,
   Order name, Staff member name, Total sales, Net sales, Orders
   (first-time), Orders (returning), ...).
3. Upload the shipping company's export and confirm its mapping.
4. Optional filters, all visible with counts, nothing silently dropped:
   cancelled-order exclusion, fulfillment-status inclusion list, and an
   optional **Shipping** column (Total sales minus Net sales, see below).
5. Click **Merge files**. Matching is by Reference/Order Number, ignoring
   stray spaces, invisible characters, and a leading "#" either way (the
   same hidden-character problem that caused the #116339 duplicate-row bug
   won't silently break a match here either).
6. Anything that doesn't find a shipping match yet is flagged, not dropped --
   review it in the expander before downloading.
7. Download the Excel and paste it into the raw sheet.

## Order Value: why it's SUMMED across an order's rows (Aug 2026 fix)

The Monthly POS Report has ONE row per order, PLUS one extra row for every
LATER edit against that same Order name -- either a return (a row that's the
exact negative of the original) or an item added afterward (a row with its
own positive Total sales). An earlier version of this tool just took the
first row and dropped every `Orders = 0` row, assuming they were all harmless
return-adjustment noise -- **that was wrong**: it silently dropped the
revenue from a later item-addition along with the harmless return rows.

The fix: **sum every row sharing that Order name first** (`Total sales` and
`Net sales` separately), THEN derive the two output columns from those
sums. Confirmed against the real Iraq export -- every "item added" row found
had `Total sales == Net sales` (no extra shipping is charged just for adding
an item to an already-created order), so an order's real shipping fee always
stays anchored to the original row, and summing Total/Net separately then
subtracting gives the identical answer as summing the per-row difference
(basic linearity -- there's no discrepancy to reconcile). Every other field
(date, city, salesman, new/returning-customer) is taken from the **earliest
row** in the group -- a later edit doesn't change who ordered it or when.

**Value vs Shipping (Aug 2026 fix):** `Total sales` INCLUDES shipping
(subtotal + taxes + fees + shipping + reversals, per Shopify's own
definition). An earlier version of this tool put `Total sales` straight into
the Value column -- that double-counts once Shipping is broken out as its
own column. Value is now **Net sales** (the goods-only amount) and Shipping
is `Total sales - Net sales`, so **Value + Shipping = Total sales**, with
nothing counted twice.

Worked example from the real data (order #IQ-1050):

| Day | Total sales | Net sales |
|---|---|---|
| Mar 12 (original) | 107,000 | 103,000 |
| Mar 16 (item added) | 25,750 | 25,750 |

Total sales summed = 132,750. Net sales summed = 128,750. **Value = 128,750**
(Net), **Shipping = 132,750 − 128,750 = 4,000** -- same 4,000 as the original
row alone, since the added-item row contributes 0 to the difference either
way.

## Join direction: which file is the "base"

UAE & Oman and Gulf are joined **Shopify-first**: the download has one row
per Shopify order, with shipping data (phone, waybill) filled in wherever a
match is found and left blank otherwise -- so an order that hasn't shipped
yet still shows up (as Pending, to be filled in once it ships).

Iraq is joined **shipping-first**: the download has one row per shipping
company row, with Shopify data (date, value, new/returning-customer) filled
in wherever a match is found. A Shopify order that hasn't reached the
shipping company yet is NOT included at all -- it'll appear once the
shipping export catches up to it. This matches how the Iraq raw sheet is
actually used: an order isn't logged there until it's actually with the
shipping company.

## Why the date comes out fixed, not typed

Order date is written as a **real Excel date cell**, not typed text like
"25/1/2026". That sidesteps the whole day-first/month-first ambiguity that
caused the Iraq and Gulf date bugs earlier -- a real date cell carries its
value unambiguously when pasted into Google Sheets.

## What still needs a real sample to finish

- **UAE & Oman's shipping export** -- only Gulf's "Golden Collection" format
  was sampled. If the column names differ, just re-point the mapping
  dropdowns (no code change needed).
- If a cancelled/fulfillment-status filter, or the multi-row Order Value
  summing, ever produces a number that looks wrong on a real batch, send me
  that file and I'll dig into the specific rows.

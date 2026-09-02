from datetime import date, timedelta
import json
import traceback
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

PaymentFunnelOptions = {"No": ["10xTechies", "AI", "AI BootcampPaid", "AI TV", "AI UAE", "AI Unknown", "AI USA", "DivineLane", "DRF",
                        "Excel", "PBI", "PU", "Python", "SMAI", "SQL","AI US"], 
                        "Yes": ["AI Unknown", "AI Bootcamp", "AI Exotic",  "PU Exotic", "SMAI Exotic"]}

SPREADSHEET_ID = "1v0UI5B4rkWJm3N8cbqnRCa4olvwV6-h-YC2mafNYnjU"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

KEY_Payment_Slug = "Payment Slug"
KEY_ID = "ID"


# ── Google Sheets helpers (ported from UpdateSlugs.ipynb) ───────────────────

@st.cache_resource(show_spinner=False)
def get_gspread_client(creds_json_bytes: bytes):
    creds_dict = json.loads(creds_json_bytes)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def copy_data_validation(worksheet, template_row, start_row, end_row, col_indices):
    """Copy data validation (dropdown rules) from `template_row` down to the
    newly appended rows [start_row, end_row], for the given 0-indexed columns."""
    sheet_id = worksheet.id
    requests = [
        {
            "copyPaste": {
                "source": {
                    "sheetId": sheet_id,
                    "startRowIndex": template_row - 1,
                    "endRowIndex": template_row,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1,
                },
                "destination": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row - 1,
                    "endRowIndex": end_row,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1,
                },
                "pasteType": "PASTE_DATA_VALIDATION",
            }
        }
        for col_idx in col_indices
    ]
    worksheet.spreadsheet.batch_update({"requests": requests})


def get_ws_records(spreadsheet, worksheet_name):
    worksheet = spreadsheet.worksheet(worksheet_name)
    records = worksheet.get_all_records()
    return records, worksheet


def build_value_to_row_map(records):
    """Map every existing Payment Slug / ID value to its 1-indexed sheet row
    (row 1 is the header, so record i sits at sheet row i+2)."""
    value_to_row = {}
    for i, row in enumerate(records):
        sheet_row = i + 2
        ps = str(row.get(KEY_Payment_Slug, "")).strip()
        idv = str(row.get(KEY_ID, "")).strip()
        if ps:
            value_to_row[ps] = sheet_row
        if idv:
            value_to_row[idv] = sheet_row
    return value_to_row


def delete_rows(worksheet, row_numbers):
    """Delete the given 1-indexed sheet rows. Must be applied highest-row-first
    so earlier deletions in the same batch don't shift the rows still to delete."""
    if not row_numbers:
        return
    sheet_id = worksheet.id
    requests = [
        {
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": r - 1,
                    "endIndex": r,
                }
            }
        }
        for r in sorted(set(row_numbers), reverse=True)
    ]
    worksheet.spreadsheet.batch_update({"requests": requests})


def update_records(df, worksheet, records):
    df = df.copy()
    for col in ["PaymentFunnel", "isExotic", "Payment Slug", "ID"]:
        df[col] = df[col].fillna("").astype(str)

    if worksheet.row_count == 0 or not worksheet.row_values(1):
        worksheet.append_row(df.columns.tolist(), value_input_option="RAW")

    # The identifying value for a row is whichever column it's routed to:
    # ID for exotic rows, Payment Slug otherwise.
    if df.empty:
        df["_value"] = pd.Series(dtype=str)
    else:
        df["_value"] = [
            (idv.strip() if exotic == "Yes" else slug.strip())
            for exotic, slug, idv in zip(df["isExotic"], df["Payment Slug"], df["ID"])
        ]

    value_to_row = build_value_to_row_map(records)
    rows_to_delete = [value_to_row[v] for v in df["_value"] if v and v in value_to_row]
    updated_count = len(rows_to_delete)

    if rows_to_delete:
        delete_rows(worksheet, rows_to_delete)

    rows = df.drop(columns="_value").values.tolist()

    # Row count *after* deletions, so new rows land in the right place.
    existing_row_count = len(worksheet.get_all_values())
    start_row = existing_row_count + 1

    worksheet.append_rows(rows, value_input_option="USER_ENTERED")

    end_row = start_row + len(rows) - 1
    template_row = existing_row_count if existing_row_count >= 2 else 2

    col_positions = {col: idx for idx, col in enumerate(["PaymentFunnel", "isExotic", "Payment Slug", "ID"])}
    dropdown_cols = [col_positions["PaymentFunnel"], col_positions["isExotic"]]

    if start_row > template_row:
        copy_data_validation(worksheet, template_row=template_row,
                              start_row=start_row, end_row=end_row,
                              col_indices=dropdown_cols)

    return len(rows), updated_count, start_row, end_row

def next_sunday():
    """
    Returns the date of the upcoming Sunday in 'YYYY-MM-DD' format.
    If today is already a Sunday, returns today's date.
    """
    today = date.today()
    # Monday=0 ... Sunday=6
    days_until_sunday = (6 - today.weekday()) % 7
    result = today + timedelta(days=days_until_sunday)
    return result.strftime("%Y-%m-%d")

# ── App ───────────────────────────────────────────────────────────────────

st.set_page_config("Update Payment Slugs", layout="wide")
st.title("Update Payment Slugs")

with st.sidebar:
    st.subheader("Google Sheet connection")
    creds_file = st.file_uploader("Upload service account credentials.json", type="json")
    worksheet_date = st.date_input("Worksheet date", value=next_sunday())
    worksheet_name = worksheet_date.strftime("%Y-%m-%d")
    st.caption(f"Target worksheet: `{worksheet_name}`")

is_exotic_batch = st.checkbox(
    "This batch is exotic — uploaded values are IDs, not Payment Slugs",
    key="is_exotic_batch",
)

uploaded_file = st.file_uploader(
    "Upload CSV with Payment Slugs" + (" (or IDs, since this batch is exotic)" if is_exotic_batch else ""),
    type="csv",
)

lookup_file = None
if is_exotic_batch:
    lookup_file = st.file_uploader(
        "Upload lookup CSV (columns: 'Payment Slug', 'ID') to resolve the Payment Slug for each ID",
        type="csv",
        key="lookup_csv",
    )

if uploaded_file is not None:
    raw_df = pd.read_csv(uploaded_file)
    raw_col = raw_df.columns[0]
    raw_values = raw_df[raw_col].astype(str).str.strip().tolist()

    slug_lookup = {}
    unmatched = []

    if is_exotic_batch:
        if lookup_file is None:
            st.warning("Upload the lookup CSV (Payment Slug + ID) to resolve Payment Slugs for these IDs.")
            st.stop()

        lookup_df = pd.read_csv(lookup_file)
        lookup_df.rename(columns={"slug":"Payment Slug", "id":"ID"  }, inplace=True)
        lookup_df.columns = [c.strip() for c in lookup_df.columns]
        missing_cols = {"Payment Slug", "ID"} - set(lookup_df.columns)
        if missing_cols:
            st.error(f"Lookup CSV is missing column(s): {', '.join(sorted(missing_cols))}")
            st.stop()

        lookup_df["ID"] = lookup_df["ID"].astype(str).str.strip()
        lookup_df["Payment Slug"] = lookup_df["Payment Slug"].astype(str).str.strip()
        slug_lookup = dict(zip(lookup_df["ID"], lookup_df["Payment Slug"]))

        unmatched = [v for v in raw_values if v not in slug_lookup]
        if unmatched:
            preview = ", ".join(unmatched[:10]) + ("…" if len(unmatched) > 10 else "")
            st.warning(f"{len(unmatched)} ID(s) had no match in the lookup file: {preview}")

    values = raw_values
    st.caption(f"Loaded {len(values)} record(s) from column '{raw_col}'. One row per record.")

    cols_spec = [2, 1, 2, 2, 1] if is_exotic_batch else [2, 1, 3, 1]
    headers = st.columns(cols_spec)
    headers[0].markdown("**Payment Funnel**")
    headers[1].markdown("**isExotic?**")
    headers[2].markdown("**Value (from CSV)**")
    if is_exotic_batch:
        headers[3].markdown("**Resolved Payment Slug**")
        headers[4].markdown("**Goes to**")
    else:
        headers[3].markdown("**Goes to**")

    results = []
    for i, value in enumerate(values):
        row_cols = st.columns(cols_spec)

        with row_cols[0]:
            pf = st.selectbox(
                "Payment Funnel", options=PaymentFunnelOptions["Yes" if is_exotic_batch else "No"],
                key=f"pf_{i}", label_visibility="collapsed"
            )
        with row_cols[1]:
            ie = st.selectbox(
                "isExotic?", options=["No", "Yes"],
                index=1 if is_exotic_batch else 0,
                key=f"ie_{i}", label_visibility="collapsed", disabled=True
            )
        with row_cols[2]:
            st.text_input(
                "Value", value=value, disabled=True,
                key=f"val_{i}", label_visibility="collapsed"
            )

        resolved_slug = slug_lookup.get(value, "") if is_exotic_batch else value

        if is_exotic_batch:
            with row_cols[3]:
                st.text_input(
                    "Resolved Payment Slug", value=resolved_slug or "⚠️ no match",
                    disabled=True, key=f"resolved_{i}", label_visibility="collapsed"
                )
            with row_cols[4]:
                st.markdown("`ID`" if ie == "Yes" else "`Payment Slug`")
        else:
            with row_cols[3]:
                st.markdown("`ID`" if ie == "Yes" else "`Payment Slug`")

        # Route: isExotic == Yes -> ID column, isExotic == No -> Payment Slug column.
        # In exotic-batch mode the CSV value IS the ID; the Payment Slug comes from the lookup merge.
        if ie == "Yes":
            id_val = value
            payment_slug = resolved_slug if is_exotic_batch else ""
        else:
            id_val = ""
            payment_slug = value

        results.append({
            "PaymentFunnel": pf,
            "isExotic": ie,
            "Payment Slug": payment_slug,
            "ID": id_val,
        })

    st.divider()

    if st.button("Upload Data?"):
        result_df = pd.DataFrame(results)
        st.dataframe(result_df, use_container_width=True)

        csv_bytes = result_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download as CSV",
            data=csv_bytes,
            file_name="payment_slugs_updated.csv",
            mime="text/csv",
        )

        if creds_file is None:
            st.warning("Upload a service account credentials.json in the sidebar to push these rows to the sheet.")
        else:
            try:
                with st.spinner(f"Connecting to worksheet '{worksheet_name}'..."):
                    client = get_gspread_client(creds_file.getvalue())
                    spreadsheet = client.open_by_key(SPREADSHEET_ID)
                    records, worksheet = get_ws_records(spreadsheet, worksheet_name)
                    total, updated, start_row, end_row = update_records(result_df, worksheet, records)

                new_count = total - updated
                msg = f"Wrote {total} row(s) to worksheet '{worksheet_name}' (rows {start_row}-{end_row})"
                if updated:
                    msg += f" — {updated} replaced existing row(s), {new_count} new."
                else:
                    msg += " — all new."
                st.success(msg)
            except gspread.exceptions.WorksheetNotFound:
                st.error(f"Worksheet '{worksheet_name}' not found in the spreadsheet.")
            except Exception:
                st.error("Failed to update sheet. Full traceback below — please share this if it recurs:")
                st.code(traceback.format_exc())
else:
    st.info("Upload a CSV file to begin.")
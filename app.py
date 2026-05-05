import streamlit as st
import pandas as pd
from io import BytesIO
from processor import load_mapping, process_sales, write_excel

st.set_page_config(
    page_title="StubHub Channel Chare Report Creator",
    page_icon="🎟️",
    layout="centered",
)

st.title("🎟️ StubHub Channel Chare Report Creator")
st.caption("Upload your company mapping file and monthly sales exports to generate split output files.")

# ── Step 1: Mapping file ──────────────────────────────────────────────────────

st.divider()
st.subheader("Step 1 — Company mapping file")

mapping_file = st.file_uploader(
    "Upload mapping file",
    type=["xlsx"],
    key="mapping",
    help='Must contain "Company Name" and "Account" columns',
)

mapping_lookup = None
if mapping_file:
    try:
        mapping_lookup, display_df = load_mapping(mapping_file)
        st.success(f"✓ {mapping_file.name} — {len(mapping_lookup)} companies loaded")
        with st.expander("View mapping table"):
            st.dataframe(display_df, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Error reading mapping file: {e}")

# ── Step 2: Sales files ───────────────────────────────────────────────────────

st.divider()
st.subheader("Step 2 — Sales export files")

sales_files = st.file_uploader(
    "Upload one or more TicketVault exports",
    type=["xlsx"],
    accept_multiple_files=True,
    key="sales",
    help="Multiple files are combined and deduplicated automatically",
)

# ── Process ───────────────────────────────────────────────────────────────────

st.divider()

ready = mapping_lookup is not None and len(sales_files) > 0

if not ready:
    missing = []
    if mapping_lookup is None:
        missing.append("mapping file")
    if not sales_files:
        missing.append("at least one sales export")
    st.info(f"Waiting for: {' and '.join(missing)}")

if st.button("Process files", disabled=not ready, type="primary", use_container_width=True):
    with st.spinner("Processing..."):
        try:
            result = process_sales(sales_files, mapping_lookup)

            month_label   = result["month_label"]
            ys_rows       = result["ys_rows"]
            yitz_rows     = result["yitz_rows"]
            unmapped      = result["unmapped"]
            total_raw     = result["total_raw"]
            total_deduped = result["total_deduped"]
            total_filtered = result["total_filtered"]

            st.success(f"Done — {total_deduped:,} unique rows from {total_raw:,} total ({total_raw - total_deduped:,} duplicates removed, {total_deduped - total_filtered:,} zero-dollar rows excluded)")

            if unmapped:
                st.warning(
                    f"**{len(unmapped)} company name(s) not found in mapping file** — rows excluded. "
                    f"Add them to the mapping file and reprocess:\n\n"
                    + "\n".join(f"- `{c}`" for c in sorted(unmapped))
                )

            # Summary stats
            ys_sales   = sum(r["Ticket Sales"] for r in ys_rows)
            yitz_sales = sum(r["Ticket Sales"] for r in yitz_rows)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("ystickets rows",  f"{len(ys_rows):,}")
            col2.metric("ystickets sales", f"${ys_sales:,.2f}")
            col3.metric("yitzknopf rows",  f"{len(yitz_rows):,}")
            col4.metric("yitzknopf sales", f"${yitz_sales:,.2f}")

            # Customer breakdown
            with st.expander("Sales by customer"):
                breakdown = {}
                for r in ys_rows + yitz_rows:
                    c = r["Customer"]
                    breakdown.setdefault(c, {"Rows": 0, "Total Sales": 0.0})
                    breakdown[c]["Rows"] += 1
                    breakdown[c]["Total Sales"] += r["Ticket Sales"]
                bd_df = pd.DataFrame([
                    {"Customer": k, "Rows": v["Rows"], "Total Sales": f"${v['Total Sales']:,.2f}"}
                    for k, v in sorted(breakdown.items(), key=lambda x: -x[1]["Total Sales"])
                ])
                st.dataframe(bd_df, use_container_width=True, hide_index=True)

            st.divider()
            st.subheader(f"Download — {month_label}")

            col_a, col_b = st.columns(2)

            with col_a:
                ys_buf = BytesIO()
                write_excel(ys_rows, ys_buf)
                st.download_button(
                    label=f"⬇ {month_label} - ystickets.xlsx",
                    data=ys_buf.getvalue(),
                    file_name=f"{month_label} - ystickets.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            with col_b:
                yitz_buf = BytesIO()
                write_excel(yitz_rows, yitz_buf)
                st.download_button(
                    label=f"⬇ {month_label} - yitzknopf.xlsx",
                    data=yitz_buf.getvalue(),
                    file_name=f"{month_label} - yitzknopf.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Processing error: {e}")
            raise e

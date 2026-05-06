import streamlit as st
import pandas as pd
from io import BytesIO
from processor import load_mapping, process_sales, write_excel

st.set_page_config(
    page_title="TicketVault Report Processor",
    page_icon="🎟️",
    layout="centered",
)

st.title("🎟️ TicketVault Report Processor")
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
    # Clear any previous results when files are removed
    st.session_state.pop("result", None)
    st.session_state.pop("ys_bytes", None)
    st.session_state.pop("yitz_bytes", None)

if st.button("Process files", disabled=not ready, type="primary", use_container_width=True):
    with st.spinner("Processing..."):
        try:
            result = process_sales(sales_files, mapping_lookup)

            # Pre-generate both Excel files and store everything in session_state
            # so download button clicks (which trigger reruns) don't lose the data
            ys_buf = BytesIO()
            write_excel(result["ys_rows"], ys_buf)

            yitz_buf = BytesIO()
            write_excel(result["yitz_rows"], yitz_buf)

            st.session_state["result"]     = result
            st.session_state["ys_bytes"]   = ys_buf.getvalue()
            st.session_state["yitz_bytes"] = yitz_buf.getvalue()

        except Exception as e:
            st.error(f"Processing error: {e}")
            raise e

# ── Results (rendered from session_state so they survive download-button reruns) ──

if "result" in st.session_state:
    result        = st.session_state["result"]
    ys_bytes      = st.session_state["ys_bytes"]
    yitz_bytes    = st.session_state["yitz_bytes"]

    month_label    = result["month_label"]
    ys_rows        = result["ys_rows"]
    yitz_rows      = result["yitz_rows"]
    unmapped       = result["unmapped"]
    total_raw      = result["total_raw"]
    total_deduped  = result["total_deduped"]
    total_filtered = result["total_filtered"]

    st.success(
        f"Done — {total_deduped:,} unique rows from {total_raw:,} total "
        f"({total_raw - total_deduped:,} duplicates removed, "
        f"{total_deduped - total_filtered:,} zero-dollar rows excluded)"
    )

    if unmapped:
        st.warning(
            f"**{len(unmapped)} company name(s) not found in mapping file** — rows excluded. "
            f"Add them to the mapping file and reprocess:\n\n"
            + "\n".join(f"- `{c}`" for c in sorted(unmapped))
        )

    # Summary metrics
    ys_sales   = sum(r["Ticket Sales"] for r in ys_rows)
    yitz_sales = sum(r["Ticket Sales"] for r in yitz_rows)

    ys_sales   = sum(r["Ticket Sales"] for r in ys_rows)
    yitz_sales = sum(r["Ticket Sales"] for r in yitz_rows)

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**ystickets**")
        st.metric("Rows",  f"{len(ys_rows):,}")
        st.metric("Sales", f"${ys_sales:,.2f}")
    with col_r:
        st.markdown("**yitzknopf**")
        st.metric("Rows",  f"{len(yitz_rows):,}")
        st.metric("Sales", f"${yitz_sales:,.2f}")

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

    # Downloads
    st.divider()
    st.subheader(f"Download — {month_label}")

    col_a, col_b = st.columns(2)

    with col_a:
        st.download_button(
            label=f"⬇ {month_label} - ystickets.xlsx",
            data=ys_bytes,
            file_name=f"{month_label} - ystickets.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_b:
        st.download_button(
            label=f"⬇ {month_label} - yitzknopf.xlsx",
            data=yitz_bytes,
            file_name=f"{month_label} - yitzknopf.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

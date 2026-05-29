"""
Significance Triangle Sync — Streamlit web app
==============================================
Upload a survey data Excel file and a PowerPoint deck; the app inserts ▲/▼
significance markers into the deck's charts, tables, and text boxes, then lets
you download the updated deck plus a progress report.

Sidebar holds three groups of settings:
  • Core settings   — sheet name, default region/sub-column, min % diff, skip slides
  • Excel layout     — header rows, letter-code row, label column
  • Shape overrides  — per-shape custom column targeting (editable table)

Run locally:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import pandas as pd
import streamlit as st

import triangle_engine as engine


st.set_page_config(page_title="Significance Triangle Sync", page_icon="▲", layout="wide")

# ── Header ────────────────────────────────────────────────────────────────────
st.title("▲ Significance Triangle Sync")
st.caption(
    "Insert ▲/▼ significance markers into a PowerPoint deck from survey "
    "cross-tab data. Shapes are matched by their Selection-Pane names "
    "(those starting with “x ”)."
)


# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_int_list(raw):
    return [int(p.strip()) for p in str(raw).split(",") if p.strip().lstrip("-").isdigit()]


# ── Sidebar: Core settings ──────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")

    with st.expander("Core settings", expanded=True):
        excel_sheet = st.text_input("Data sheet name", value=engine.EXCEL_SHEET)
        default_region = st.text_input("Default region", value=engine.DEFAULT_REGION)
        default_sub = st.text_input("Default sub-column", value=engine.DEFAULT_SUB_COLUMN)
        min_diff = st.number_input(
            "Minimum % point difference", min_value=0, max_value=100,
            value=engine.MIN_DIFF_PCT,
        )
        skip_slides_raw = st.text_input(
            "Skip slides (comma-separated)",
            value=", ".join(str(s) for s in engine.SKIP_SLIDES),
        )

    with st.expander("Excel layout", expanded=False):
        st.caption(
            "Where things live in the data sheet. Only change if the data "
            "template's structure changes."
        )
        st.code(
            "Row 1-4  : Hierarchical column headers (Total/Country/Market -> Region -> Sub-col)\n"
            "Row 5    : 'Wave' label row\n"
            "Row 6    : 'Wave 4' / 'Wave 5' labels\n"
            "Row 7    : Letter codes (A, B, C, D ...)\n"
            "Row 8+   : Data, sig row immediately below each value row\n"
            "Column 1 : Section headers & row labels",
            language=None,
        )
        header_rows_raw = st.text_input(
            "Header rows (region / sub-col matching)",
            value=", ".join(str(r) for r in engine.HEADER_ROWS),
            help="Comma-separated, e.g. 1, 2, 3, 4, 6",
        )
        letter_row = st.number_input(
            "Letter-code row", min_value=1, max_value=100, value=engine.LETTER_ROW,
            help="Row holding the A, B, C significance letters.",
        )
        label_col = st.number_input(
            "Label column", min_value=1, max_value=100, value=engine.LABEL_COL,
            help="Column with section headers & row labels (1 = column A).",
        )
        st.caption("The Wave 4 / Wave 5 label row is found automatically within the header rows.")


# ── Main: file uploaders ─────────────────────────────────────────────────────
left, right = st.columns(2)
with left:
    excel_file = st.file_uploader("1 — Data Excel (.xlsx)", type=["xlsx"])
with right:
    pptx_file = st.file_uploader("2 — PowerPoint (.pptx)", type=["pptx"])


# ── Shape overrides editor ──────────────────────────────────────────────────────
st.subheader("Shape overrides")
st.caption(
    "Shapes that read from a specific column instead of the default region / "
    "sub-column (e.g. D5a from “All Bundle (NET) → Good Value”). Edit cells "
    "directly; use the ＋ at the bottom of the table to add a row."
)

if "overrides_df" not in st.session_state:
    st.session_state.overrides_df = pd.DataFrame(
        [
            {"Slide": s, "Shape name": name, "Region": region, "Sub-column": sub}
            for (s, name, region, sub) in engine.SHAPE_OVERRIDES
        ]
    )

overrides_df = st.data_editor(
    st.session_state.overrides_df,
    num_rows="dynamic",
    use_container_width=True,
    key="overrides_editor",
    column_config={
        "Slide": st.column_config.NumberColumn("Slide", min_value=1, step=1, width="small"),
        "Shape name": st.column_config.TextColumn("Shape name", width="large"),
        "Region": st.column_config.TextColumn("Region", width="medium"),
        "Sub-column": st.column_config.TextColumn("Sub-column", width="medium"),
    },
)

run = st.button(
    "Apply triangles", type="primary",
    disabled=not (excel_file and pptx_file),
)


# ── Run ───────────────────────────────────────────────────────────────────────
if run and excel_file and pptx_file:
    # Validate + collect overrides from the edited table
    overrides = []
    ov_errors = []
    for i, row in overrides_df.iterrows():
        slide = row.get("Slide")
        name  = (row.get("Shape name") or "")
        region = (row.get("Region") or "")
        sub    = (row.get("Sub-column") or "")
        # Skip fully blank rows the editor may leave behind
        if pd.isna(slide) and not str(name).strip():
            continue
        if pd.isna(slide) or not str(name).strip() or not str(region).strip() or not str(sub).strip():
            ov_errors.append(f"Override row {i + 1}: all four fields are required.")
            continue
        overrides.append((int(slide), str(name).strip(), str(region).strip(), str(sub).strip()))

    header_rows = parse_int_list(header_rows_raw)
    if not header_rows:
        st.error("Excel layout: header rows must list at least one row number.")
        st.stop()
    if ov_errors:
        for e in ov_errors:
            st.error(e)
        st.stop()

    # Push all settings into the engine
    engine.EXCEL_SHEET        = excel_sheet.strip()
    engine.DEFAULT_REGION     = default_region.strip()
    engine.DEFAULT_SUB_COLUMN = default_sub.strip()
    engine.MIN_DIFF_PCT       = int(min_diff)
    engine.SKIP_SLIDES        = parse_int_list(skip_slides_raw)
    engine.HEADER_ROWS        = header_rows
    engine.LETTER_ROW         = int(letter_row)
    engine.LABEL_COL          = int(label_col)
    engine.SHAPE_OVERRIDES    = overrides

    log_lines = []
    try:
        with st.spinner("Processing… this can take a few moments for large data files."):
            result = engine.run_sync(
                excel_file.getvalue(), pptx_file.getvalue(), log=log_lines.append,
            )
    except Exception as exc:
        st.error(f"Something went wrong: {exc}")
        with st.expander("Show log"):
            st.code("\n".join(log_lines) or "(no output)")
        st.stop()

    rows = result["report_rows"]
    n_ok    = sum(1 for r in rows if r["status"] == engine.STATUS_OK)
    n_nosig = sum(1 for r in rows if r["status"] == engine.STATUS_NO_SIG)
    n_warn  = sum(1 for r in rows if r["status"] not in (engine.STATUS_OK, engine.STATUS_NO_SIG))

    st.success(f"Done — {result['total']} triangles applied across {len(rows)} shapes.")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Triangles", result["total"])
    m2.metric("Shapes updated", n_ok)
    m3.metric("No sig", n_nosig)
    m4.metric("Warnings", n_warn)

    base = pptx_file.name.rsplit(".", 1)[0]
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "⬇ Updated PowerPoint", data=result["pptx_bytes"],
            file_name=f"{base} - updated.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )
    with d2:
        st.download_button(
            "⬇ Progress report (.xlsx)", data=result["report_bytes"],
            file_name="sync_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with st.expander("Per-shape results", expanded=True):
        st.dataframe(
            [
                {
                    "Slide": r["slide"], "Shape": r["shape_name"],
                    "Status": r["status"], "Applied": r.get("items_applied", ""),
                }
                for r in rows
            ],
            use_container_width=True, hide_index=True,
        )

    with st.expander("Processing log"):
        st.code("\n".join(log_lines))

else:
    st.info("Upload both files to enable the **Apply triangles** button.")

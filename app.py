"""
Significance Triangle Sync — Streamlit web app
==============================================
Upload a survey data Excel file and a PowerPoint deck; the app inserts ▲/▼
significance markers into the deck's charts, tables, and text boxes, then lets
you download the updated deck plus a progress report.

Flow:
  1. Configure the Excel layout (sheet + where the header/letter/label rows are).
  2. The app scans the hierarchical column headers and offers cascading dropdowns
     for the default column cut.
  3. Optionally add per-shape overrides (same cascading picker per row).
  4. Upload the PowerPoint and run.

Run locally:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import openpyxl
import pandas as pd
import streamlit as st

import triangle_engine as engine


st.set_page_config(page_title="Significance Triangle Sync", page_icon="▲", layout="wide")

st.title("▲ Significance Triangle Sync")
st.caption(
    "Insert ▲/▼ significance markers into a PowerPoint deck from survey "
    "cross-tab data. Shapes are matched by their Selection-Pane names "
    "(those starting with “x ”)."
)


# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_int_list(raw):
    return [int(p.strip()) for p in str(raw).split(",") if p.strip().lstrip("-").isdigit()]


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    return list(wb.sheetnames)


@st.cache_data(show_spinner=False)
def build_tree(file_bytes, sheet, header_rows, letter_row, label_col):
    """Load the chosen sheet under the given layout and build the header tree."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb[sheet]
    # Apply layout to the engine so its scanners use the right rows
    engine.HEADER_ROWS = list(header_rows)
    engine.LETTER_ROW  = int(letter_row)
    engine.LABEL_COL   = int(label_col)
    engine.get_col_ancestry.__defaults__[0].clear()
    return engine.build_header_tree(ws)


def cascade_options(paths, chosen):
    """
    Given the list of full path tuples and the values chosen so far (a list,
    one per level, '' = not chosen yet), return the valid options for the NEXT
    unchosen level — i.e. the distinct values at that level among paths whose
    earlier levels match what's chosen.
    """
    level = len(chosen)
    opts = []
    for p in paths:
        if level >= len(p):
            continue
        if all(p[i] == chosen[i] for i in range(level)):
            opts.append(p[level])
    # Preserve order, drop dups
    seen, out = set(), []
    for o in opts:
        if o not in seen:
            seen.add(o); out.append(o)
    return out


def cascade_picker(paths, levels, key_prefix, defaults=None):
    """
    Render `levels` cascading selectboxes. Each becomes active only once the
    previous one is chosen. Returns the chosen path as a list (may contain ''
    for legitimately blank levels). Blank ('') options are shown as “—”.
    """
    BLANK = "—"
    chosen = []
    defaults = defaults or []
    for lvl in range(levels):
        opts = cascade_options(paths, chosen)
        if not opts:
            break
        display = [BLANK if o == "" else o for o in opts]
        # Default selection if provided and valid at this level
        idx = 0
        if lvl < len(defaults):
            want = defaults[lvl]
            disp_want = BLANK if want == "" else want
            if disp_want in display:
                idx = display.index(disp_want)
        sel = st.selectbox(
            f"Level {lvl + 1}", display, index=idx,
            key=f"{key_prefix}_lvl{lvl}",
        )
        chosen.append("" if sel == BLANK else sel)
    return chosen


# ── File uploaders (one each) ─────────────────────────────────────────────────
left, right = st.columns(2)
with left:
    excel_file = st.file_uploader(
        "Data Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)
with right:
    pptx_file = st.file_uploader(
        "PowerPoint (.pptx)", type=["pptx"], accept_multiple_files=False)

if excel_file is None:
    st.info("Upload the data Excel to begin — the column settings are built from it.")
    st.stop()

excel_bytes = excel_file.getvalue()
try:
    sheets = list_sheets(excel_bytes)
except Exception as exc:
    st.error(f"Couldn't read that Excel file: {exc}")
    st.stop()


# ── STEP 1 — Advanced / layout settings (must be set first) ───────────────────
st.subheader("1 · Excel layout")
st.caption(
    "Tell the tool where things live in the data sheet. The column pickers below "
    "are built from this, so set it first."
)

lc1, lc2, lc3, lc4 = st.columns([2, 1.4, 1, 1])
with lc1:
    sheet = st.selectbox(
        "Data sheet", sheets,
        index=sheets.index(engine.EXCEL_SHEET) if engine.EXCEL_SHEET in sheets else 0,
    )
with lc2:
    header_rows_raw = st.text_input(
        "Header rows", value=", ".join(str(r) for r in engine.HEADER_ROWS),
        help="Comma-separated rows holding the column hierarchy, e.g. 1, 2, 3, 4, 6",
    )
with lc3:
    letter_row = st.number_input("Letter row", min_value=1, max_value=100,
                                 value=engine.LETTER_ROW)
with lc4:
    label_col = st.number_input("Label col", min_value=1, max_value=100,
                                value=engine.LABEL_COL)

with st.expander("What the layout means"):
    st.code(
        "Row 1-4  : Hierarchical column headers (Country/Market -> Region -> Group -> Sub-col)\n"
        "Row 5    : 'Wave' label row\n"
        "Row 6    : 'Wave 4' / 'Wave 5' labels\n"
        "Row 7    : Letter codes (A, B, C, D ...)\n"
        "Row 8+   : Data, sig row immediately below each value row\n"
        "Column 1 : Section headers & row labels",
        language=None,
    )

header_rows = parse_int_list(header_rows_raw)
if not header_rows:
    st.error("Header rows must list at least one row number, e.g. 1, 2, 3, 4, 6.")
    st.stop()

# Build the hierarchy tree from the chosen layout
try:
    tree = build_tree(excel_bytes, sheet, tuple(header_rows), int(letter_row), int(label_col))
except Exception as exc:
    st.error(f"Couldn't read the header hierarchy with these layout settings: {exc}")
    st.stop()

paths_raw = tree["paths"]
paths  = [p for (p, _col) in paths_raw]   # bare path tuples for cascading
levels = tree["levels"]
if not paths:
    st.error(
        "No Wave 4 / Wave 5 columns were found with these layout settings. "
        "Check the header rows and sheet."
    )
    st.stop()

st.success(f"Found {len(paths)} column cuts across {levels} hierarchy levels.")


# ── STEP 2 — Core settings (cascading dropdowns, unlocked after layout) ───────
st.subheader("2 · Default column cut")
st.caption(
    "Pick the column the tool uses for every shape unless overridden. Each level "
    "unlocks the next."
)

BLANK = "—"
chosen_default = []
cols = st.columns(levels)
for lvl in range(levels):
    opts = cascade_options(paths, chosen_default)
    if not opts:
        # No further levels apply for this branch — show a disabled placeholder
        with cols[lvl]:
            st.selectbox(f"Level {lvl + 1}", ["—"], index=0,
                         key=f"def_lvl{lvl}_disabled", disabled=True)
        chosen_default.append("")
        continue
    display = [BLANK if o == "" else o for o in opts]
    key = f"def_lvl{lvl}"
    # If the persisted choice is no longer valid under the parent selections,
    # drop it so the box falls back to the first valid option (true cascade).
    if key in st.session_state and st.session_state[key] not in display:
        del st.session_state[key]
    with cols[lvl]:
        sel = st.selectbox(f"Level {lvl + 1}", display, key=key)
    chosen_default.append("" if sel == BLANK else sel)

st.caption("Selected cut → " + " > ".join(c if c else "—" for c in chosen_default))

cset1, cset2 = st.columns(2)
with cset1:
    min_diff = st.number_input("Minimum % point difference", min_value=0, max_value=100,
                               value=engine.MIN_DIFF_PCT)
with cset2:
    skip_slides_raw = st.text_input(
        "Skip slides (comma-separated)",
        value=", ".join(str(s) for s in engine.SKIP_SLIDES))


# ── STEP 3 — Shape overrides ──────────────────────────────────────────────────
st.subheader("3 · Shape overrides (optional)")
st.caption(
    "Shapes that read from a different column than the default. Each row picks a "
    "full column path. Add a row, then choose its levels."
)

# Build level-value option lists for the editable table. To keep the table
# usable we offer ALL values seen at each level; invalid combinations are
# validated on run.
level_values = []
for lvl in range(levels):
    seen, vals = set(), []
    for p in paths:
        if lvl < len(p):
            v = p[lvl]
            disp = "—" if v == "" else v
            if disp not in seen:
                seen.add(disp); vals.append(disp)
    level_values.append(vals)

if "overrides_df" not in st.session_state:
    init_rows = []
    for entry in engine.SHAPE_OVERRIDES:
        # Entries may be legacy (slide, name, region, sub) or (slide, name, [path])
        if len(entry) == 4:
            s, name, c1, c2 = entry
            wanted = [v for v in (c1, c2) if v]
        else:
            s, name, path_list = entry
            wanted = [v for v in path_list if v]
        row = {"Slide": s, "Shape name": name}
        # Best-effort: find the full path whose values contain all wanted values
        match = next((p for p in paths if all(w in p for w in wanted)), None)
        for lvl in range(levels):
            if match is not None:
                row[f"Sub column {lvl + 1}"] = "—" if match[lvl] == "" else match[lvl]
            else:
                row[f"Sub column {lvl + 1}"] = ""
        init_rows.append(row)
    st.session_state.overrides_df = pd.DataFrame(init_rows)

col_cfg = {
    "Slide": st.column_config.NumberColumn("Slide", min_value=1, step=1, width="small"),
    "Shape name": st.column_config.TextColumn("Shape name", width="large"),
}
for lvl in range(levels):
    col_cfg[f"Sub column {lvl + 1}"] = st.column_config.SelectboxColumn(
        f"Sub column {lvl + 1}", options=[""] + level_values[lvl], width="medium")

overrides_df = st.data_editor(
    st.session_state.overrides_df,
    num_rows="dynamic", use_container_width=True, key="overrides_editor",
    column_config=col_cfg,
)


# ── STEP 4 — Run ──────────────────────────────────────────────────────────────
st.subheader("4 · Run")
run = st.button("Apply triangles", type="primary", disabled=pptx_file is None)
if pptx_file is None:
    st.info("Upload the PowerPoint to enable the run button.")

if run and pptx_file is not None:
    # Resolve overrides into (slide, shape, [path levels])
    overrides = []
    ov_errors = []
    for i, row in overrides_df.iterrows():
        slide = row.get("Slide")
        name  = (row.get("Shape name") or "")
        if pd.isna(slide) and not str(name).strip():
            continue
        if pd.isna(slide) or not str(name).strip():
            ov_errors.append(f"Override row {i + 1}: slide and shape name are required.")
            continue
        path_vals = []
        for lvl in range(levels):
            v = row.get(f"Sub column {lvl + 1}", "")
            v = "" if (v is None or (isinstance(v, float) and pd.isna(v)) or v == "—") else str(v).strip()
            path_vals.append(v)
        if not any(path_vals):
            ov_errors.append(f"Override row {i + 1}: choose at least one column level.")
            continue
        overrides.append((int(slide), str(name).strip(), path_vals))

    if ov_errors:
        for e in ov_errors:
            st.error(e)
        st.stop()

    # Push settings into the engine
    engine.EXCEL_SHEET     = sheet
    engine.MIN_DIFF_PCT    = int(min_diff)
    engine.SKIP_SLIDES     = parse_int_list(skip_slides_raw)
    engine.HEADER_ROWS     = header_rows
    engine.LETTER_ROW      = int(letter_row)
    engine.LABEL_COL       = int(label_col)
    # Default cut + overrides now carry full paths
    engine.DEFAULT_CUT     = [c for c in chosen_default]
    engine.SHAPE_OVERRIDES = overrides
    # Keep the legacy two fields populated from the first/last chosen levels so
    # any code still referencing them stays sane.
    nonblank = [c for c in chosen_default if c]
    engine.DEFAULT_REGION     = nonblank[0] if nonblank else ""
    engine.DEFAULT_SUB_COLUMN = nonblank[-1] if nonblank else ""

    log_lines = []
    try:
        with st.spinner("Processing…"):
            result = engine.run_sync(excel_bytes, pptx_file.getvalue(), log=log_lines.append)
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
            type="primary")
    with d2:
        st.download_button(
            "⬇ Progress report (.xlsx)", data=result["report_bytes"],
            file_name="sync_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with st.expander("Per-shape results", expanded=True):
        st.dataframe(
            [{"Slide": r["slide"], "Shape": r["shape_name"], "Status": r["status"],
              "Applied": r.get("items_applied", "")} for r in rows],
            use_container_width=True, hide_index=True)

    with st.expander("Processing log"):
        st.code("\n".join(log_lines))

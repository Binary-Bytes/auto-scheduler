import io

import streamlit as st
import pandas as pd
import openpyxl

from backend2 import (
    load_data_rows, required_weekdays, build_workbook, WEEKDAY_NAMES,
    detect_start_year, shift_rows_to_year, suggest_term_label,
    Theme, PRESET_ACCENTS, FONT_CHOICES, BORDER_CHOICES, DENSITY_CHOICES, DEFAULT_ACCENT,
)

st.set_page_config(page_title="Night School Weekday Schedule Auto-Filler", layout="wide")
st.title("🌙 North Night School Scheduler")
st.caption(
    "Upload the course list, fill in Room + Time for each course, and get back one sheet "
    "per weekday with every class's active weeks shaded automatically."
)

st.markdown(
    """
    <style>
    div[data-testid="stForm"] { border: 2px solid #2F5597; border-radius: 12px; padding: 1rem 1.2rem; }
    .stButton > button, .stDownloadButton > button { border-radius: 10px; font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Step 1: Upload
# ----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload the workbook (must contain a 'Data' sheet with Starts/Ends/Course/Name/Days/Instructor columns)",
    type=["xlsx"],
)

if uploaded_file is None:
    st.info("Waiting for a workbook to be uploaded...")
    st.stop()

if "wb_bytes" not in st.session_state or st.session_state.get("uploaded_name") != uploaded_file.name:
    st.session_state["wb_bytes"] = uploaded_file.getvalue()
    st.session_state["uploaded_name"] = uploaded_file.name
    st.session_state.pop("assignments_df2", None)

try:
    wb_in = openpyxl.load_workbook(io.BytesIO(st.session_state["wb_bytes"]))
except Exception as e:
    st.error(f"Could not open this file as an Excel workbook: {e}")
    st.stop()

if "Data" not in wb_in.sheetnames:
    st.error("No sheet named 'Data' found. Please make sure your course list sheet is named exactly 'Data'.")
    st.stop()

try:
    data_rows = load_data_rows(wb_in["Data"])
except Exception as e:
    st.error(f"Could not read the 'Data' sheet: {e}")
    st.stop()

if not data_rows:
    st.warning("No course rows found in the 'Data' sheet.")
    st.stop()

needed_weekdays = required_weekdays(data_rows)
st.success(
    f"Loaded **{len(data_rows)}** course rows. "
    f"Sheets will be created for: **{', '.join(WEEKDAY_NAMES[i] for i in needed_weekdays)}**."
)

school_name = st.text_input("School name (used in each sheet's title bar)", value="Night School")

# Year / term override: the dates in the uploaded file may carry the wrong year
# (e.g. a Jan-Jun 2027 list typed in as 2026). Everything downstream - week
# columns, month headers, Start/End cells, sheet title - follows this choice.
file_year = detect_start_year(data_rows)
yc1, yc2 = st.columns(2)
term_year = yc1.number_input(
    "Year of the term's first class",
    min_value=2000, max_value=2100, value=file_year, step=1,
    help=f"The uploaded file's dates say {file_year}. Change this to move every date to the correct year.",
)
term_label = yc2.text_input("Term name (used in each sheet's title bar)", value=suggest_term_label(data_rows))
term_label = term_label.strip().upper() or suggest_term_label(data_rows)

if term_year != file_year:
    data_rows = shift_rows_to_year(data_rows, int(term_year))
    st.info(f"Dates shifted from {file_year} to {int(term_year)}.")
first_start = min(c["start"] for c in data_rows)
last_end = max(c["end"] for c in data_rows)
st.caption(f"Term runs {first_start:%a %b %d, %Y} \u2192 {last_end:%a %b %d, %Y}.")

# ----------------------------------------------------------------------------
# Appearance of the generated Excel file
# ----------------------------------------------------------------------------
with st.expander("🎨 Customise the Excel look (optional)"):
    c1, c2, c3, c4 = st.columns([1.3, 1, 1, 1])
    preset = c1.selectbox("Colour theme", list(PRESET_ACCENTS) + ["Custom colour..."], index=0)
    if preset == "Custom colour...":
        accent_in = c1.color_picker("Pick an accent colour", value="#" + DEFAULT_ACCENT)
        c1.caption("Very light colours are automatically deepened so the white header text stays readable.")
    else:
        accent_in = PRESET_ACCENTS[preset]
    font_choice = c2.selectbox("Font", FONT_CHOICES, index=0)
    border_choice = c3.selectbox("Border weight", BORDER_CHOICES, index=0)
    density_choice = c4.selectbox("Row height", DENSITY_CHOICES, index=1)
    zebra_choice = st.checkbox("Alternate row shading (zebra stripes)", value=True)

    theme = Theme(accent=accent_in, font=font_choice, border=border_choice,
                  density=density_choice, zebra=zebra_choice)

    # Mini preview built from the theme's real palette
    _p = theme.palette
    _row = lambda bg, cells: "".join(
        f'<td style="background:{c};border:1px solid #A6A6A6;height:{18 if density_choice == "Compact" else 26 if density_choice == "Standard" else 34}px;'
        f'text-align:center;color:#fff;font-weight:700;font-size:11px">{t}</td>' for c, t in cells)
    _bw = 3 if border_choice == "Bold" else 2
    st.markdown(
        f"""
        <div style="font-family:{font_choice},sans-serif;max-width:560px;border:{_bw}px solid #{_p['dark']};">
          <div style="background:#{_p['dark']};color:#fff;text-align:center;font-weight:700;padding:6px">Night School &ndash; Room Schedule &ndash; Preview</div>
          <table style="border-collapse:collapse;width:100%;font-size:12px">
            <tr><td style="background:#{_p['accent']};color:#fff;font-weight:700;padding:4px 8px;border:1px solid #A6A6A6">Course</td>
                <td style="background:#{_p['month']};color:#{_p['dark']};font-weight:700;text-align:center;border:1px solid #A6A6A6" colspan="6">JANUARY</td></tr>
            <tr><td style="background:#{_p['accent']};border:1px solid #A6A6A6"></td>
                {''.join(f'<td style="background:#{_p["date"]};color:#{_p["dark"]};font-weight:700;text-align:center;border:1px solid #A6A6A6">{d}</td>' for d in (4, 11, 18, 25, 1, 8))}</tr>
            <tr><td style="padding:4px 8px;font-weight:700;border:1px solid #A6A6A6">Bookkeeping (MW)</td>
                {_row('', [('#'+_p['accent'], 'S'), ('#'+_p['shade'], ''), ('#'+_p['shade'], ''), ('#'+_p['shade'], ''), ('#'+_p['shade'], ''), ('#'+_p['accent'], 'E')])}</tr>
            <tr style="background:{'#' + _p['zebra'] if zebra_choice else 'transparent'}"><td style="padding:4px 8px;font-weight:700;border:1px solid #A6A6A6">Korean - Level 1 (M)</td>
                {_row('', [('#'+_p['accent'], 'S'), ('#'+_p['shade'], ''), ('#'+_p['shade'], ''), ('#'+_p['shade'], ''), ('#'+_p['shade'], ''), ('#'+_p['accent'], 'E')])}</tr>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------------------------
# Step 2: Editable Room / Time table (one row per course - not per weekday,
# since a course meeting on multiple days uses the same room & time on all of them)
# ----------------------------------------------------------------------------
st.subheader("1. Add Room and Time (optional)")
st.caption(
    "Fill in whatever you already know - any Room or Time left blank simply stays blank in the Excel file, "
    "so you can complete it there later. Each entry applies to every weekday that course meets on."
)

if "assignments_df2" not in st.session_state:
    records = []
    for course in data_rows:
        records.append({
            "Course Row": course["row"],
            "Course #": course["course_raw"],
            "Name": course["name"],
            "Days": course["days"],
            "Instructor": course["instructor"],
            "Start Date": str(course["start"]),
            "End Date": str(course["end"]),
            "Notes": course["notes"],
            "Room": "",
            "Time": "",
        })
    st.session_state["assignments_df2"] = pd.DataFrame(records)

# Keep the Start/End columns in sync with the chosen year (rooms/times untouched)
_df = st.session_state["assignments_df2"]
_df["Start Date"] = [str(c["start"]) for c in data_rows]
_df["End Date"] = [str(c["end"]) for c in data_rows]

with st.form("assignment_form2"):
    edited_df = st.data_editor(
        st.session_state["assignments_df2"],
        column_config={
            "Course Row": st.column_config.NumberColumn(disabled=True),
            "Course #": st.column_config.TextColumn(disabled=True),
            "Name": st.column_config.TextColumn(disabled=True),
            "Days": st.column_config.TextColumn(disabled=True),
            "Instructor": st.column_config.TextColumn(disabled=True),
            "Start Date": st.column_config.TextColumn(disabled=True),
            "End Date": st.column_config.TextColumn(disabled=True),
            "Notes": st.column_config.TextColumn(disabled=True),
            "Room": st.column_config.TextColumn(help="e.g. 2104"),
            "Time": st.column_config.TextColumn(help="e.g. 6:15-9:15"),
        },
        hide_index=True,
        use_container_width=True,
        key="assignment_editor2",
    )
    generate_clicked = st.form_submit_button("Generate Weekday Schedules", type="primary")

st.session_state["assignments_df2"] = edited_df

# ----------------------------------------------------------------------------
# Step 3: Generate
# ----------------------------------------------------------------------------
st.subheader("2. Generate")

if generate_clicked:
    def _clean(v):
        return "" if v is None or pd.isna(v) else str(v).strip()

    with st.spinner("Building the schedule..."):
        assignments = {}
        for _, row in edited_df.iterrows():
            assignments[int(row["Course Row"])] = {"room": _clean(row["Room"]), "time": _clean(row["Time"])}

        wb_out, placed_counts = build_workbook(data_rows, assignments, school_name=school_name, term_label=term_label, theme=theme)

        out_buffer = io.BytesIO()
        wb_out.save(out_buffer)
        out_buffer.seek(0)

    total_placed = sum(placed_counts.values())
    st.success(
        f"Done! Created {len(wb_out.sheetnames)} weekday sheet(s) "
        f"({', '.join(wb_out.sheetnames)}) and shaded {total_placed} class-week cell(s) total."
    )
    blank_room = int(edited_df["Room"].map(_clean).eq("").sum())
    blank_time = int(edited_df["Time"].map(_clean).eq("").sum())
    if blank_room or blank_time:
        st.info(f"{blank_room} course(s) have no Room and {blank_time} have no Time - those cells are left blank in the file.")
    st.caption("Each course's first class day is marked **S** and its last class day **E** (on the one weekday sheet where each falls).")

    st.download_button(
        "⬇️ Download Weekday Schedule (.xlsx)",
        data=out_buffer,
        file_name="Night_School_Weekday_Schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

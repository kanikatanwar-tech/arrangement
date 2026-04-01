import streamlit as st
import pandas as pd
import math
import re
import io
from datetime import date
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Marking Arrangement", layout="wide")
st.title("📋 Teacher Marking Arrangement")

PERIOD_COLS = ["P0","P1","P2","P3","P4","P5","P6","P7","P8"]
DAY_MAP = {
    "Monday":"MON","Tuesday":"TUE","Wednesday":"WED",
    "Thursday":"THU","Friday":"FRI","Saturday":"SAT"
}
FIRST_HALF  = PERIOD_COLS[:math.ceil(len(PERIOD_COLS)/2)]   # P0-P4
SECOND_HALF = PERIOD_COLS[math.ceil(len(PERIOD_COLS)/2):]   # P5-P8

# ═══════════════════════════════════════════════
#  PURE HELPERS
# ═══════════════════════════════════════════════

def clean_name(raw: str) -> str:
    m = re.match(r'^(.+?)-(\d{7,8})\s*$', raw.strip())
    return m.group(1).strip() if m else raw.strip()

def extract_teacher_keys(cell: str) -> list:
    if not cell:
        return []
    results = []
    for chunk in re.findall(r'\(([^)]+)\)', cell):
        for part in chunk.split('/'):
            part = part.strip()
            if re.search(r'-[A-Z0-9]{2,}', part, re.IGNORECASE):
                results.append(part)
    if not results:
        for m in re.findall(r'([A-Z][A-Za-z\s]+-(?:\d{5,8}|[A-Z]{2,6}))', cell):
            results.append(m.strip())
    return results

def get_half_cols(leave_type: str) -> list:
    if leave_type == "Full Day":     return PERIOD_COLS[:]
    elif leave_type == "First Half": return FIRST_HALF[:]
    else:                            return SECOND_HALF[:]

def which_half(pc: str) -> str:
    return "First Half" if pc in FIRST_HALF else "Second Half"

def get_class_from_cell(cell: str):
    if not cell: return None
    m = re.match(r'^(\d+[A-Z])', cell.strip())
    return m.group(1) if m else None

# ═══════════════════════════════════════════════
#  DATA LOADING
# ═══════════════════════════════════════════════

@st.cache_data
def load_data(file_bytes: bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))

    # ── Teacher TT ──────────────────────────────
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Teacher TT Summary", header=None)
    tt_rows, cur = [], None
    for _, row in raw.iloc[2:].iterrows():
        if pd.notna(row[0]) and str(row[0]).strip():
            cur = str(row[0]).strip()
        if pd.isna(row[1]) or cur is None:
            continue
        rec = {"Teacher": cur, "Day": str(row[1]).strip()}
        for i, pc in enumerate(PERIOD_COLS):
            v = row[i+2] if i+2 < len(row) else None
            rec[pc] = str(v).strip() if pd.notna(v) and str(v).strip() not in ["nan",""] else ""
        tt_rows.append(rec)
    teacher_tt = pd.DataFrame(tt_rows)

    # ── Class TT ────────────────────────────────
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Class TT Summary", header=None)
    ct_rows, cur = [], None
    for _, row in raw.iloc[2:].iterrows():
        if pd.notna(row[0]) and str(row[0]).strip():
            cur = str(row[0]).strip()
        if pd.isna(row[1]) or cur is None:
            continue
        rec = {"Class": cur, "Day": str(row[1]).strip()}
        for i, pc in enumerate(PERIOD_COLS):
            v = row[i+2] if i+2 < len(row) else None
            rec[pc] = str(v).strip() if pd.notna(v) and str(v).strip() not in ["nan",""] else ""
        ct_rows.append(rec)
    class_tt = pd.DataFrame(ct_rows)

    # ── Unavailability ───────────────────────────
    unavail = {}   # { raw_teacher_key: set of period col names }
    if "Unavailability" in xl.sheet_names:
        raw_u = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Unavailability", header=None)
        for _, row in raw_u.iloc[1:].iterrows():
            t, p = row[0], row[1]
            if pd.isna(t) or pd.isna(p): continue
            t = str(t).strip()
            pset = set()
            for x in str(p).split(','):
                x = x.strip()
                if x.isdigit():
                    pset.add(f"P{x}")
            if pset:
                unavail[t] = pset

    # ── Name maps ────────────────────────────────
    disp2raw, raw2disp = {}, {}
    for rk in teacher_tt["Teacher"].unique():
        disp = clean_name(rk)
        disp2raw.setdefault(disp, []).append(rk)
        raw2disp[rk] = disp

    return teacher_tt, class_tt, unavail, disp2raw, raw2disp

# ═══════════════════════════════════════════════
#  BUSY COUNTING
# ═══════════════════════════════════════════════

def get_unavail_periods(raw_keys: list, unavail: dict) -> set:
    unav = set()
    for rk in raw_keys:
        for uk, uperiods in unavail.items():
            if uk == rk or clean_name(uk) == clean_name(rk):
                unav |= uperiods
    return unav

def count_busy(raw_keys: list, day: str, period_subset: list,
               teacher_tt, unavail: dict, extra_busy: dict) -> dict:
    """
    Count busy periods within period_subset.
    extra_busy: { display_name: set(period_cols) } — dynamically assigned arrangements.
    """
    tt_busy = set()
    for rk in raw_keys:
        rows = teacher_tt[(teacher_tt["Teacher"]==rk) & (teacher_tt["Day"]==day)]
        if not rows.empty:
            r = rows.iloc[0]
            for pc in period_subset:
                if r.get(pc, ""):
                    tt_busy.add(pc)

    unav_set  = get_unavail_periods(raw_keys, unavail)
    unav_busy = {pc for pc in period_subset if pc in unav_set}

    extra_set = set()
    for rk in raw_keys:
        disp = clean_name(rk)
        if disp in extra_busy:
            extra_set |= (extra_busy[disp] & set(period_subset))

    total = tt_busy | unav_busy | extra_set
    return {"tt": len(tt_busy), "unav": len(unav_busy),
            "extra": len(extra_set), "total": len(total)}

def is_free(raw_keys: list, day: str, pc: str,
            teacher_tt, unavail: dict, extra_busy: dict) -> bool:
    """True only if teacher has no class, not unavailable, and not already assigned."""
    for rk in raw_keys:
        rows = teacher_tt[(teacher_tt["Teacher"]==rk) & (teacher_tt["Day"]==day)]
        if not rows.empty and rows.iloc[0].get(pc, ""):
            return False
    if pc in get_unavail_periods(raw_keys, unavail):
        return False
    for rk in raw_keys:
        disp = clean_name(rk)
        if disp in extra_busy and pc in extra_busy[disp]:
            return False
    return True

# ═══════════════════════════════════════════════
#  SUGGESTION ENGINE
# ═══════════════════════════════════════════════

def suggest_teacher(mode: str, absent_raw_keys: list, pc: str, day: str,
                    teacher_tt, class_tt, unavail: dict,
                    disp2raw: dict, raw2disp: dict,
                    all_on_leave_raw: list,
                    extra_busy: dict) -> tuple:
    """Returns (best_display_name | None, reason_str)."""
    half_cols = FIRST_HALF if pc in FIRST_HALF else SECOND_HALF

    # ── Build candidate pool ─────────────────────
    if mode == "Same Class Teacher":
        klass = None
        for rk in absent_raw_keys:
            rows = teacher_tt[(teacher_tt["Teacher"]==rk) & (teacher_tt["Day"]==day)]
            if not rows.empty:
                klass = get_class_from_cell(rows.iloc[0].get(pc,""))
                break
        if not klass:
            return None, "Could not determine class for this period"
        ct_rows = class_tt[(class_tt["Class"]==klass) & (class_tt["Day"]==day)]
        if ct_rows.empty:
            return None, f"No class TT found for {klass}"
        ct_row = ct_rows.iloc[0]
        cand_disps = set()
        for col in PERIOD_COLS:
            for tk in extract_teacher_keys(ct_row.get(col,"")):
                disp = clean_name(tk)
                if disp in disp2raw:
                    cand_disps.add(disp)
    else:  # Any Teacher
        cand_disps = set(disp2raw.keys())

    # ── Exclude on-leave and absent teachers ─────
    exclude = ({clean_name(rk) for rk in all_on_leave_raw}
               | {clean_name(rk) for rk in absent_raw_keys})
    cand_disps -= exclude

    # ── Filter: must be free in this specific period ─
    free_cands = [
        d for d in cand_disps
        if is_free(disp2raw.get(d, [d]), day, pc, teacher_tt, unavail, extra_busy)
    ]
    if not free_cands:
        return None, "No free teacher available for this period"

    # ── Rank by fewest busy in this half ─────────
    scored = []
    for d in free_cands:
        info = count_busy(disp2raw.get(d,[d]), day, half_cols,
                          teacher_tt, unavail, extra_busy)
        scored.append((d, info["total"]))
    scored.sort(key=lambda x: x[1])

    best, cnt = scored[0]
    return best, f"Least busy in {which_half(pc)}: {cnt} period(s)"

# ═══════════════════════════════════════════════
#  REMARKS  (workload info of a specific teacher)
# ═══════════════════════════════════════════════

def build_remarks(disp: str, day: str, leave_type: str,
                  disp2raw: dict, teacher_tt, unavail: dict,
                  extra_busy: dict) -> str:
    if not disp: return "—"
    raws = disp2raw.get(disp, [disp])
    fhi  = count_busy(raws, day, FIRST_HALF,  teacher_tt, unavail, extra_busy)
    shi  = count_busy(raws, day, SECOND_HALF, teacher_tt, unavail, extra_busy)

    def fmt(info, n, label):
        parts = [f"TT: **{info['tt']}**", f"Unavail: **{info['unav']}**"]
        if info["extra"] > 0:
            parts.append(f"Arranged: **{info['extra']}**")
        parts.append(f"Total busy {label}: **{info['total']}/{n}**")
        return "  |  ".join(parts)

    if leave_type == "First Half":
        return fmt(fhi, len(FIRST_HALF), "1st Half")
    elif leave_type == "Second Half":
        return fmt(shi, len(SECOND_HALF), "2nd Half")
    else:
        total = fhi["total"] + shi["total"]
        return (f"1st Half: **{fhi['total']}/{len(FIRST_HALF)}**  |  "
                f"2nd Half: **{shi['total']}/{len(SECOND_HALF)}**  |  "
                f"Full Day total: **{total}/{len(PERIOD_COLS)}**"
                + (f"  |  Arranged: **{fhi['extra']+shi['extra']}**"
                   if fhi["extra"]+shi["extra"] > 0 else ""))

# ═══════════════════════════════════════════════
#  EXCEL EXPORT — APPENDED (original + new sheet)
# ═══════════════════════════════════════════════

def _write_arrangement_sheet(ws, sel_date: date, day_code: str,
                              arrangement_rows: list, arr_state: list):
    """Writes the arrangement data into an openpyxl worksheet."""
    heading_font    = Font(bold=True, size=13, color="FFFFFF")
    heading_fill    = PatternFill("solid", fgColor="1F3864")
    col_hdr_font    = Font(bold=True, size=11, color="FFFFFF")
    col_hdr_fill    = PatternFill("solid", fgColor="2E75B6")
    leave_font      = Font(bold=True)
    leave_fill      = PatternFill("solid", fgColor="FCE4D6")
    arranged_fill   = PatternFill("solid", fgColor="E2EFDA")
    pending_fill    = PatternFill("solid", fgColor="FFF2CC")
    dash_fill       = PatternFill("solid", fgColor="F2F2F2")
    center_align    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin            = Side(style="thin")
    border          = Border(left=thin, right=thin, top=thin, bottom=thin)

    total_cols      = 1 + len(PERIOD_COLS)
    last_col        = get_column_letter(total_cols)

    # Row 1: heading
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value     = (f"Arrangement Sheet  |  "
                   f"Date: {sel_date.strftime('%d %B %Y')}  |  "
                   f"Day: {sel_date.strftime('%A')}  ({day_code})")
    c.font      = heading_font
    c.fill      = heading_fill
    c.alignment = center_align
    ws.row_dimensions[1].height = 26

    # Row 2: column headers
    for ci, h in enumerate(["Teacher on Leave"] + PERIOD_COLS, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font=col_hdr_font; c.fill=col_hdr_fill
        c.alignment=center_align; c.border=border
    ws.row_dimensions[2].height = 20

    # Build teacher → period map
    teacher_period_map = defaultdict(dict)
    for row, state in zip(arrangement_rows, arr_state):
        if not row["period_col"]: continue
        method      = state["method"]
        arranged_by = (state["manual"] if method == "Manual Entry"
                       else state.get("_locked_suggested") or state.get("_suggested") or "")
        teacher_period_map[row["teacher"]][row["period_col"]] = {
            "class_subject": row["class_subject"],
            "arranged_by"  : arranged_by,
            "saved"        : state["saved"],
        }

    # Data rows
    data_row = 3
    for teacher_disp, period_data in teacher_period_map.items():
        c = ws.cell(row=data_row, column=1, value=teacher_disp)
        c.font=leave_font; c.fill=leave_fill; c.alignment=left_align; c.border=border

        for ci, pc in enumerate(PERIOD_COLS, 2):
            c = ws.cell(row=data_row, column=ci)
            c.alignment=center_align; c.border=border
            if pc in period_data:
                info   = period_data[pc]
                by     = info["arranged_by"]
                c.value = f"{info['class_subject']}\n→ {by}" if by else info["class_subject"]
                c.fill  = arranged_fill if info["saved"] else pending_fill
            else:
                c.value = "-"; c.fill = dash_fill

        ws.row_dimensions[data_row].height = 38
        data_row += 1

    # Column widths
    ws.column_dimensions["A"].width = 28
    for ci in range(2, 2 + len(PERIOD_COLS)):
        ws.column_dimensions[get_column_letter(ci)].width = 20


def build_appended_excel(file_bytes: bytes, sel_date: date, day_code: str,
                         arrangement_rows: list, arr_state: list) -> bytes:
    """Returns original workbook bytes with arrangement sheet appended."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    sheet_name = sel_date.strftime("%d-%m-%Y")
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)
    _write_arrangement_sheet(ws, sel_date, day_code, arrangement_rows, arr_state)
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.read()


def build_standalone_excel(sel_date: date, day_code: str,
                            arrangement_rows: list, arr_state: list) -> bytes:
    """Returns a standalone workbook with only the arrangement sheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sel_date.strftime("%d-%m-%Y")
    _write_arrangement_sheet(ws, sel_date, day_code, arrangement_rows, arr_state)
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.read()


# ═══════════════════════════════════════════════════════════
#  UI — STEP 1: Upload
# ═══════════════════════════════════════════════════════════

st.header("Step 1: Upload Timetable Excel")
uploaded = st.file_uploader("Upload Summary_Report.xlsx", type=["xlsx"])
if not uploaded:
    st.info("Please upload the timetable Excel file to continue.")
    st.stop()

file_bytes = uploaded.read()
teacher_tt, class_tt, unavail, disp2raw, raw2disp = load_data(file_bytes)
display_names = sorted(disp2raw.keys())

has_unav = "Unavailability" in pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names
st.success(
    f"✅ Loaded — Teachers: **{teacher_tt['Teacher'].nunique()}** · "
    f"Classes: **{class_tt['Class'].nunique()}** · "
    f"Unavailability: **{'Yes – ' + str(len(unavail)) + ' entries' if has_unav else 'Sheet not found'}**"
)
st.caption(
    f"Period split → First Half: {FIRST_HALF} ({len(FIRST_HALF)} periods)  |  "
    f"Second Half: {SECOND_HALF} ({len(SECOND_HALF)} periods)"
)

# ═══════════════════════════════════════════════════════════
#  UI — STEP 2: Date + Leave entries
# ═══════════════════════════════════════════════════════════

st.header("Step 2: Select Date & Add Teachers on Leave")

col_d, _ = st.columns([2, 5])
sel_date  = col_d.date_input("Select Date", value=date.today())
weekday   = sel_date.strftime("%A")
day_code  = DAY_MAP.get(weekday)
if not day_code:
    st.warning(f"{weekday} is not a school day.")
    st.stop()
st.write(f"**Day:** {weekday} ({day_code})")

if "leave_list" not in st.session_state:
    st.session_state.leave_list = []

with st.expander("➕ Add Teacher on Leave", expanded=True):
    a, b, c = st.columns([3, 2, 1])
    sel_t = a.selectbox("Teacher", display_names, key="inp_t")
    sel_l = b.selectbox("Leave Type", ["First Half","Second Half","Full Day"], key="inp_l")
    c.write(""); c.write("")
    if c.button("Add"):
        if any(e["display"] == sel_t for e in st.session_state.leave_list):
            st.warning(f"{sel_t} is already in the list.")
        else:
            st.session_state.leave_list.append({
                "display"   : sel_t,
                "raw_keys"  : disp2raw.get(sel_t, [sel_t]),
                "leave_type": sel_l
            })
            st.rerun()

if st.session_state.leave_list:
    st.subheader("Teachers on Leave")
    for idx, e in enumerate(st.session_state.leave_list):
        x, y, z = st.columns([3, 2, 1])
        x.write(f"**{e['display']}**")
        y.write(e["leave_type"])
        if z.button("❌", key=f"del_{idx}"):
            st.session_state.leave_list.pop(idx)
            st.rerun()
else:
    st.stop()

# ═══════════════════════════════════════════════════════════
#  UI — STEP 3: Arrangement Table
# ═══════════════════════════════════════════════════════════

st.divider()
st.header("Step 3: Arrangement Table")

all_on_leave_raw = [rk for e in st.session_state.leave_list for rk in e["raw_keys"]]

# ── Build arrangement rows ─────────────────────────────────
arrangement_rows = []
for entry in st.session_state.leave_list:
    raw_keys   = entry["raw_keys"]
    leave_type = entry["leave_type"]

    tt_row = None
    for rk in raw_keys:
        rows = teacher_tt[(teacher_tt["Teacher"]==rk) & (teacher_tt["Day"]==day_code)]
        if not rows.empty:
            tt_row = rows.iloc[0]; break

    if tt_row is None:
        arrangement_rows.append({
            "teacher": entry["display"], "leave_type": leave_type,
            "period_col": None, "class_subject": "⚠️ No timetable for this day",
            "raw_keys": raw_keys
        })
        continue

    half_cols = get_half_cols(leave_type)
    periods_to_arrange = [pc for pc in half_cols if tt_row.get(pc,"")]

    if not periods_to_arrange:
        arrangement_rows.append({
            "teacher": entry["display"], "leave_type": leave_type,
            "period_col": None, "class_subject": "No classes in this half",
            "raw_keys": raw_keys
        })
        continue

    for pc in periods_to_arrange:
        arrangement_rows.append({
            "teacher"      : entry["display"],
            "leave_type"   : leave_type,
            "period_col"   : pc,
            "class_subject": tt_row.get(pc,""),
            "raw_keys"     : raw_keys
        })

# ── Init per-row state ─────────────────────────────────────
# State keys:
#   method              → "Same Class Teacher" | "Any Teacher" | "Manual Entry"
#   manual              → display name chosen by user (Manual Entry only)
#   _locked_suggested   → the suggested name at the moment Add was clicked (locked-in)
#   saved               → bool
if ("arr_state" not in st.session_state
        or len(st.session_state.arr_state) != len(arrangement_rows)):
    st.session_state.arr_state = [
        {"method":"Same Class Teacher","manual":"","_locked_suggested":None,"saved":False}
        for _ in arrangement_rows
    ]
arr = st.session_state.arr_state

# ── Rebuild extra_busy from SAVED rows only ───────────────
# This is the authoritative dynamic load that persists across rerenders.
extra_busy: dict = {}
for row, state in zip(arrangement_rows, arr):
    if state["saved"] and row["period_col"]:
        method   = state["method"]
        assigned = (state["manual"] if method == "Manual Entry"
                    else state.get("_locked_suggested"))
        if assigned:
            extra_busy.setdefault(assigned, set()).add(row["period_col"])

# ── Table header ───────────────────────────────────────────
h0, h1, h2, h3, h4 = st.columns([2, 2, 2, 4, 1])
h0.markdown("**Teacher / Period**")
h1.markdown("**Class & Subject**")
h2.markdown("**Arrangement**")
h3.markdown("**Remarks – Workload of Assigned Teacher**")
h4.markdown("**Action**")
st.divider()

for i, row in enumerate(arrangement_rows):
    c0, c1, c2, c3, c4 = st.columns([2, 2, 2, 4, 1])

    # ── Col 0: Teacher on leave + period info ──
    with c0:
        st.write(f"**{row['teacher']}**")
        if row["period_col"]:
            st.caption(f"{row['period_col']} · {which_half(row['period_col'])} · {row['leave_type']}")
        else:
            st.caption(row["leave_type"])

    # ── Col 1: Class & Subject ─────────────────
    with c1:
        st.write(row["class_subject"])

    # ── Col 2: Method + assigned teacher name ──
    with c2:
        if row["period_col"] is None:
            st.write("—")

        elif arr[i]["saved"]:
            # Row is locked — show method and assigned teacher clearly
            method   = arr[i]["method"]
            assigned = (arr[i]["manual"] if method == "Manual Entry"
                        else arr[i].get("_locked_suggested") or "—")
            st.write(f"**{method}**")
            st.write(f"👤 {assigned}")     # normal text, not grey caption

        else:
            # Editable row
            method = st.radio(
                "Method",
                ["Same Class Teacher","Any Teacher","Manual Entry"],
                key=f"m_{i}", horizontal=False,
                label_visibility="collapsed",
                index=["Same Class Teacher","Any Teacher","Manual Entry"]
                      .index(arr[i]["method"])
            )
            arr[i]["method"] = method

            if method == "Manual Entry":
                # keep previous manual selection if any
                prev_idx = (display_names.index(arr[i]["manual"])
                            if arr[i]["manual"] in display_names else 0)
                manual = st.selectbox(
                    "Select Teacher", display_names,
                    index=prev_idx,
                    key=f"man_{i}",
                    label_visibility="collapsed"
                )
                arr[i]["manual"] = manual

    # ── Col 3: Suggestion + Remarks ───────────
    with c3:
        if row["period_col"] is None:
            st.write("—")

        elif arr[i]["saved"]:
            # Row is locked — show updated workload of the locked-in teacher
            method   = arr[i]["method"]
            assigned = (arr[i]["manual"] if method == "Manual Entry"
                        else arr[i].get("_locked_suggested"))
            if assigned:
                remark = build_remarks(
                    assigned, day_code, row["leave_type"],
                    disp2raw, teacher_tt, unavail, extra_busy
                )
                st.markdown(remark)
            else:
                st.write("—")

        else:
            # Unsaved row — run suggestion fresh using current extra_busy
            method = arr[i]["method"]

            if method != "Manual Entry":
                suggested, reason = suggest_teacher(
                    method,
                    row["raw_keys"],
                    row["period_col"],
                    day_code,
                    teacher_tt, class_tt, unavail,
                    disp2raw, raw2disp,
                    all_on_leave_raw,
                    extra_busy       # includes all previously saved rows
                )
                if suggested:
                    st.markdown(f"💡 **Suggested:** {suggested}")
                    st.caption(reason)
                    # Store temporarily so Add button can lock it
                    arr[i]["_pending_suggested"] = suggested
                else:
                    st.warning(f"⚠️ {reason}")
                    arr[i]["_pending_suggested"] = None
                info_teacher = suggested
            else:
                info_teacher = arr[i].get("manual","")

            # Show workload of suggested/manual teacher
            if info_teacher:
                remark = build_remarks(
                    info_teacher, day_code, row["leave_type"],
                    disp2raw, teacher_tt, unavail, extra_busy
                )
                st.markdown(remark)

    # ── Col 4: Add / Edit ──────────────────────
    with c4:
        if row["period_col"] is None:
            st.write("")
        elif not arr[i]["saved"]:
            if st.button("✅ Add", key=f"add_{i}"):
                method = arr[i]["method"]
                if method == "Manual Entry":
                    # Lock in manual choice
                    arr[i]["_locked_suggested"] = None
                else:
                    # Lock in whatever was suggested this render
                    arr[i]["_locked_suggested"] = arr[i].get("_pending_suggested")
                arr[i]["saved"] = True
                # Immediately add to extra_busy before rerun
                assigned = (arr[i]["manual"] if method == "Manual Entry"
                            else arr[i]["_locked_suggested"])
                if assigned and row["period_col"]:
                    extra_busy.setdefault(assigned, set()).add(row["period_col"])
                st.rerun()
        else:
            if st.button("✏️ Edit", key=f"edit_{i}"):
                # Un-lock: remove from extra_busy
                method   = arr[i]["method"]
                assigned = (arr[i]["manual"] if method == "Manual Entry"
                            else arr[i].get("_locked_suggested"))
                if assigned and row["period_col"] and assigned in extra_busy:
                    extra_busy[assigned].discard(row["period_col"])
                arr[i]["saved"] = False
                arr[i]["_locked_suggested"] = None
                st.rerun()

    st.divider()

# ═══════════════════════════════════════════════════════════
#  SUMMARY + DOWNLOADS
# ═══════════════════════════════════════════════════════════

saved_pairs = [
    (r, s) for r, s in zip(arrangement_rows, arr)
    if s["saved"] and r["period_col"]
]

if saved_pairs:
    st.subheader("📄 Saved Arrangements Summary")
    out_rows = []
    for r, s in saved_pairs:
        method      = s["method"]
        arranged_by = (s["manual"] if method == "Manual Entry"
                       else s.get("_locked_suggested") or "—")
        out_rows.append({
            "Teacher on Leave": r["teacher"],
            "Period"          : r["period_col"],
            "Half"            : which_half(r["period_col"]),
            "Class & Subject" : r["class_subject"],
            "Leave Type"      : r["leave_type"],
            "Method"          : method,
            "Arranged By"     : arranged_by,
        })
    st.dataframe(pd.DataFrame(out_rows), use_container_width=True, hide_index=True)

    # ── Download section ─────────────────────────
    total_arrangeable = sum(1 for r in arrangement_rows if r["period_col"])
    total_saved       = len(saved_pairs)
    sheet_name        = sel_date.strftime("%d-%m-%Y")

    if total_saved < total_arrangeable:
        st.info(f"ℹ️ {total_arrangeable - total_saved} period(s) still pending. "
                f"Save all rows to enable downloads.")
    else:
        st.success("✅ All periods arranged!")
        st.subheader("📥 Download Options")

        dl1, dl2 = st.columns(2)

        with dl1:
            standalone = build_standalone_excel(
                sel_date, day_code, arrangement_rows, arr
            )
            st.download_button(
                label           = f"📄 Download Arrangement Only ({sheet_name})",
                data            = standalone,
                file_name       = f"Arrangement_{sheet_name}.xlsx",
                mime            = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help            = "Single sheet with today's arrangement"
            )

        with dl2:
            appended = build_appended_excel(
                file_bytes, sel_date, day_code, arrangement_rows, arr
            )
            st.download_button(
                label           = f"📚 Download Full Report (Original + {sheet_name})",
                data            = appended,
                file_name       = f"Summary_Report_{sheet_name}.xlsx",
                mime            = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help            = "Original workbook with arrangement sheet appended"
            )

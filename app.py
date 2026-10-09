from __future__ import annotations

import hashlib
import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

APP_VERSION = "0.1.1"
MIRO_SCORE_COLUMN = 8  # H: Moodle grade
FIRST_WEEK_COLUMN = 10  # J: keep I (last downloaded) intact
SETTINGS_SHEET = "Miro_Settings"
DEFAULT_WEEKS = 14
DEFAULT_WEEK_MAX = 2.0
DEFAULT_FINAL_MAX = 20.0
DEFAULT_DECIMALS = 1


def normalize_number(value):
    if value is None or value == "":
        return None
    try:
        n = float(value)
        if n.is_integer():
            return int(n)
        return n
    except (TypeError, ValueError):
        return value


def id_last4(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = str(value).strip()
    digits = "".join(re.findall(r"\d", s))
    if digits:
        return digits[-4:].zfill(4) if len(digits) < 4 else digits[-4:]
    compact = re.sub(r"\s+", "", s)
    return compact[-4:]


def find_header_row(ws) -> int | None:
    for r in range(1, min(ws.max_row, 12) + 1):
        b = str(ws.cell(r, 2).value or "").strip()
        c = str(ws.cell(r, 3).value or "").strip()
        h = str(ws.cell(r, MIRO_SCORE_COLUMN).value or "").strip()
        if b == "氏名" and "ID" in c.upper() and "Miro" in h:
            return r
    return None


def candidate_sheets(wb):
    result = []
    for ws in wb.worksheets:
        if ws.title == SETTINGS_SHEET:
            continue
        row = find_header_row(ws)
        if row:
            result.append((ws.title, row))
    return result


def ensure_settings_sheet(wb):
    if SETTINGS_SHEET in wb.sheetnames:
        ws = wb[SETTINGS_SHEET]
    else:
        ws = wb.create_sheet(SETTINGS_SHEET)
        ws["A1"] = "Miro Grader settings"
        ws["A2"] = "app_version"
        ws["B2"] = APP_VERSION
        ws["A3"] = "num_weeks"
        ws["B3"] = DEFAULT_WEEKS
        ws["A4"] = "week_max"
        ws["B4"] = DEFAULT_WEEK_MAX
        ws["A5"] = "final_max"
        ws["B5"] = DEFAULT_FINAL_MAX
        ws["A6"] = "decimals"
        ws["B6"] = DEFAULT_DECIMALS
        ws["A7"] = "current_week"
        ws["B7"] = 1
        ws["A8"] = "active_weeks"
        ws["B8"] = ""
    ws.sheet_state = "hidden"
    return ws


def read_settings(wb):
    ws = ensure_settings_sheet(wb)
    settings = {
        "num_weeks": int(ws["B3"].value or DEFAULT_WEEKS),
        "week_max": float(ws["B4"].value or DEFAULT_WEEK_MAX),
        "final_max": float(ws["B5"].value or DEFAULT_FINAL_MAX),
        "decimals": int(ws["B6"].value if ws["B6"].value is not None else DEFAULT_DECIMALS),
        "current_week": int(ws["B7"].value or 1),
        "active_weeks": [],
    }
    active_raw = str(ws["B8"].value or "").strip()
    if active_raw:
        settings["active_weeks"] = sorted({int(x) for x in active_raw.split(",") if x.strip().isdigit()})
    settings["num_weeks"] = max(1, min(40, settings["num_weeks"]))
    settings["current_week"] = max(1, min(settings["num_weeks"], settings["current_week"]))
    settings["active_weeks"] = [w for w in settings["active_weeks"] if 1 <= w <= settings["num_weeks"]]
    return settings


def write_settings(wb, settings):
    ws = ensure_settings_sheet(wb)
    ws["B2"] = APP_VERSION
    ws["B3"] = int(settings["num_weeks"])
    ws["B4"] = float(settings["week_max"])
    ws["B5"] = float(settings["final_max"])
    ws["B6"] = int(settings["decimals"])
    ws["B7"] = int(settings["current_week"])
    ws["B8"] = ",".join(str(x) for x in sorted(settings["active_weeks"]))
    ws.sheet_state = "hidden"


def week_header(week_no: int) -> str:
    return f"Week {week_no}"


def week_column_map(ws, header_row: int):
    mapping = {}
    for col in range(FIRST_WEEK_COLUMN, ws.max_column + 1):
        value = str(ws.cell(header_row, col).value or "").strip()
        m = re.fullmatch(r"Week\s*(\d+)", value, flags=re.I)
        if m:
            mapping[int(m.group(1))] = col
    return mapping


def ensure_week_columns(ws, header_row: int, num_weeks: int):
    mapping = week_column_map(ws, header_row)
    next_col = max(FIRST_WEEK_COLUMN, ws.max_column + 1)
    for week in range(1, num_weeks + 1):
        if week not in mapping:
            col = next_col
            ws.cell(header_row, col).value = week_header(week)
            ws.cell(header_row, col).font = Font(bold=True)
            ws.cell(header_row, col).alignment = Alignment(horizontal="center")
            ws.cell(header_row, col).fill = PatternFill(fill_type="solid", fgColor="E8F0FE")
            mapping[week] = col
            next_col += 1
    return mapping


def student_rows(ws, header_row: int):
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(r, 2).value
        student_id = ws.cell(r, 3).value
        if (name is None or str(name).strip() == "") and (student_id is None or str(student_id).strip() == ""):
            continue
        rows.append(
            {
                "row": r,
                "id4": id_last4(student_id),
                "name": str(name or "").strip(),
                "id_full": student_id,
            }
        )
    return rows


def compute_raw_total(ws, row: int, week_cols: dict[int, int], active_weeks: list[int]) -> float:
    total = 0.0
    for week in active_weeks:
        col = week_cols.get(week)
        if not col:
            continue
        value = ws.cell(row, col).value
        try:
            if value is not None and value != "":
                total += float(value)
        except (TypeError, ValueError):
            pass
    return total


def recalc_final_scores(ws, header_row: int, settings, week_cols):
    active_weeks = settings["active_weeks"]
    week_max = float(settings["week_max"])
    final_max = float(settings["final_max"])
    decimals = int(settings["decimals"])
    possible = len(active_weeks) * week_max

    for stu in student_rows(ws, header_row):
        row = stu["row"]
        if possible <= 0:
            ws.cell(row, MIRO_SCORE_COLUMN).value = None
            continue
        raw = compute_raw_total(ws, row, week_cols, active_weeks)
        final = raw / possible * final_max
        final = max(0.0, min(final_max, final))
        rounded = round(final, decimals)
        ws.cell(row, MIRO_SCORE_COLUMN).value = int(rounded) if decimals == 0 else rounded


def workbook_bytes(wb) -> bytes:
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def build_table(ws, header_row: int, settings, week_cols, current_week: int):
    data = []
    active = settings["active_weeks"]
    possible = len(active) * settings["week_max"]
    current_col = week_cols[current_week]
    for stu in student_rows(ws, header_row):
        row = stu["row"]
        raw = compute_raw_total(ws, row, week_cols, active)
        final = None
        if possible > 0:
            final = round(raw / possible * settings["final_max"], settings["decimals"])
        data.append(
            {
                "下4桁": stu["id4"],
                "氏名": stu["name"],
                week_header(current_week): normalize_number(ws.cell(row, current_col).value),
                "素点合計": round(raw, 2),
                f"最終 / {settings['final_max']:g}": final,
            }
        )
    return pd.DataFrame(data)


st.set_page_config(page_title="Miro Grader", page_icon="🧩", layout="wide")
st.title("🧩 Miro Group Work Grader")
st.caption("ExcelのJ列以降にWeek別点数を保存し、H列へMiroの最終評定を自動反映します。I列は保持します。")

uploaded = st.file_uploader("評定Excel（.xlsx）を選択", type=["xlsx"])
if uploaded is None:
    st.info("まず評定Excelを選択してください。元ファイルは上書きせず、更新版をダウンロードできます。")
    st.stop()

file_data = uploaded.getvalue()
file_hash = hashlib.sha256(file_data).hexdigest()
if st.session_state.get("loaded_hash") != file_hash:
    try:
        wb = load_workbook(io.BytesIO(file_data))
    except Exception as e:
        st.error(f"Excelを開けませんでした: {e}")
        st.stop()
    candidates = candidate_sheets(wb)
    if not candidates:
        st.error("B列「氏名」、C列「IDナンバ」、H列「グループワーク (Miro)」の見出しを持つシートが見つかりませんでした。")
        st.stop()
    st.session_state.loaded_hash = file_hash
    st.session_state.wb = wb
    st.session_state.target_sheet = candidates[0][0]
    st.session_state.header_row = candidates[0][1]
    st.session_state.original_name = uploaded.name
    st.session_state.flash = "Excelを読み込みました。"

wb = st.session_state.wb
available_sheets = [ws.title for ws in wb.worksheets if ws.title != SETTINGS_SHEET]
selected_sheet = st.sidebar.selectbox(
    "評定シート",
    available_sheets,
    index=available_sheets.index(st.session_state.target_sheet) if st.session_state.target_sheet in available_sheets else 0,
)
if selected_sheet != st.session_state.target_sheet:
    ws_test = wb[selected_sheet]
    detected = find_header_row(ws_test)
    if not detected:
        st.sidebar.error("このシートでは見出し行を自動認識できません。")
        st.stop()
    st.session_state.target_sheet = selected_sheet
    st.session_state.header_row = detected

ws = wb[st.session_state.target_sheet]
header_row = st.session_state.header_row
settings = read_settings(wb)

st.sidebar.subheader("設定")
with st.sidebar.form("settings_form"):
    num_weeks = st.number_input("Week数", min_value=1, max_value=40, value=int(settings["num_weeks"]), step=1)
    week_max = st.number_input("1回あたりの満点", min_value=0.1, max_value=100.0, value=float(settings["week_max"]), step=0.5)
    final_max = st.number_input("Miro最終配点（/100）", min_value=0.1, max_value=100.0, value=float(settings["final_max"]), step=1.0)
    decimals = st.number_input("最終点の小数桁", min_value=0, max_value=3, value=int(settings["decimals"]), step=1)
    all_weeks_for_form = list(range(1, int(num_weeks) + 1))
    active_default = [w for w in settings["active_weeks"] if w in all_weeks_for_form]
    active_weeks = st.multiselect(
        "採点対象Week（飛ばした週は外す）",
        all_weeks_for_form,
        default=active_default,
        format_func=lambda x: f"Week {x}",
    )
    save_settings = st.form_submit_button("設定を保存", use_container_width=True)

if save_settings:
    settings["num_weeks"] = int(num_weeks)
    settings["week_max"] = float(week_max)
    settings["final_max"] = float(final_max)
    settings["decimals"] = int(decimals)
    settings["active_weeks"] = sorted(active_weeks)
    settings["current_week"] = min(settings["current_week"], settings["num_weeks"])
    week_cols = ensure_week_columns(ws, header_row, settings["num_weeks"])
    recalc_final_scores(ws, header_row, settings, week_cols)
    write_settings(wb, settings)
    st.session_state.flash = "設定を保存し、H列を再計算しました。"
    st.rerun()

# Ensure columns exist based on current settings.
week_cols = ensure_week_columns(ws, header_row, settings["num_weeks"])
roster = student_rows(ws, header_row)

if st.session_state.get("flash"):
    st.success(st.session_state.pop("flash"))

c1, c2, c3, c4 = st.columns(4)
c1.metric("学生数", len(roster))
c2.metric("1回満点", f"{settings['week_max']:g}")
c3.metric("採点対象Week", len(settings["active_weeks"]))
c4.metric("最終配点", f"{settings['final_max']:g} / 100")

week_options = list(range(1, settings["num_weeks"] + 1))
default_index = max(0, min(len(week_options) - 1, settings["current_week"] - 1))
current_week = st.selectbox("今回のWeek", week_options, index=default_index, format_func=lambda x: f"Week {x}")
settings["current_week"] = current_week
write_settings(wb, settings)

current_col = week_cols[current_week]
entered_count = sum(1 for stu in roster if ws.cell(stu["row"], current_col).value not in (None, ""))
st.progress(entered_count / len(roster) if roster else 0, text=f"Week {current_week}: {entered_count} / {len(roster)} 人入力済み")

st.subheader(f"Week {current_week} の採点")
search = st.text_input("学生を絞り込み（氏名または下4桁）", placeholder="例: 1234 / 山田")
filtered = roster
if search.strip():
    q = search.strip().lower()
    filtered = [s for s in roster if q in s["id4"].lower() or q in s["name"].lower()]

option_map = {f"{s['id4']} | {s['name']}": s for s in filtered}
selected_labels = st.multiselect(
    "このグループの学生を選択",
    list(option_map.keys()),
    placeholder="名前または下4桁を入力すると検索できます",
)
selected_students = [option_map[x] for x in selected_labels]

score_col, action_col = st.columns([1, 2])
with score_col:
    score = st.number_input(
        "点数",
        min_value=0.0,
        max_value=float(settings["week_max"]),
        value=float(settings["week_max"]),
        step=0.5 if settings["week_max"] % 1 else 1.0,
    )
with action_col:
    st.write(" ")
    st.write(" ")
    apply_score = st.button(
        f"選択した {len(selected_students)} 人に {score:g} 点を入力",
        type="primary",
        disabled=not selected_students,
        use_container_width=True,
    )

# Quick score buttons for small integer scales.
if float(settings["week_max"]).is_integer() and int(settings["week_max"]) <= 5:
    quick_values = list(range(0, int(settings["week_max"]) + 1))
    quick_cols = st.columns(len(quick_values))
    quick_clicked = None
    for idx, value in enumerate(quick_values):
        if quick_cols[idx].button(f"{value}点", disabled=not selected_students, use_container_width=True, key=f"q_{value}"):
            quick_clicked = value
    if quick_clicked is not None:
        score = float(quick_clicked)
        apply_score = True

if apply_score:
    for stu in selected_students:
        ws.cell(stu["row"], current_col).value = normalize_number(score)
    if current_week not in settings["active_weeks"]:
        settings["active_weeks"].append(current_week)
        settings["active_weeks"] = sorted(set(settings["active_weeks"]))
    settings["current_week"] = current_week
    recalc_final_scores(ws, header_row, settings, week_cols)
    write_settings(wb, settings)
    st.session_state.flash = f"Week {current_week}: {len(selected_students)}人に {score:g} 点を入力しました。"
    st.rerun()

clear_score = st.button("選択した学生のこのWeekの点数を空欄に戻す", disabled=not selected_students)
if clear_score:
    for stu in selected_students:
        ws.cell(stu["row"], current_col).value = None
    recalc_final_scores(ws, header_row, settings, week_cols)
    write_settings(wb, settings)
    st.session_state.flash = f"Week {current_week}: {len(selected_students)}人の点数を空欄に戻しました。"
    st.rerun()

st.divider()
st.subheader("確認")
df = build_table(ws, header_row, settings, week_cols, current_week)
st.dataframe(df, use_container_width=True, hide_index=True, height=min(700, 80 + 35 * max(1, len(df))))

if settings["active_weeks"]:
    st.caption(
        "H列は、採点対象Weekの素点合計 ÷ 採点対象Weekの満点合計 × Miro最終配点、で計算します。"
        " 採点対象Week内の空欄は、その学生について0点相当になります。"
    )
else:
    st.warning("採点対象Weekがまだありません。最初の採点を入力すると、そのWeekが自動で採点対象になります。")

# Recalculate once more before export.
recalc_final_scores(ws, header_row, settings, week_cols)
write_settings(wb, settings)

original = Path(st.session_state.original_name)
out_name = f"{original.stem}_Miro更新{original.suffix}"
export_data = workbook_bytes(wb)
st.download_button(
    "更新済みExcelをダウンロード",
    data=export_data,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)
st.caption("元ファイルは上書きしません。次回は、ここからダウンロードした更新版Excelを読み込んで続けてください。")

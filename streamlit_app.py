from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter


COURSE_HEADERS = [
    "#", "課程代碼", "學年度", "醫院別", "開課單位", "課程名稱", "課程類型", "排除重複課程",
    "上課時間", "開始日期", "結束日期", "上課時間", "上課分鐘數", "表單發送時間", "課程分類",
    "課程備註", "上課地點", "該課程授課老師為住院醫師", "授課老師", "授課老師員編", "協同老師",
    "協同老師員編", "該課程協同&授課老師皆為同一位住院醫師", "學生", "輔助教材", "職類", "計畫類別",
    "訓練計畫", "訓練科室", "符合/不符合",
]

SUMMARY_HEADERS = COURSE_HEADERS[:-1] + ["A.住院醫師主授課程/B.住院醫師擔任協同老師"]


@dataclass(frozen=True)
class ColumnSettings:
    main_teacher_name: str = "S"
    main_teacher_status: str = "T"
    co_teacher_name: str = "U"
    co_teacher_status: str = "V"
    student: str = "X"


def column_index(letter: str) -> int:
    cleaned = str(letter).strip().upper()
    if not re.fullmatch(r"[A-Z]+", cleaned):
        raise ValueError(f"欄位代號不正確：{letter}")

    value = 0
    for char in cleaned:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value - 1


def ensure_column_count(df: pd.DataFrame, min_count: int) -> pd.DataFrame:
    while len(df.columns) < min_count:
        df[f"__空白欄位_{len(df.columns) + 1}__"] = ""
    return df


def normalize_name(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", "", str(value).strip())


def normalize_value(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def split_teacher_names(value: object) -> list[str]:
    text = normalize_value(value)
    if not text:
        return []
    parts = re.split(r"[、,，;；/／\\\n\r]+", text)
    return [normalize_name(part) for part in parts if normalize_name(part)]


def first_matching_column(columns: Iterable[object], keywords: tuple[str, ...]) -> object | None:
    normalized = [(column, normalize_name(column)) for column in columns]

    for keyword in keywords:
        key = normalize_name(keyword)
        for column, column_name in normalized:
            if column_name == key:
                return column

    for keyword in keywords:
        key = normalize_name(keyword)
        for column, column_name in normalized:
            if key in column_name:
                return column

    return None


def find_col(df: pd.DataFrame, keywords: tuple[str, ...], fallback_letter: str | None = None) -> object:
    col = first_matching_column(df.columns, keywords)
    if col is not None:
        return col
    if fallback_letter:
        idx = column_index(fallback_letter)
        ensure_column_count(df, idx + 1)
        return df.columns[idx]
    raise ValueError(f"找不到欄位：{keywords}")


def build_teacher_lookup(mapping_df: pd.DataFrame) -> dict[str, str]:
    name_col = first_matching_column(
        mapping_df.columns,
        ("教師姓名", "老師姓名", "住院醫師姓名", "醫師姓名", "姓名", "授課老師", "協同老師"),
    )
    employee_col = first_matching_column(
        mapping_df.columns,
        ("員編", "員工編號", "教師員編", "醫師員編", "住院醫師員編", "職編", "ID"),
    )

    if name_col is None or employee_col is None:
        raise ValueError("比對用住院醫師名單需包含姓名欄位與員編欄位，例如「姓名」與「員編」。")

    lookup: dict[str, str] = {}
    for _, row in mapping_df.iterrows():
        name = normalize_name(row[name_col])
        employee_id = normalize_value(row[employee_col])
        if name and employee_id:
            lookup[name] = employee_id
    return lookup


def teacher_ids_for_value(value: object, lookup: dict[str, str]) -> str:
    ids = [lookup[name] for name in split_teacher_names(value) if name in lookup]
    seen: list[str] = []
    for employee_id in ids:
        if employee_id not in seen:
            seen.append(employee_id)
    return "、".join(seen)


def any_teacher_is_resident(value: object, lookup: dict[str, str]) -> bool:
    return any(name in lookup for name in split_teacher_names(value))


def same_resident_teacher(main_value: object, co_value: object, lookup: dict[str, str]) -> bool:
    main_names = {name for name in split_teacher_names(main_value) if name in lookup}
    co_names = {name for name in split_teacher_names(co_value) if name in lookup}
    return bool(main_names and co_names and main_names.intersection(co_names))


def read_excel_first_sheet(file) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=0, dtype=object, engine="openpyxl").astype(object)


def make_dedupe_key(df: pd.DataFrame) -> pd.Series:
    ensure_column_count(df, column_index("S") + 1)
    i_idx = column_index("I")
    q_idx = column_index("Q")
    s_idx = column_index("S")
    return (
        df.iloc[:, i_idx].map(normalize_value)
        + "|"
        + df.iloc[:, q_idx].map(normalize_value)
        + "|"
        + df.iloc[:, s_idx].map(normalize_value)
    )


def standardize_report(report_df: pd.DataFrame, lookup: dict[str, str], settings: ColumnSettings) -> pd.DataFrame:
    """建立課程總表：完整保留原始每一筆，不刪除、不去重、不篩選。"""
    df = report_df.copy().astype(object)
    ensure_column_count(df, max(column_index("AD"), column_index(settings.student)) + 1)

    main_col = find_col(df, ("授課老師", "授課教師", "主授課老師"), settings.main_teacher_name)
    co_col = find_col(df, ("協同老師", "協同教師", "協同授課老師"), settings.co_teacher_name)
    student_col = find_col(df, ("學生", "學員", "受訓學員", "受訓人員"), settings.student)

    # 依指定輸出格式建置，缺少的欄位補空白。重複欄名「上課時間」採原始欄位順序依序帶入。
    original_cols = list(df.columns)
    course_time_cols = [c for c in original_cols if normalize_name(c) == "上課時間"]
    output = pd.DataFrame(index=df.index)

    for header in COURSE_HEADERS:
        if header == "#":
            output[header] = range(1, len(df) + 1)
        elif header == "排除重複課程":
            output[header] = make_dedupe_key(df)
        elif header == "該課程授課老師為住院醫師":
            output[header] = df[main_col].map(lambda v: "授課老師有員編" if any_teacher_is_resident(v, lookup) else "")
        elif header == "授課老師":
            output[header] = df[main_col].map(normalize_value)
        elif header == "授課老師員編":
            output[header] = df[main_col].map(lambda v: teacher_ids_for_value(v, lookup))
        elif header == "協同老師":
            output[header] = df[co_col].map(normalize_value)
        elif header == "協同老師員編":
            output[header] = df[co_col].map(lambda v: teacher_ids_for_value(v, lookup))
        elif header == "該課程協同&授課老師皆為同一位住院醫師":
            output[header] = ["是" if same_resident_teacher(m, c, lookup) else "" for m, c in zip(df[main_col], df[co_col])]
        elif header == "學生":
            output[header] = df[student_col].map(normalize_value)
        elif header == "符合/不符合":
            output[header] = ""  # 下方統一計算
        elif header == "上課時間":
            # pandas 不允許以相同欄名重複建立兩欄，先暫用內部名稱，最後再改回指定欄名。
            count = sum(str(c).startswith("上課時間") for c in output.columns)
            source = course_time_cols[count] if count < len(course_time_cols) else first_matching_column(df.columns, ("上課時間",))
            output[f"上課時間__{count + 1}"] = df[source].map(normalize_value) if source is not None else ""
        else:
            source = first_matching_column(df.columns, (header,))
            output[header] = df[source].map(normalize_value) if source is not None else ""

    # 若學生空白，即使授課老師或協同老師是住院醫師，也不列入計算。
    student_not_blank = output["學生"].map(normalize_value).ne("")
    main_is_resident = output["授課老師員編"].map(normalize_value).ne("")
    co_is_resident = output["協同老師員編"].map(normalize_value).ne("")

    output["符合/不符合"] = "不符合"
    output.loc[student_not_blank & (main_is_resident | co_is_resident), "符合/不符合"] = "符合"
    output.loc[~student_not_blank & (main_is_resident | co_is_resident), "符合/不符合"] = "不符合(學生空白)"
    output.loc[student_not_blank & ~(main_is_resident | co_is_resident), "符合/不符合"] = "不符合(非住院醫師授課/協同)"

    # 還原重複欄名：兩個「上課時間」。
    output.columns = COURSE_HEADERS
    return output.astype(object)


def build_summary(course_df: pd.DataFrame) -> pd.DataFrame:
    """彙整只列入符合者，且學生不可空白。A 類優先於 B 類。"""
    df = course_df.copy().astype(object)
    student_not_blank = df["學生"].map(normalize_value).ne("")
    eligible = df["符合/不符合"].map(normalize_value).eq("符合") & student_not_blank

    main_mask = eligible & df["授課老師員編"].map(normalize_value).ne("")
    co_mask = eligible & ~main_mask & df["協同老師員編"].map(normalize_value).ne("")

    summary = df.loc[main_mask | co_mask, COURSE_HEADERS[:-1]].copy()
    category = pd.Series("", index=df.index, dtype=object)
    category.loc[main_mask] = "A.住院醫師主授課程"
    category.loc[co_mask] = "B.住院醫師擔任協同老師"
    summary["A.住院醫師主授課程/B.住院醫師擔任協同老師"] = category.loc[summary.index].values
    summary["#"] = range(1, len(summary) + 1)
    summary.columns = SUMMARY_HEADERS
    return summary.reset_index(drop=True).astype(object)


def autosize_and_style(ws) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAD3")
    header_font = Font(bold=True)
    thin = Border()

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            text = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(len(text), 40))
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col_idx)].width = max(10, min(max_len + 2, 35))


def write_dataframe(ws, df: pd.DataFrame) -> None:
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(["" if pd.isna(value) else value for value in row])
    autosize_and_style(ws)


def process_report(report_file, mapping_file) -> tuple[io.BytesIO, int, int, int]:
    report_df = read_excel_first_sheet(report_file)
    mapping_df = read_excel_first_sheet(mapping_file)
    lookup = build_teacher_lookup(mapping_df)

    course_df = standardize_report(report_df, lookup, ColumnSettings())
    summary_df = build_summary(course_df)

    wb = Workbook()
    ws_course = wb.active
    ws_course.title = "課程總表"
    write_dataframe(ws_course, course_df)

    ws_summary = wb.create_sheet("彙整")
    write_dataframe(ws_summary, summary_df)

    ws_lookup = wb.create_sheet("比對用")
    lookup_df = pd.DataFrame(
        [{"姓名": name, "員編": employee_id} for name, employee_id in lookup.items()]
    )
    write_dataframe(ws_lookup, lookup_df)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(course_df), len(summary_df), int((course_df["符合/不符合"] == "不符合(學生空白)").sum())


st.set_page_config(page_title="RAST 自動化報表", page_icon="📊", layout="centered")
st.title("RAST 自動化報表")
st.caption("完整保留課程總表原始資料；彙整僅計入住院醫師授課/協同且學生欄不空白者。")

report_file = st.file_uploader("上傳原始 RAST Excel 報表", type=["xlsx"])
mapping_file = st.file_uploader("上傳比對用住院醫師名單", type=["xlsx"])

st.info("新規則：若授課老師或協同老師為住院醫師，但學生欄位空白，會標示為「不符合(學生空白)」，不納入彙整計算。")

if st.button("產生整理後 RAST 報表", type="primary"):
    if report_file is None or mapping_file is None:
        st.error("請同時上傳原始 RAST Excel 報表與比對用住院醫師名單。")
    else:
        try:
            output, course_rows, summary_rows, blank_student_rows = process_report(report_file, mapping_file)
            st.success(
                f"完成：課程總表 {course_rows} 筆；彙整 {summary_rows} 筆；學生空白未納入 {blank_student_rows} 筆。"
            )
            st.download_button(
                label="下載整理後 RAST 報表",
                data=output,
                file_name=f"整理後RAST報表_課程總表{course_rows}筆_彙整{summary_rows}筆.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"處理失敗：{exc}")

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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
    co_teacher_name: str = "U"
    student: str = "X"


def column_index(letter: str) -> int:
    cleaned = str(letter).strip().upper()
    if not re.fullmatch(r"[A-Z]+", cleaned):
        raise ValueError(f"欄位代號不正確：{letter}")
    value = 0
    for char in cleaned:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value - 1


def normalize_name(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    # pandas 讀取重複欄名時可能變成「上課時間.1」，比對欄位時移除尾端 .數字
    text = re.sub(r"\.\d+$", "", text)
    return re.sub(r"\s+", "", text)


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


def ensure_column_count(df: pd.DataFrame, min_count: int) -> None:
    while df.shape[1] < min_count:
        df[f"__空白欄位_{df.shape[1] + 1}__"] = ""


def matching_column_indices(columns: Iterable[object], keywords: tuple[str, ...]) -> list[int]:
    normalized_columns = [normalize_name(column) for column in columns]
    normalized_keywords = [normalize_name(keyword) for keyword in keywords]

    exact: list[int] = []
    for i, column_name in enumerate(normalized_columns):
        if column_name in normalized_keywords:
            exact.append(i)
    if exact:
        return exact

    contains: list[int] = []
    for i, column_name in enumerate(normalized_columns):
        if any(key and key in column_name for key in normalized_keywords):
            contains.append(i)
    return contains


def first_matching_index(df: pd.DataFrame, keywords: tuple[str, ...], fallback_letter: str | None = None) -> int:
    indices = matching_column_indices(df.columns, keywords)
    if indices:
        return indices[0]
    if fallback_letter:
        idx = column_index(fallback_letter)
        ensure_column_count(df, idx + 1)
        return idx
    raise ValueError(f"找不到欄位：{'、'.join(keywords)}")


def series_by_index(df: pd.DataFrame, idx: int) -> pd.Series:
    ensure_column_count(df, idx + 1)
    return df.iloc[:, idx].map(normalize_value)


def series_by_keywords(
    df: pd.DataFrame,
    keywords: tuple[str, ...],
    fallback_letter: str | None = None,
    occurrence: int = 0,
) -> pd.Series:
    indices = matching_column_indices(df.columns, keywords)
    if indices and occurrence < len(indices):
        return series_by_index(df, indices[occurrence])
    if indices:
        return series_by_index(df, indices[0])
    if fallback_letter:
        return series_by_index(df, column_index(fallback_letter))
    return pd.Series([""] * len(df), index=df.index, dtype=object)


def build_teacher_lookup(mapping_df: pd.DataFrame) -> dict[str, str]:
    name_idx = first_matching_index(
        mapping_df,
        ("教師姓名", "老師姓名", "住院醫師姓名", "醫師姓名", "姓名", "授課老師", "協同老師"),
    )
    employee_idx = first_matching_index(
        mapping_df,
        ("員編", "員工編號", "教師員編", "醫師員編", "住院醫師員編", "職編", "ID"),
    )

    lookup: dict[str, str] = {}
    for _, row in mapping_df.iterrows():
        name = normalize_name(row.iloc[name_idx])
        employee_id = normalize_value(row.iloc[employee_idx])
        if name and employee_id:
            lookup[name] = employee_id
    return lookup


def teacher_ids_for_value(value: object, lookup: dict[str, str]) -> str:
    ids = [lookup[name] for name in split_teacher_names(value) if name in lookup]
    unique_ids: list[str] = []
    for employee_id in ids:
        if employee_id not in unique_ids:
            unique_ids.append(employee_id)
    return "、".join(unique_ids)


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
    return (
        series_by_index(df, column_index("I"))
        + "|"
        + series_by_index(df, column_index("Q"))
        + "|"
        + series_by_index(df, column_index("S"))
    )


def standardize_report(report_df: pd.DataFrame, lookup: dict[str, str], settings: ColumnSettings) -> pd.DataFrame:
    """建立課程總表：保留每一筆原始資料，不刪除、不去重、不篩選。"""
    df = report_df.copy().astype(object)
    ensure_column_count(df, max(column_index("AD"), column_index(settings.student)) + 1)

    main_idx = first_matching_index(df, ("授課老師", "授課教師", "主授課老師"), settings.main_teacher_name)
    co_idx = first_matching_index(df, ("協同老師", "協同教師", "協同授課老師"), settings.co_teacher_name)
    student_idx = first_matching_index(df, ("學生", "學員", "受訓學員", "受訓人員"), settings.student)

    main_series = series_by_index(df, main_idx)
    co_series = series_by_index(df, co_idx)
    student_series = series_by_index(df, student_idx)

    output_series: list[pd.Series] = []
    course_time_count = 0

    for header in COURSE_HEADERS:
        if header == "#":
            s = pd.Series(range(1, len(df) + 1), index=df.index, dtype=object)
        elif header == "排除重複課程":
            s = make_dedupe_key(df)
        elif header == "該課程授課老師為住院醫師":
            s = main_series.map(lambda v: "授課老師有員編" if any_teacher_is_resident(v, lookup) else "")
        elif header == "授課老師":
            s = main_series
        elif header == "授課老師員編":
            s = main_series.map(lambda v: teacher_ids_for_value(v, lookup))
        elif header == "協同老師":
            s = co_series
        elif header == "協同老師員編":
            s = co_series.map(lambda v: teacher_ids_for_value(v, lookup))
        elif header == "該課程協同&授課老師皆為同一位住院醫師":
            s = pd.Series(
                ["是" if same_resident_teacher(m, c, lookup) else "" for m, c in zip(main_series, co_series)],
                index=df.index,
                dtype=object,
            )
        elif header == "學生":
            s = student_series
        elif header == "符合/不符合":
            s = pd.Series([""] * len(df), index=df.index, dtype=object)
        elif header == "上課時間":
            s = series_by_keywords(df, ("上課時間",), occurrence=course_time_count)
            course_time_count += 1
        else:
            s = series_by_keywords(df, (header,))
        output_series.append(s.astype(object).reset_index(drop=True))

    if len(output_series) != len(COURSE_HEADERS):
        raise RuntimeError(f"輸出欄位數異常：{len(output_series)}，應為 {len(COURSE_HEADERS)}")

    output = pd.concat(output_series, axis=1)
    output.columns = COURSE_HEADERS

    student_not_blank = output["學生"].map(normalize_value).ne("")
    main_is_resident = output["授課老師員編"].map(normalize_value).ne("")
    co_is_resident = output["協同老師員編"].map(normalize_value).ne("")

    output["符合/不符合"] = "不符合"
    output.loc[student_not_blank & (main_is_resident | co_is_resident), "符合/不符合"] = "符合"
    output.loc[~student_not_blank & (main_is_resident | co_is_resident), "符合/不符合"] = "不符合(學生空白)"
    output.loc[student_not_blank & ~(main_is_resident | co_is_resident), "符合/不符合"] = "不符合(非住院醫師授課/協同)"

    return output.astype(object)


def build_summary(course_df: pd.DataFrame) -> pd.DataFrame:
    """彙整只列入符合者，且學生不可空白。A 類優先於 B 類。"""
    df = course_df.copy().astype(object)
    eligible = df["符合/不符合"].map(normalize_value).eq("符合") & df["學生"].map(normalize_value).ne("")

    main_mask = eligible & df["授課老師員編"].map(normalize_value).ne("")
    co_mask = eligible & ~main_mask & df["協同老師員編"].map(normalize_value).ne("")

    selected_mask = main_mask | co_mask
    summary_base = df.loc[selected_mask].iloc[:, : len(COURSE_HEADERS) - 1].copy()
    summary_base.columns = SUMMARY_HEADERS[:-1]

    category = pd.Series("", index=df.index, dtype=object)
    category.loc[main_mask] = "A.住院醫師主授課程"
    category.loc[co_mask] = "B.住院醫師擔任協同老師"
    summary_base["A.住院醫師主授課程/B.住院醫師擔任協同老師"] = category.loc[summary_base.index].values
    summary_base["#"] = range(1, len(summary_base) + 1)

    return summary_base.reset_index(drop=True).astype(object)


def autosize_and_style(ws) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAD3")
    header_font = Font(bold=True)
    thin_side = Side(style="thin", color="D9D9D9")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in col_cells:
            text = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(len(text), 40))
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = border
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
    lookup_df = pd.DataFrame([{"姓名": name, "員編": employee_id} for name, employee_id in lookup.items()])
    write_dataframe(ws_lookup, lookup_df)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    blank_student_rows = int((course_df["符合/不符合"] == "不符合(學生空白)").sum())
    return output, len(course_df), len(summary_df), blank_student_rows


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

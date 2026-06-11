from __future__ import annotations

import io
import re
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


TEMPLATE_HEADERS = [
    "#", "課程代碼", "學年度", "醫院別", "開課單位", "課程名稱", "課程類型", "排除重複課程",
    "上課時間", "開始日期", "結束日期", "上課時間", "上課分鐘數", "表單發送時間", "課程分類", "課程備註",
    "上課地點", "該課程授課老師為住院醫師", "授課老師", "授課老師員編", "協同老師", "協同老師員編",
    "該課程協同&授課老師皆為同一位住院醫師", "學生", "輔助教材", "職類", "計畫類別", "訓練計畫", "訓練科室", None,
]

DEFAULT_WIDTHS = {
    1: 20, 25: 50, 26: 20,
}


@dataclass(frozen=True)
class ColumnSettings:
    main_teacher_name: str = "S"
    co_teacher_name: str = "U"

    def normalized(self) -> "ColumnSettings":
        return ColumnSettings(
            main_teacher_name=self.main_teacher_name.strip().upper(),
            co_teacher_name=self.co_teacher_name.strip().upper(),
        )


def column_index(letter: str) -> int:
    value = 0
    cleaned = str(letter).strip().upper()
    if not re.fullmatch(r"[A-Z]+", cleaned):
        raise ValueError(f"欄位代號不正確：{letter}")
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
    if text.lower() == "nan":
        return ""
    # 避免 Excel 員編 223004.0 這種浮點尾巴
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def split_teacher_names(value: object) -> list[str]:
    text = normalize_value(value)
    if not text:
        return []
    parts = re.split(r"[、,，;；/／\n\r]+", text)
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
            if key and key in column_name:
                return column
    return None


def find_column_index_by_keywords(df: pd.DataFrame, keywords: tuple[str, ...]) -> int | None:
    normalized_columns = [normalize_name(column) for column in df.columns]
    for keyword in keywords:
        key = normalize_name(keyword)
        for index, column_name in enumerate(normalized_columns):
            if column_name == key:
                return index
    for keyword in keywords:
        key = normalize_name(keyword)
        for index, column_name in enumerate(normalized_columns):
            if key and key in column_name:
                return index
    return None


def build_teacher_lookup(mapping_df: pd.DataFrame) -> dict[str, str]:
    name_col = first_matching_column(mapping_df.columns, ("中文姓名", "教師姓名", "老師姓名", "姓名", "授課老師", "協同老師"))
    employee_col = first_matching_column(mapping_df.columns, ("員工編號", "員編", "教師員編", "醫師員編", "住院醫師員編", "職編", "ID"))
    if name_col is None or employee_col is None:
        raise ValueError("比對用住院醫師名單需包含姓名欄位與員編欄位，例如「中文姓名」與「員工編號」。")

    lookup: dict[str, str] = {}
    for _, row in mapping_df.iterrows():
        name = normalize_name(row[name_col])
        employee_id = normalize_value(row[employee_col])
        if name and employee_id:
            lookup[name] = employee_id
    return lookup


def ids_for_names(value: object, lookup: dict[str, str]) -> str:
    ids = []
    for name in split_teacher_names(value):
        emp_id = lookup.get(name)
        if emp_id and emp_id not in ids:
            ids.append(emp_id)
    return "、".join(ids)


def unique_column_name(df: pd.DataFrame, base_name: str) -> str:
    if base_name not in df.columns:
        return base_name
    suffix = 2
    while f"{base_name}_{suffix}" in df.columns:
        suffix += 1
    return f"{base_name}_{suffix}"


def split_multi_co_teacher_rows(df: pd.DataFrame, lookup: dict[str, str], settings: ColumnSettings) -> pd.DataFrame:
    """
    保留相容用函式。

    v4 起「課程總表」不再拆列、不篩選、不刪除任何原始資料列，
    因此這個函式直接回傳原資料。若未來需要另做「拆分協同老師明細表」，
    建議新增獨立工作表，不要改動課程總表。
    """
    return df.copy().reset_index(drop=True)


def get_series_by_template_header(df: pd.DataFrame, header: str | None, used: set[int]) -> pd.Series:
    if header is None:
        return pd.Series([""] * len(df), dtype=object)

    normalized_header = normalize_name(header)
    for idx, col in enumerate(df.columns):
        if idx not in used and normalize_name(col) == normalized_header:
            used.add(idx)
            return df.iloc[:, idx]

    # Excel 讀取重複欄名時，第 2 個「上課時間」常變成「上課時間.1」
    for idx, col in enumerate(df.columns):
        clean_col = re.sub(r"\.\d+$", "", normalize_name(col))
        if idx not in used and clean_col == normalized_header:
            used.add(idx)
            return df.iloc[:, idx]

    return pd.Series([""] * len(df), dtype=object)


def build_course_total(report_df: pd.DataFrame, lookup: dict[str, str], settings: ColumnSettings) -> pd.DataFrame:
    settings = settings.normalized()
    ensure_column_count(report_df, column_index("AD") + 1)

    i_idx = column_index("I")
    q_idx = column_index("Q")
    s_idx = column_index("S")
    dedupe_col = unique_column_name(report_df, "排除重複課程")
    # 只建立與 Excel 公式 I&Q&S 對應的判斷 key，不刪除任何重複列。
    # 使用分隔符避免不同欄位組合產生相同字串，例如 AB+C 與 A+BC。
    report_df[dedupe_col] = (
        report_df.iloc[:, i_idx].map(normalize_value) + "|" +
        report_df.iloc[:, q_idx].map(normalize_value) + "|" +
        report_df.iloc[:, s_idx].map(normalize_value)
    )
    report_df = report_df.reset_index(drop=True)

    main_idx = find_column_index_by_keywords(report_df, ("授課老師", "授課教師", "主授課老師"))
    co_idx = find_column_index_by_keywords(report_df, ("協同老師", "協同教師", "協同授課老師"))
    if main_idx is None:
        main_idx = column_index(settings.main_teacher_name)
    if co_idx is None:
        co_idx = column_index(settings.co_teacher_name)
    ensure_column_count(report_df, max(main_idx, co_idx, column_index("AD")) + 1)

    used: set[int] = set()
    output = pd.DataFrame()
    for header in TEMPLATE_HEADERS:
        label = header if header is not None else ""
        output[label] = get_series_by_template_header(report_df, header, used).map(normalize_value)

    # 重新依完成版邏輯產出指定欄位
    output["#"] = range(1, len(output) + 1)
    output["排除重複課程"] = report_df[dedupe_col].map(normalize_value).str.replace("|", "", regex=False)
    output["授課老師"] = report_df.iloc[:, main_idx].map(normalize_value)
    output["協同老師"] = report_df.iloc[:, co_idx].map(normalize_value)
    output["授課老師員編"] = output["授課老師"].map(lambda value: ids_for_names(value, lookup))
    output["協同老師員編"] = output["協同老師"].map(lambda value: ids_for_names(value, lookup))
    output["該課程授課老師為住院醫師"] = output["授課老師員編"]
    output["該課程協同&授課老師皆為同一位住院醫師"] = output.apply(
        lambda row: "是" if row["授課老師員編"] and row["授課老師員編"] == row["協同老師員編"] else "否",
        axis=1,
    )
    return output


def build_summary(course_total: pd.DataFrame) -> pd.DataFrame:
    mask = (
        course_total["該課程授課老師為住院醫師"].map(normalize_value).ne("") |
        course_total["協同老師員編"].map(normalize_value).ne("")
    )
    summary = course_total.loc[mask].copy().reset_index(drop=True)
    summary["#"] = range(1, len(summary) + 1)
    return summary


def copy_sheet_layout_from_template(template_ws, target_ws, max_rows: int, max_cols: int) -> None:
    for row in range(1, max_rows + 1):
        target_ws.row_dimensions[row].height = template_ws.row_dimensions[row].height
    for col in range(1, max_cols + 1):
        letter = get_column_letter(col)
        target_ws.column_dimensions[letter].width = template_ws.column_dimensions[letter].width

    style_source_row = 2 if template_ws.max_row >= 2 else 1
    for col in range(1, max_cols + 1):
        src = template_ws.cell(1, col)
        dst = target_ws.cell(1, col)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format

        src_body = template_ws.cell(style_source_row, col)
        for row in range(2, max_rows + 1):
            dst_body = target_ws.cell(row, col)
            dst_body.font = copy(src_body.font)
            dst_body.fill = copy(src_body.fill)
            dst_body.border = copy(src_body.border)
            dst_body.alignment = copy(src_body.alignment)
            dst_body.number_format = src_body.number_format


def apply_default_layout(ws, max_rows: int, max_cols: int, summary: bool = False) -> None:
    header_fill = PatternFill("solid", fgColor="CCCCCC")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, max_cols + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = DEFAULT_WIDTHS.get(col, 13)
    ws.row_dimensions[1].height = 72 if summary else 43.2
    for row in range(1, max_rows + 1):
        for col in range(1, max_cols + 1):
            cell = ws.cell(row, col)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if row == 1:
                cell.fill = header_fill
                cell.font = Font(name="Calibri", size=11, color="000000")


def write_dataframe(ws, df: pd.DataFrame) -> None:
    for c, col_name in enumerate(df.columns, 1):
        ws.cell(1, c).value = col_name
    for r, row in enumerate(df.itertuples(index=False, name=None), 2):
        for c, value in enumerate(row, 1):
            ws.cell(r, c).value = None if pd.isna(value) or value == "" else value
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{max(len(df) + 1, 1)}"


def create_output_workbook(course_total: pd.DataFrame, summary: pd.DataFrame, mapping_df: pd.DataFrame, template_file=None) -> io.BytesIO:
    template_wb = None
    if template_file is not None:
        template_wb = load_workbook(template_file)
    elif Path("115年6月RAST.xlsx").exists():
        template_wb = load_workbook("115年6月RAST.xlsx")

    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, df in (("課程總表", course_total), ("彙整", summary)):
        ws = wb.create_sheet(sheet_name)
        write_dataframe(ws, df)
        max_rows = len(df) + 1
        max_cols = len(df.columns)
        if template_wb and sheet_name in template_wb.sheetnames:
            copy_sheet_layout_from_template(template_wb[sheet_name], ws, max_rows, max_cols)
        else:
            apply_default_layout(ws, max_rows, max_cols, summary=(sheet_name == "彙整"))

    ws_map = wb.create_sheet("比對用")
    mapping_output = mapping_df.copy().map(normalize_value)
    write_dataframe(ws_map, mapping_output)
    if template_wb and "比對用" in template_wb.sheetnames:
        copy_sheet_layout_from_template(template_wb["比對用"], ws_map, len(mapping_output) + 1, len(mapping_output.columns))
    else:
        for col in range(1, len(mapping_output.columns) + 1):
            ws_map.column_dimensions[get_column_letter(col)].width = 18
            ws_map.cell(1, col).font = Font(bold=True)
            ws_map.cell(1, col).alignment = Alignment(horizontal="center")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def process_report(report_file, mapping_file, settings: ColumnSettings, template_file=None) -> tuple[io.BytesIO, int, int, int]:
    report_df = pd.read_excel(report_file, dtype=object, engine="openpyxl").astype(object)
    mapping_df = pd.read_excel(mapping_file, dtype=object, engine="openpyxl").astype(object)
    original_rows = len(report_df)
    lookup = build_teacher_lookup(mapping_df)
    course_total = build_course_total(report_df, lookup, settings)
    summary = build_summary(course_total)
    output = create_output_workbook(course_total, summary, mapping_df, template_file)
    return output, original_rows, len(course_total), len(summary)


st.set_page_config(page_title="EP/RAST 課程報表整理", layout="centered")
st.title("EP/RAST 課程報表整理")
st.caption("上傳原始版報表與比對用住院醫師名單，輸出格式會比照完成版；課程總表會完整保留每一筆原始資料，不會去重、不會篩掉、不會拆列。")

report_file = st.file_uploader("上傳原始版 Excel 報表", type=["xlsx"])
mapping_file = st.file_uploader("上傳比對用住院醫師名單", type=["xlsx"])
template_file = st.file_uploader("上傳完成版樣板（選填；若要完全複製排版請上傳 115年6月RAST.xlsx）", type=["xlsx"])

with st.expander("欄位設定"):
    left, right = st.columns(2)
    with left:
        main_teacher_name = st.text_input("授課老師姓名欄", value="S")
    with right:
        co_teacher_name = st.text_input("協同老師姓名欄", value="U")

if st.button("整理報表", type="primary"):
    if report_file is None or mapping_file is None:
        st.error("請同時上傳原始版 Excel 報表與比對用住院醫師名單。")
    else:
        try:
            output, original_rows, total_rows, summary_rows = process_report(
                report_file,
                mapping_file,
                ColumnSettings(main_teacher_name=main_teacher_name, co_teacher_name=co_teacher_name),
                template_file=template_file,
            )
            st.success(f"整理完成：原始 {original_rows} 筆，課程總表保留 {total_rows} 筆，彙整 {summary_rows} 筆。")
            st.download_button(
                "下載整理後 Excel",
                data=output,
                file_name=f"整理後RAST報表_總表{total_rows}筆_彙整{summary_rows}筆.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"處理失敗：{exc}")

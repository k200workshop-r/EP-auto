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


COURSE_TOTAL_HEADERS = [
    "#", "課程代碼", "學年度", "醫院別", "開課單位", "課程名稱", "課程類型", "排除重複課程",
    "上課時間", "開始日期", "結束日期", "上課時間", "上課分鐘數", "表單發送時間", "課程分類", "課程備註",
    "上課地點", "該課程授課老師為住院醫師", "授課老師", "授課老師員編", "協同老師", "協同老師員編",
    "該課程協同&授課老師皆為同一位住院醫師", "學生", "輔助教材", "職類", "計畫類別", "訓練計畫", "訓練科室", "符合/不符合",
]

SUMMARY_HEADERS = [
    "#", "課程代碼", "學年度", "醫院別", "開課單位", "課程名稱", "課程類型", "排除重複課程",
    "上課時間", "開始日期", "結束日期", "上課時間", "上課分鐘數", "表單發送時間", "課程分類", "課程備註",
    "上課地點", "該課程授課老師為住院醫師", "授課老師", "授課老師員編", "協同老師", "協同老師員編",
    "該課程協同&授課老師皆為同一位住院醫師", "學生", "輔助教材", "職類", "計畫類別", "訓練計畫", "訓練科室", "A.住院醫師主授課程/B.住院醫師擔任協同老師",
]


@dataclass(frozen=True)
class ColumnSettings:
    # 只有原始檔找不到欄名時，才會使用欄位代號備援。
    main_teacher_name: str = "S"
    co_teacher_name: str = "U"

    def normalized(self) -> "ColumnSettings":
        return ColumnSettings(
            main_teacher_name=str(self.main_teacher_name).strip().upper(),
            co_teacher_name=str(self.co_teacher_name).strip().upper(),
        )


def column_index(letter: str) -> int:
    value = 0
    cleaned = str(letter).strip().upper()
    if not re.fullmatch(r"[A-Z]+", cleaned):
        raise ValueError(f"欄位代號不正確：{letter}")
    for char in cleaned:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value - 1


def normalize_header(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", "", str(value).strip())


def normalize_value(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def strip_employee_id_from_name(text: object) -> str:
    value = normalize_value(text)
    if not value:
        return ""
    # 移除半形/全形括號中的員編，例如 王小明(223001)、王小明（223001）
    value = re.sub(r"[（(]\s*\d+(?:\.0)?\s*[）)]", "", value)
    return normalize_header(value)


def split_teacher_names(value: object) -> list[str]:
    text = normalize_value(value)
    if not text:
        return []
    parts = re.split(r"[、,，;；/／\n\r]+", text)
    names: list[str] = []
    for part in parts:
        name = strip_employee_id_from_name(part)
        if name and name not in names:
            names.append(name)
    return names


def join_names(value: object) -> str:
    return "、".join(split_teacher_names(value))


def employee_ids_in_text(value: object) -> list[str]:
    text = normalize_value(value)
    if not text:
        return []
    ids: list[str] = []
    for found in re.findall(r"[（(]\s*(\d+(?:\.0)?)\s*[）)]", text):
        emp_id = normalize_value(found)
        if emp_id and emp_id not in ids:
            ids.append(emp_id)
    return ids


def first_matching_column(columns: Iterable[object], keywords: tuple[str, ...]) -> object | None:
    normalized = [(column, normalize_header(column)) for column in columns]
    for keyword in keywords:
        key = normalize_header(keyword)
        for column, column_name in normalized:
            if column_name == key:
                return column
    for keyword in keywords:
        key = normalize_header(keyword)
        for column, column_name in normalized:
            if key and key in column_name:
                return column
    return None


def find_column_index_by_keywords(df: pd.DataFrame, keywords: tuple[str, ...]) -> int | None:
    normalized_columns = [normalize_header(column) for column in df.columns]
    for keyword in keywords:
        key = normalize_header(keyword)
        for index, column_name in enumerate(normalized_columns):
            if column_name == key:
                return index
    for keyword in keywords:
        key = normalize_header(keyword)
        for index, column_name in enumerate(normalized_columns):
            if key and key in column_name:
                return index
    return None


def get_series_by_header_occurrence(df: pd.DataFrame, header: str, occurrence: int = 1) -> pd.Series:
    """依欄名抓資料；可處理重複欄名，例如第 1 個/第 2 個「上課時間」。"""
    target = normalize_header(header)
    count = 0
    for idx, col in enumerate(df.columns):
        clean_col = re.sub(r"\.\d+$", "", normalize_header(col))
        if clean_col == target:
            count += 1
            if count == occurrence:
                return df.iloc[:, idx].map(normalize_value)
    return pd.Series([""] * len(df), dtype=object)


def get_series_by_index(df: pd.DataFrame, index: int) -> pd.Series:
    if 0 <= index < len(df.columns):
        return df.iloc[:, index].map(normalize_value)
    return pd.Series([""] * len(df), dtype=object)


def build_teacher_lookup(mapping_df: pd.DataFrame) -> dict[str, str]:
    name_col = first_matching_column(mapping_df.columns, ("中文姓名", "教師姓名", "老師姓名", "姓名", "授課老師", "協同老師"))
    employee_col = first_matching_column(mapping_df.columns, ("員工編號", "員編", "教師員編", "醫師員編", "住院醫師員編", "職編", "ID"))
    if name_col is None or employee_col is None:
        raise ValueError("比對用住院醫師名單需包含姓名欄位與員編欄位，例如「中文姓名」與「員工編號」。")

    lookup: dict[str, str] = {}
    for _, row in mapping_df.iterrows():
        name = strip_employee_id_from_name(row[name_col])
        employee_id = normalize_value(row[employee_col])
        if name and employee_id:
            lookup[name] = employee_id
    return lookup


def lookup_ids_for_names(value: object, lookup: dict[str, str]) -> list[str]:
    ids: list[str] = []
    for name in split_teacher_names(value):
        emp_id = lookup.get(name)
        if emp_id and emp_id not in ids:
            ids.append(emp_id)
    return ids


def lookup_ids_text(value: object, lookup: dict[str, str], na_text: str = "#N/A") -> str:
    ids = lookup_ids_for_names(value, lookup)
    return "、".join(ids) if ids else na_text


def build_course_total(report_df: pd.DataFrame, lookup: dict[str, str], settings: ColumnSettings) -> pd.DataFrame:
    """建立「課程總表」：完整保留原始資料列，不去重、不刪除、不篩選。"""
    settings = settings.normalized()
    report_df = report_df.astype(object).reset_index(drop=True)

    main_idx = find_column_index_by_keywords(report_df, ("授課老師", "授課教師", "主授課老師"))
    co_idx = find_column_index_by_keywords(report_df, ("協同老師", "協同教師", "協同授課老師"))
    main_emp_idx = find_column_index_by_keywords(report_df, ("授課老師員編", "授課教師員編", "主授課老師員編"))

    if main_idx is None:
        main_idx = column_index(settings.main_teacher_name)
    if co_idx is None:
        co_idx = column_index(settings.co_teacher_name)

    raw_main = get_series_by_index(report_df, main_idx)
    raw_co = get_series_by_index(report_df, co_idx)
    clean_main = raw_main.map(join_names)
    clean_co = raw_co.map(join_names)

    if main_emp_idx is not None:
        main_emp = get_series_by_index(report_df, main_emp_idx)
    else:
        main_emp = raw_main.map(lambda value: "、".join(employee_ids_in_text(value)))

    main_resident_ids = raw_main.map(lambda value: lookup_ids_text(value, lookup))
    co_resident_ids = raw_co.map(lambda value: lookup_ids_text(value, lookup))

    # 依照完成版 Excel 公式邏輯：排除重複課程 = I欄 & Q欄 & S欄
    dedupe_key = (
        get_series_by_index(report_df, column_index("I")) +
        get_series_by_index(report_df, column_index("Q")) +
        get_series_by_index(report_df, column_index("S"))
    )

    data = {
        "#": get_series_by_header_occurrence(report_df, "#"),
        "課程代碼": get_series_by_header_occurrence(report_df, "課程代碼"),
        "學年度": get_series_by_header_occurrence(report_df, "學年度"),
        "醫院別": get_series_by_header_occurrence(report_df, "醫院別"),
        "開課單位": get_series_by_header_occurrence(report_df, "開課單位"),
        "課程名稱": get_series_by_header_occurrence(report_df, "課程名稱"),
        "課程類型": get_series_by_header_occurrence(report_df, "課程類型"),
        "排除重複課程": dedupe_key,
        "上課時間_1": get_series_by_header_occurrence(report_df, "上課時間", 1),
        "開始日期": get_series_by_header_occurrence(report_df, "開始日期"),
        "結束日期": get_series_by_header_occurrence(report_df, "結束日期"),
        "上課時間_2": get_series_by_header_occurrence(report_df, "上課時間", 2),
        "上課分鐘數": get_series_by_header_occurrence(report_df, "上課分鐘數"),
        "表單發送時間": get_series_by_header_occurrence(report_df, "表單發送時間"),
        "課程分類": get_series_by_header_occurrence(report_df, "課程分類"),
        "課程備註": get_series_by_header_occurrence(report_df, "課程備註"),
        "上課地點": get_series_by_header_occurrence(report_df, "上課地點"),
        "該課程授課老師為住院醫師": main_resident_ids,
        "授課老師": clean_main,
        "授課老師員編": main_emp,
        "協同老師": clean_co,
        "協同老師員編": co_resident_ids,
        "該課程協同&授課老師皆為同一位住院醫師": [
            "是" if r != "#N/A" and r == c else "否"
            for r, c in zip(main_resident_ids, co_resident_ids)
        ],
        "學生": get_series_by_header_occurrence(report_df, "學生"),
        "輔助教材": get_series_by_header_occurrence(report_df, "輔助教材"),
        "職類": get_series_by_header_occurrence(report_df, "職類"),
        "計畫類別": get_series_by_header_occurrence(report_df, "計畫類別"),
        "訓練計畫": get_series_by_header_occurrence(report_df, "訓練計畫"),
        "訓練科室": get_series_by_header_occurrence(report_df, "訓練科室"),
        "符合/不符合": [
            "需刪除" if normalize_header(co) and normalize_header(co) == normalize_header(stu) else "符合"
            for co, stu in zip(clean_co, get_series_by_header_occurrence(report_df, "學生"))
        ],
    }

    rows = []
    for i in range(len(report_df)):
        rows.append([
            data["#"].iloc[i] or i + 1,
            data["課程代碼"].iloc[i],
            data["學年度"].iloc[i],
            data["醫院別"].iloc[i],
            data["開課單位"].iloc[i],
            data["課程名稱"].iloc[i],
            data["課程類型"].iloc[i],
            data["排除重複課程"].iloc[i],
            data["上課時間_1"].iloc[i],
            data["開始日期"].iloc[i],
            data["結束日期"].iloc[i],
            data["上課時間_2"].iloc[i],
            data["上課分鐘數"].iloc[i],
            data["表單發送時間"].iloc[i],
            data["課程分類"].iloc[i],
            data["課程備註"].iloc[i],
            data["上課地點"].iloc[i],
            data["該課程授課老師為住院醫師"].iloc[i],
            data["授課老師"].iloc[i],
            data["授課老師員編"].iloc[i],
            data["協同老師"].iloc[i],
            data["協同老師員編"].iloc[i],
            data["該課程協同&授課老師皆為同一位住院醫師"][i],
            data["學生"].iloc[i],
            data["輔助教材"].iloc[i],
            data["職類"].iloc[i],
            data["計畫類別"].iloc[i],
            data["訓練計畫"].iloc[i],
            data["訓練科室"].iloc[i],
            data["符合/不符合"][i],
        ])

    return pd.DataFrame(rows, columns=COURSE_TOTAL_HEADERS)


def build_summary(course_total: pd.DataFrame, lookup: dict[str, str]) -> pd.DataFrame:
    """建立「彙整」：只彙整住院醫師主授或住院醫師協同；不影響課程總表原始列。"""
    summary_rows: list[list[object]] = []

    for _, row in course_total.iterrows():
        row_values = list(row.iloc[:30])
        is_valid = normalize_value(row_values[29]) != "需刪除"
        if not is_valid:
            continue

        main_resident_id = normalize_value(row_values[17])
        if main_resident_id and main_resident_id != "#N/A":
            new_row = row_values.copy()
            new_row[29] = "A.住院醫師主授課程"
            summary_rows.append(new_row)

        # 協同老師若有多位，依完成版彙整習慣拆成一位一列。
        co_names = split_teacher_names(row_values[20])
        for co_name in co_names:
            co_id = lookup.get(co_name)
            if not co_id:
                continue
            new_row = row_values.copy()
            new_row[20] = co_name
            new_row[21] = co_id
            new_row[22] = "是" if main_resident_id != "#N/A" and main_resident_id == co_id else "否"
            new_row[29] = "B.住院醫師擔任協同老師"
            summary_rows.append(new_row)

    summary = pd.DataFrame(summary_rows, columns=SUMMARY_HEADERS)
    return summary


def copy_sheet_layout_from_template(template_ws, target_ws, max_rows: int, max_cols: int) -> None:
    # 欄寬與列高
    for col in range(1, max_cols + 1):
        letter = get_column_letter(col)
        target_ws.column_dimensions[letter].width = template_ws.column_dimensions[letter].width
    for row in range(1, max_rows + 1):
        target_ws.row_dimensions[row].height = template_ws.row_dimensions[row].height

    # 凍結窗格、篩選
    target_ws.freeze_panes = template_ws.freeze_panes

    # 樣式：第 1 列用樣板第 1 列，資料列用樣板第 2 列
    body_source_row = 2 if template_ws.max_row >= 2 else 1
    for col in range(1, max_cols + 1):
        for row in range(1, max_rows + 1):
            src = template_ws.cell(1 if row == 1 else body_source_row, col)
            dst = target_ws.cell(row, col)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)


def apply_default_layout(ws, max_rows: int, max_cols: int) -> None:
    header_fill = PatternFill("solid", fgColor="D9D9D9")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, max_cols + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 18
    ws.row_dimensions[1].height = 45
    for row in range(1, max_rows + 1):
        for col in range(1, max_cols + 1):
            cell = ws.cell(row, col)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if row == 1:
                cell.fill = header_fill
                cell.font = Font(bold=True)


def write_dataframe(ws, df: pd.DataFrame, headers: list[str], add_course_formula: bool = False) -> None:
    for c, col_name in enumerate(headers, 1):
        ws.cell(1, c).value = col_name

    for r, row in enumerate(df.itertuples(index=False, name=None), 2):
        for c, value in enumerate(row, 1):
            ws.cell(r, c).value = None if normalize_value(value) == "" else value
        if add_course_formula:
            # 完成版公式：AD = IF(U = X, "需刪除", "符合")
            ws.cell(r, 30).value = f'=IF(U{r}=X{r},"需刪除","符合")'

    if len(df) > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(df) + 1}"


def write_mapping_sheet(ws, mapping_df: pd.DataFrame) -> None:
    clean_df = mapping_df.copy().astype(object)
    for c, col_name in enumerate(clean_df.columns, 1):
        ws.cell(1, c).value = col_name
    for r, row in enumerate(clean_df.itertuples(index=False, name=None), 2):
        for c, value in enumerate(row, 1):
            ws.cell(r, c).value = None if normalize_value(value) == "" else normalize_value(value)
    if len(clean_df) > 0 and len(clean_df.columns) > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(clean_df.columns))}{len(clean_df) + 1}"


def create_output_workbook(course_total: pd.DataFrame, summary: pd.DataFrame, mapping_df: pd.DataFrame, template_file=None) -> io.BytesIO:
    template_wb = None
    if template_file is not None:
        template_wb = load_workbook(template_file)

    wb = Workbook()
    wb.remove(wb.active)

    ws_total = wb.create_sheet("課程總表")
    write_dataframe(ws_total, course_total, COURSE_TOTAL_HEADERS, add_course_formula=True)
    if template_wb and "課程總表" in template_wb.sheetnames:
        copy_sheet_layout_from_template(template_wb["課程總表"], ws_total, len(course_total) + 1, len(COURSE_TOTAL_HEADERS))
    else:
        apply_default_layout(ws_total, len(course_total) + 1, len(COURSE_TOTAL_HEADERS))

    ws_summary = wb.create_sheet("彙整")
    write_dataframe(ws_summary, summary, SUMMARY_HEADERS, add_course_formula=False)
    if template_wb and "彙整" in template_wb.sheetnames:
        copy_sheet_layout_from_template(template_wb["彙整"], ws_summary, len(summary) + 1, len(SUMMARY_HEADERS))
    else:
        apply_default_layout(ws_summary, len(summary) + 1, len(SUMMARY_HEADERS))

    ws_map = wb.create_sheet("比對用")
    write_mapping_sheet(ws_map, mapping_df)
    if template_wb and "比對用" in template_wb.sheetnames:
        copy_sheet_layout_from_template(template_wb["比對用"], ws_map, len(mapping_df) + 1, max(len(mapping_df.columns), 1))
    else:
        apply_default_layout(ws_map, len(mapping_df) + 1, max(len(mapping_df.columns), 1))

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
    summary = build_summary(course_total, lookup)
    output = create_output_workbook(course_total, summary, mapping_df, template_file)
    return output, original_rows, len(course_total), len(summary)


st.set_page_config(page_title="RAST 課程自動報表", layout="centered")
st.title("RAST 課程自動報表")
st.caption("依完成版 RAST 格式輸出：課程總表完整保留原始每一筆資料；彙整表另依住院醫師主授／協同授課自動產生。")

report_file = st.file_uploader("上傳原始版 Excel 報表", type=["xlsx"])
mapping_file = st.file_uploader("上傳比對用住院醫師名單", type=["xlsx"])
template_file = st.file_uploader("上傳完成版樣板（建議上傳 115年6月RAST.xlsx，可複製相同排版）", type=["xlsx"])

with st.expander("欄位設定（原始檔找不到欄名時才會使用）", expanded=False):
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
            st.success(f"整理完成：原始 {original_rows} 筆，課程總表完整保留 {total_rows} 筆，彙整 {summary_rows} 筆。")
            st.download_button(
                "下載整理後 Excel",
                data=output,
                file_name=f"整理後RAST報表_課程總表{total_rows}筆_彙整{summary_rows}筆.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"處理失敗：{exc}")

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd
import streamlit as st


@dataclass(frozen=True)
class ColumnSettings:
    main_teacher_name: str = "S"
    main_teacher_status: str = "T"
    co_teacher_name: str = "U"
    co_teacher_status: str = "V"


def column_index(letter: str) -> int:
    value = 0
    cleaned = letter.strip().upper()
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
    return "" if text.lower() == "nan" else text


def split_teacher_names(value: object) -> list[str]:
    text = normalize_value(value)
    if not text:
        return []

    parts = re.split(r"[、,，;；/／\n\r]+", text)
    return [normalize_name(part) for part in parts if normalize_name(part)]


def first_matching_column(
    columns: Iterable[object],
    keywords: tuple[str, ...],
) -> object | None:
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


def find_column_index_by_keywords(df: pd.DataFrame, keywords: tuple[str, ...]) -> int | None:
    matched_column = first_matching_column(df.columns, keywords)
    if matched_column is None:
        return None
    return df.columns.get_loc(matched_column)


def unique_column_name(df: pd.DataFrame, base_name: str) -> str:
    if base_name not in df.columns:
        return base_name

    suffix = 2
    while f"{base_name}_{suffix}" in df.columns:
        suffix += 1

    return f"{base_name}_{suffix}"


def build_teacher_lookup(mapping_df: pd.DataFrame) -> dict[str, str]:
    name_col = first_matching_column(
        mapping_df.columns,
        ("教師姓名", "老師姓名", "姓名", "授課老師", "協同老師"),
    )
    employee_col = first_matching_column(
        mapping_df.columns,
        ("員編", "員工編號", "教師員編", "醫師員編", "住院醫師員編", "職編", "ID"),
    )

    if name_col is None or employee_col is None:
        raise ValueError("住院醫師名單需包含姓名欄位與員編欄位，例如「姓名」與「員編」。")

    lookup: dict[str, str] = {}
    for _, row in mapping_df.iterrows():
        name = normalize_name(row[name_col])
        employee_id = normalize_value(row[employee_col])

        if name and employee_id:
            lookup[name] = employee_id

    return lookup


def teacher_ids_for_value(value: object, lookup: dict[str, str]) -> str:
    employee_ids = [
        lookup[name]
        for name in split_teacher_names(value)
        if name in lookup
    ]
    return "、".join(employee_ids)


def mark_teacher_status(
    df: pd.DataFrame,
    lookup: dict[str, str],
    name_letter: str,
    status_letter: str,
    found_text: str,
) -> None:
    name_idx = column_index(name_letter)
    status_idx = column_index(status_letter)
    ensure_column_count(df, max(name_idx, status_idx) + 1)

    df.iloc[:, status_idx] = df.iloc[:, name_idx].map(
        lambda value: found_text
        if any(name in lookup for name in split_teacher_names(value))
        else ""
    )


def split_multi_co_teacher_rows(
    df: pd.DataFrame,
    lookup: dict[str, str],
    settings: ColumnSettings,
) -> pd.DataFrame:
    main_teacher_idx = find_column_index_by_keywords(
        df,
        ("授課老師", "授課教師", "主授課老師"),
    )
    co_teacher_idx = find_column_index_by_keywords(
        df,
        ("協同老師", "協同教師", "協同授課老師"),
    )

    if main_teacher_idx is None:
        main_teacher_idx = column_index(settings.main_teacher_name)
    if co_teacher_idx is None:
        co_teacher_idx = column_index(settings.co_teacher_name)

    ensure_column_count(
        df,
        max(main_teacher_idx, co_teacher_idx, column_index("T")) + 1,
    )

    output_rows = []
    for _, row in df.iterrows():
        main_teacher_names = split_teacher_names(row.iloc[main_teacher_idx])
        co_teacher_names = split_teacher_names(row.iloc[co_teacher_idx])
        main_teacher_is_resident = any(name in lookup for name in main_teacher_names)

        if len(co_teacher_names) < 2:
            co_teacher_is_resident = any(name in lookup for name in co_teacher_names)
            if main_teacher_is_resident or co_teacher_is_resident:
                output_rows.append(row.copy())
            continue

        for index, co_teacher_name in enumerate(co_teacher_names):
            co_teacher_is_resident = co_teacher_name in lookup
            if not main_teacher_is_resident and not co_teacher_is_resident:
                continue

            new_row = row.copy()
            new_row.iloc[co_teacher_idx] = co_teacher_name

            if index > 0:
                for letter in ("R", "S", "T"):
                    new_row.iloc[column_index(letter)] = ""

            output_rows.append(new_row)

    if not output_rows:
        return df.iloc[0:0].copy()

    return pd.DataFrame(output_rows, columns=df.columns).reset_index(drop=True)


def prepare_output_columns(
    df: pd.DataFrame,
    lookup: dict[str, str],
    settings: ColumnSettings,
) -> pd.DataFrame:
    result_df = df.copy()

    co_teacher_idx = find_column_index_by_keywords(
        result_df,
        ("協同老師", "協同教師", "協同授課老師"),
    )
    if co_teacher_idx is None:
        co_teacher_idx = column_index(settings.co_teacher_name)

    if "協同老師員編(使用vlookup比對)" in result_df.columns:
        result_df = result_df.drop(columns=["協同老師員編(使用vlookup比對)"])
        co_teacher_idx = min(co_teacher_idx, len(result_df.columns) - 1)

    co_teacher_col = result_df.columns[co_teacher_idx]
    co_teacher_ids = result_df[co_teacher_col].map(
        lambda value: teacher_ids_for_value(value, lookup)
    )
    result_df.insert(
        co_teacher_idx + 1,
        "協同老師員編(使用vlookup比對)",
        co_teacher_ids,
    )

    drop_columns = [
        column
        for column in result_df.columns
        if normalize_name(column) in {"建課日期", "課程備註"}
    ]
    if drop_columns:
        result_df = result_df.drop(columns=drop_columns)

    return result_df


def process_report(
    report_file,
    mapping_file,
    settings: ColumnSettings,
) -> tuple[io.BytesIO, int, int]:
    report_df = pd.read_excel(report_file, dtype=object, engine="openpyxl")
    mapping_df = pd.read_excel(mapping_file, dtype=object, engine="openpyxl")

    ensure_column_count(report_df, column_index("AD") + 1)
    original_rows = len(report_df)

    i_idx = column_index("I")
    q_idx = column_index("Q")
    s_idx = column_index("S")
    dedupe_col = unique_column_name(report_df, "排除重複課程")

    report_df[dedupe_col] = (
        report_df.iloc[:, i_idx].map(normalize_value)
        + "|"
        + report_df.iloc[:, q_idx].map(normalize_value)
        + "|"
        + report_df.iloc[:, s_idx].map(normalize_value)
    )
    report_df = report_df.drop_duplicates(
        subset=[dedupe_col],
        keep="first",
    ).reset_index(drop=True)

    lookup = build_teacher_lookup(mapping_df)
    report_df = split_multi_co_teacher_rows(report_df, lookup, settings)

    mark_teacher_status(
        report_df,
        lookup,
        settings.main_teacher_name,
        settings.main_teacher_status,
        "授課老師有員編",
    )
    mark_teacher_status(
        report_df,
        lookup,
        settings.co_teacher_name,
        settings.co_teacher_status,
        "協同老師有員編",
    )

    ad_idx = column_index("AD")
    t_idx = column_index(settings.main_teacher_status)
    v_idx = column_index(settings.co_teacher_status)

    columns = list(report_df.columns)
    columns[ad_idx] = "課程分類"
    report_df.columns = columns
    report_df.iloc[:, ad_idx] = ""

    main_mask = report_df.iloc[:, t_idx] == "授課老師有員編"
    co_mask = (
        (report_df.iloc[:, ad_idx] == "")
        & (report_df.iloc[:, v_idx] == "協同老師有員編")
    )
    report_df.iloc[main_mask.to_numpy(), ad_idx] = "A.住院醫師主授課程"
    report_df.iloc[co_mask.to_numpy(), ad_idx] = "B.住院醫師擔任協同老師"

    report_df = prepare_output_columns(report_df, lookup, settings)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report_df.to_excel(writer, index=False, sheet_name="整理後報表")
    output.seek(0)

    return output, original_rows, len(report_df)


st.set_page_config(
    page_title="EP 課程報表整理",
    layout="centered",
)

st.title("EP 課程報表整理")
st.caption("上傳原始版報表與比對用住院醫師名單，自動去重、拆分協同老師、比對員編並完成課程分類。")

report_file = st.file_uploader("上傳原始版 Excel 報表", type=["xlsx"])
mapping_file = st.file_uploader("上傳比對用住院醫師名單", type=["xlsx"])

with st.expander("欄位設定"):
    left, right = st.columns(2)

    with left:
        main_teacher_name = st.text_input("授課老師姓名欄", value="S")
        co_teacher_name = st.text_input("協同老師姓名欄", value="U")

    with right:
        main_teacher_status = st.text_input("授課老師比對結果欄", value="T")
        co_teacher_status = st.text_input("協同老師比對結果欄", value="V")

if st.button("整理報表", type="primary"):
    if report_file is None or mapping_file is None:
        st.error("請同時上傳原始版 Excel 報表與比對用住院醫師名單。")
    else:
        try:
            output, original_rows, final_rows = process_report(
                report_file,
                mapping_file,
                ColumnSettings(
                    main_teacher_name=main_teacher_name,
                    main_teacher_status=main_teacher_status,
                    co_teacher_name=co_teacher_name,
                    co_teacher_status=co_teacher_status,
                ),
            )

            st.success(f"整理完成：原始 {original_rows} 筆，輸出 {final_rows} 筆。")

            st.download_button(
                "下載整理後 Excel",
                data=output,
                file_name=f"整理後報表_原{original_rows}筆_輸出{final_rows}筆.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as exc:
            st.error(f"處理失敗：{exc}")

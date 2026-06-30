# -*- coding: utf-8 -*-
"""
Created on Tue Jun 30 10:45:08 2026

@author: cilia
0630
微幅修改例外處理
V2_03：遇錯直接終止
V2_03.1：遇到錯誤跳過並記錄（新增 Error_Log 工作表）
"""

import os
import re
import numpy as np
import pandas as pd
from tkinter import Tk, filedialog
from xlsxwriter.utility import xl_col_to_name


# =========================
# 1. 選擇 Excel 檔案
# =========================
root = Tk()
root.withdraw()

file_path = filedialog.askopenfilename(
    title="選擇 IPGTT Excel 檔案",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not file_path:
    raise ValueError("沒有選擇檔案")


# =========================
# 2. 讀取資料
# =========================
df = pd.read_excel(file_path, sheet_name=0)
df = df.dropna(how="all").copy()
df = df.fillna("")
required_cols = ["Group", "Mouse_ID"]

for col in required_cols:
    if col not in df.columns:
        raise ValueError(f"缺少必要欄位：{col}")

df["Group"] = df["Group"].ffill()


# =========================
# 3. 特殊血糖值轉換
# =========================
def is_special_value(value):

    if pd.isna(value):
        return False

    text = str(value).strip().upper()

    return text.startswith("HI/") or text == "LO"


def convert_glucose(value):
    # 這裡幫你加一個保險：如果 value 根本不是數字也不是預期的格式，直接回傳 np.nan
    try:
        # 先轉字串處理空格
        text = str(value).strip().upper()
        
        # 你的原本處理 HI/LO 的邏輯 (維持不變)
        hi_match = re.match(r"^HI/(\d+(\.\d+)?)$", text)
        if hi_match:
            n = float(hi_match.group(1))
            return 605 if n < 120 else n * 5
        if text == "LO":
            return 19
            
        # 最後再嘗試轉 float
        return float(text)
    except:
        # 如果上面任何一步掛掉，通通視為 NaN，程式繼續跑，不會閃退
        return np.nan

# =========================
# 4. 轉成 long format
# =========================
time_cols = [col for col in df.columns if col not in ["Group", "Mouse_ID"]]

long_df = df.melt(
    id_vars=["Group", "Mouse_ID"],
    value_vars=time_cols,
    var_name="Time",
    value_name="Original_Value"
)

long_df["Time"] = pd.to_numeric(long_df["Time"], errors="coerce")
long_df["Glucose"] = long_df["Original_Value"].apply(convert_glucose)
long_df["Is_Converted"] = long_df["Original_Value"].apply(is_special_value)

# -------------------------
# 4-1. Error_Log：記錄無法轉換的輸入值
# -------------------------
# 規則：
# - 空白值不視為錯誤，因為可能是刻意留空或尚未量測
# - 非空白但轉換後為 NaN，代表使用者輸入格式無法辨識
# - 例如：abc、HI/??、--、無效文字等
error_log_df = long_df[
    (long_df["Original_Value"].astype(str).str.strip() != "")
    & (long_df["Glucose"].isna())
].copy()

error_log_df = error_log_df[[
    "Group",
    "Mouse_ID",
    "Time",
    "Original_Value"
]]

error_log_df["Issue"] = "無法轉換為有效血糖值，已於計算中視為空白/NaN"

long_df = long_df.sort_values(
    ["Group", "Mouse_ID", "Time"]
).reset_index(drop=True)
# =========================
# 5. Processed_Wide
# =========================
processed_wide = long_df.pivot_table(
    index=["Group", "Mouse_ID"],
    columns="Time",
    values="Glucose",
    aggfunc="first"
).reset_index()

processed_wide.columns.name = None

converted_flag = long_df.pivot_table(
    index=["Group", "Mouse_ID"],
    columns="Time",
    values="Is_Converted",
    aggfunc="first"
).reset_index()

converted_flag.columns.name = None


# =========================
# 6. Group順序、時間順序
# =========================
group_order = list(df["Group"].dropna().drop_duplicates())
time_order = sorted(long_df["Time"].dropna().unique())

AUC_Y_AXIS_MAX = 3500   # 主任要求：AUC長條圖固定最高上限，可在這裡改


# =========================
# 7. 輸出 Excel
# =========================
folder = os.path.dirname(file_path)
base_name = os.path.splitext(os.path.basename(file_path))[0]

output_path = os.path.join(
    folder,
    f"{base_name}_IPGTT_V2_03_Report.xlsx"
)

with pd.ExcelWriter(
    output_path,
    engine="xlsxwriter",
    engine_kwargs={"options": {"nan_inf_to_errors": True}}
) as writer:
    
    def safe_write(ws, row, col, value, fmt=None):
        """
        安全寫入 Excel：
        - NaN / INF / -INF 一律寫成空白，避免 xlsxwriter 爆掉
        - 其餘值維持原樣寫入
        """
        try:
            if pd.isna(value):
                ws.write(row, col, "", fmt)
                return
        except Exception:
            pass

        try:
            if np.isinf(value):
                ws.write(row, col, "", fmt)
                return
        except Exception:
            pass

        if fmt is None:
            ws.write(row, col, value)
        else:
            ws.write(row, col, value, fmt)

    workbook = writer.book

    yellow_fmt = workbook.add_format({"bg_color": "#FFFF00"})
    header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1
    })
    number_fmt = workbook.add_format({"num_format": "0.00"})
    title_fmt = workbook.add_format({
        "bold": True,
        "font_size": 14
    })

    colors = [
        "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
        "#5B9BD5", "#70AD47", "#264478", "#9E480E"
    ]

    group_colors = {
        group: colors[i % len(colors)]
        for i, group in enumerate(group_order)
    }

    # -------------------------
    # Input_Copy
    # -------------------------
    df.to_excel(writer, sheet_name="Input_Copy", index=False)

    # -------------------------
    # Error_Log
    # -------------------------
    # 若使用者輸入非空白但無法轉換的值，會記錄在這張表。
    # 沒有異常值時，仍輸出表頭，方便確認檢查完成。
    error_log_df.to_excel(
        writer,
        sheet_name="Error_Log",
        index=False
    )

    error_ws = writer.sheets["Error_Log"]
    for col_num, value in enumerate(error_log_df.columns):
        error_ws.write(0, col_num, value, header_fmt)
    error_ws.set_column(0, 0, 15)
    error_ws.set_column(1, 1, 15)
    error_ws.set_column(2, 2, 12)
    error_ws.set_column(3, 3, 18)
    error_ws.set_column(4, 4, 45)

    # -------------------------
    # Processed_Wide 備份
    # -------------------------
    processed_wide.to_excel(
        writer,
        sheet_name="Processed_Wide",
        index=False
    )

    processed_ws = writer.sheets["Processed_Wide"]

    for col_num, value in enumerate(processed_wide.columns):
        processed_ws.write(0, col_num, value, header_fmt)

    for row in range(1, len(processed_wide) + 1):

        group = processed_wide.iloc[row - 1]["Group"]
        mouse = processed_wide.iloc[row - 1]["Mouse_ID"]

        flag_row = converted_flag[
            (converted_flag["Group"] == group)
            & (converted_flag["Mouse_ID"] == mouse)
        ]

        if len(flag_row) == 0:
            continue

        for col in range(2, len(processed_wide.columns)):

            time_col = processed_wide.columns[col]

            if bool(flag_row.iloc[0][time_col]):

                safe_write(
                    processed_ws,
                    row,
                    col,
                    processed_wide.iloc[row - 1, col],
                    yellow_fmt
                )

    processed_ws.set_column(0, 1, 15)
    processed_ws.set_column(2, len(processed_wide.columns), 12)


    # =========================
    # Mean_SD_AUC Sheet
    # =========================
    ws = workbook.add_worksheet("Mean_SD_AUC")
    writer.sheets["Mean_SD_AUC"] = ws

    ws.write("A1", "Mean SD AUC Report", title_fmt)
    ws.write("A2", "AUC Y-axis Max", header_fmt)
    ws.write("B2", AUC_Y_AXIS_MAX, number_fmt)

    # -------------------------
    # 7-1. 複製 Processed_Wide 到本 Sheet
    # -------------------------
    raw_start_row = 4   # Excel 第5列
    raw_start_col = 0

    for col_num, value in enumerate(processed_wide.columns):
        ws.write(raw_start_row, raw_start_col + col_num, value, header_fmt)

    for r in range(len(processed_wide)):
        for c in range(len(processed_wide.columns)):

            value = processed_wide.iloc[r, c]

            if c >= 2:
                group = processed_wide.iloc[r]["Group"]
                mouse = processed_wide.iloc[r]["Mouse_ID"]
                time_col = processed_wide.columns[c]

                flag_row = converted_flag[
                    (converted_flag["Group"] == group)
                    & (converted_flag["Mouse_ID"] == mouse)
                ]

                if len(flag_row) > 0 and bool(flag_row.iloc[0][time_col]):
                    safe_write(ws, raw_start_row + 1 + r, raw_start_col + c, value, yellow_fmt)
                else:
                    safe_write(ws, raw_start_row + 1 + r, raw_start_col + c, value, number_fmt)
            else:
                safe_write(ws, raw_start_row + 1 + r, raw_start_col + c, value)

    raw_first_data_row = raw_start_row + 1
    raw_last_data_row = raw_start_row + len(processed_wide)

    # 每組在 Mean_SD_AUC 內的資料列範圍
    group_ranges = {}

    for group in group_order:
        idx = processed_wide.index[processed_wide["Group"] == group]

        start_excel_row = raw_start_row + 1 + idx.min() + 1
        end_excel_row = raw_start_row + 1 + idx.max() + 1

        group_ranges[group] = (start_excel_row, end_excel_row)


    # -------------------------
    # 7-2. Mean Table：Excel公式
    # -------------------------
    mean_start_row = raw_last_data_row + 4

    ws.write(mean_start_row, 0, "Time (h)", header_fmt)

    for i, group in enumerate(group_order):
        ws.write(mean_start_row, i + 1, group, header_fmt)

    for r, time in enumerate(time_order):

        excel_row = mean_start_row + 1 + r

        ws.write(excel_row, 0, time / 60, number_fmt)

        raw_col_index = processed_wide.columns.get_loc(time)
        raw_col_letter = xl_col_to_name(raw_col_index)

        for c, group in enumerate(group_order):

            row1, row2 = group_ranges[group]

            formula = (
                f"=AVERAGE("
                f"{raw_col_letter}{row1}:"
                f"{raw_col_letter}{row2})"
            )

            ws.write_formula(
                excel_row,
                c + 1,
                formula,
                number_fmt
            )


    # -------------------------
    # 7-3. SD Table：Excel公式
    # -------------------------
    sd_start_col = len(group_order) + 3

    ws.write(mean_start_row, sd_start_col, "Time (h)", header_fmt)

    for i, group in enumerate(group_order):
        ws.write(mean_start_row, sd_start_col + i + 1, group, header_fmt)

    for r, time in enumerate(time_order):

        excel_row = mean_start_row + 1 + r

        ws.write(excel_row, sd_start_col, time / 60, number_fmt)

        raw_col_index = processed_wide.columns.get_loc(time)
        raw_col_letter = xl_col_to_name(raw_col_index)

        for c, group in enumerate(group_order):

            row1, row2 = group_ranges[group]

            formula = (
                f"=STDEV("
                f"{raw_col_letter}{row1}:"
                f"{raw_col_letter}{row2})"
            )

            ws.write_formula(
                excel_row,
                sd_start_col + c + 1,
                formula,
                number_fmt
            )


    # -------------------------
    # 7-4. AUC by Mouse：Excel公式
    # -------------------------
    auc_start_col = sd_start_col + len(group_order) + 3

    ws.write(mean_start_row, auc_start_col, "Group", header_fmt)
    ws.write(mean_start_row, auc_start_col + 1, "Mouse_ID", header_fmt)
    ws.write(mean_start_row, auc_start_col + 2, "AUC", header_fmt)

    time_data_cols = list(processed_wide.columns[2:])

    first_time_col = 2
    last_time_col = len(processed_wide.columns) - 1

    for r in range(len(processed_wide)):

        excel_row = mean_start_row + 1 + r
        raw_excel_row = raw_start_row + 1 + r + 1

        ws.write(excel_row, auc_start_col, processed_wide.iloc[r]["Group"])
        ws.write(excel_row, auc_start_col + 1, processed_wide.iloc[r]["Mouse_ID"])

        auc_parts = []

        for i in range(first_time_col, last_time_col):

            col1 = xl_col_to_name(i)
            col2 = xl_col_to_name(i + 1)

            t1 = processed_wide.columns[i]
            t2 = processed_wide.columns[i + 1]

            delta_h = (t2 - t1) / 60

            part = (
                f"(({col1}{raw_excel_row}+{col2}{raw_excel_row})/2)"
                f"*{delta_h}"
            )

            auc_parts.append(part)

        auc_formula = "=" + "+".join(auc_parts)

        ws.write_formula(
            excel_row,
            auc_start_col + 2,
            auc_formula,
            number_fmt
        )


    # -------------------------
    # 7-5. AUC Summary：Excel公式
    # -------------------------
    auc_summary_start_row = mean_start_row + len(processed_wide) + 4

    ws.write(auc_summary_start_row, auc_start_col, "Group", header_fmt)
    ws.write(auc_summary_start_row, auc_start_col + 1, "AUC Mean", header_fmt)
    ws.write(auc_summary_start_row, auc_start_col + 2, "AUC SD", header_fmt)

    auc_group_ranges = {}

    for group in group_order:
        idx = processed_wide.index[processed_wide["Group"] == group]

        start_excel_row = mean_start_row + 1 + idx.min() + 1
        end_excel_row = mean_start_row + 1 + idx.max() + 1

        auc_group_ranges[group] = (start_excel_row, end_excel_row)

    auc_col_letter = xl_col_to_name(auc_start_col + 2)

    for r, group in enumerate(group_order):

        excel_row = auc_summary_start_row + 1 + r

        row1, row2 = auc_group_ranges[group]

        ws.write(excel_row, auc_start_col, group)

        ws.write_formula(
            excel_row,
            auc_start_col + 1,
            f"=AVERAGE({auc_col_letter}{row1}:{auc_col_letter}{row2})",
            number_fmt
        )

        ws.write_formula(
            excel_row,
            auc_start_col + 2,
            f"=STDEV({auc_col_letter}{row1}:{auc_col_letter}{row2})",
            number_fmt
        )


    # -------------------------
    # 7-6. Mean ± SD 折線圖
    # -------------------------
    line_chart = workbook.add_chart({"type": "line"})

    first_data_row = mean_start_row + 1
    last_data_row = mean_start_row + len(time_order)

    for i, group in enumerate(group_order):

        line_chart.add_series({
            "name":       ["Mean_SD_AUC", mean_start_row, i + 1],
            "categories": ["Mean_SD_AUC", first_data_row, 0, last_data_row, 0],
            "values":     ["Mean_SD_AUC", first_data_row, i + 1, last_data_row, i + 1],
            "y_error_bars": {
                "type": "custom",
                "plus_values": [
                    "Mean_SD_AUC",
                    first_data_row,
                    sd_start_col + i + 1,
                    last_data_row,
                    sd_start_col + i + 1
                ],
                "minus_values": [
                    "Mean_SD_AUC",
                    first_data_row,
                    sd_start_col + i + 1,
                    last_data_row,
                    sd_start_col + i + 1
                ],
            },
            "marker": {
                "type": "circle",
                "size": 6,
                "border": {"color": group_colors[group]},
                "fill": {"color": group_colors[group]},
            },
            "line": {
                "color": group_colors[group],
                "width": 2
            },
        })

    line_chart.set_title({"name": "IPGTT Mean ± SD"})
    line_chart.set_x_axis({"name": "Time (h)"})
    line_chart.set_y_axis({"name": "Blood glucose (mg/dL)"})
    line_chart.set_legend({"position": "bottom"})
    line_chart.set_size({"width": 720, "height": 420})

    chart_start_row = mean_start_row + len(time_order) + 5
    ws.insert_chart(chart_start_row, 0, line_chart)


    # -------------------------
    # 7-7. AUC Mean ± SD 長條圖
    # -------------------------
    bar_chart = workbook.add_chart({"type": "column"})

    first_auc_summary_row = auc_summary_start_row + 1
    last_auc_summary_row = auc_summary_start_row + len(group_order)

    
    bar_chart.add_series({
        "name": "AUC Mean ± SD",
    
        "categories": [
            "Mean_SD_AUC",
            first_auc_summary_row,
            auc_start_col,
            last_auc_summary_row,
            auc_start_col
        ],
    
        "values": [
            "Mean_SD_AUC",
            first_auc_summary_row,
            auc_start_col + 1,
            last_auc_summary_row,
            auc_start_col + 1
        ],
    
        "y_error_bars": {
            "type": "custom",
            "plus_values": [
                "Mean_SD_AUC",
                first_auc_summary_row,
                auc_start_col + 2,
                last_auc_summary_row,
                auc_start_col + 2
            ],
            "minus_values": [
                "Mean_SD_AUC",
                first_auc_summary_row,
                auc_start_col + 2,
                last_auc_summary_row,
                auc_start_col + 2
            ]
        },
    
        "points": [
            {"fill": {"color": group_colors[group]}}
            for group in group_order
        ],
    
        "data_labels": {
            "value": True,
            "position": "outside_end"
        }
    })


    bar_chart.set_title({"name": "AUC Mean ± SD"})
    bar_chart.set_x_axis({"name": "Group"})
    bar_chart.set_y_axis({
            "name": "AUC (mg/dL·h)",
            "min": 0,
            "max": AUC_Y_AXIS_MAX
        })
    
    bar_chart.set_legend({"none": True})

    bar_chart.set_size({
            "width": 550,
            "height": 360
        })
    
    ws.insert_chart(chart_start_row, 8, bar_chart)
    
    ws.set_column(0, auc_start_col + 3, 13)
    
    # =========================
    # Mean_SE_AUC Sheet
    # =========================
    ws_se = workbook.add_worksheet("Mean_SE_AUC")
    writer.sheets["Mean_SE_AUC"] = ws_se

    ws_se.write("A1", "Mean SE AUC Report", title_fmt)
    ws_se.write("A2", "AUC Y-axis Max", header_fmt)
    ws_se.write("B2", AUC_Y_AXIS_MAX, number_fmt)

    # -------------------------
    # 8-1. 複製 Processed_Wide 到本 Sheet
    # -------------------------
    raw_start_row_se = 4
    raw_start_col_se = 0

    for col_num, value in enumerate(processed_wide.columns):
        ws_se.write(raw_start_row_se, raw_start_col_se + col_num, value, header_fmt)

    for r in range(len(processed_wide)):
        for c in range(len(processed_wide.columns)):

            value = processed_wide.iloc[r, c]

            if c >= 2:
                group = processed_wide.iloc[r]["Group"]
                mouse = processed_wide.iloc[r]["Mouse_ID"]
                time_col = processed_wide.columns[c]

                flag_row = converted_flag[
                    (converted_flag["Group"] == group)
                    & (converted_flag["Mouse_ID"] == mouse)
                ]

                if len(flag_row) > 0 and bool(flag_row.iloc[0][time_col]):
                    safe_write(ws_se, raw_start_row_se + 1 + r, raw_start_col_se + c, value, yellow_fmt)
                else:
                    safe_write(ws_se, raw_start_row_se + 1 + r, raw_start_col_se + c, value, number_fmt)
            else:
                safe_write(ws_se, raw_start_row_se + 1 + r, raw_start_col_se + c, value)

    raw_last_data_row_se = raw_start_row_se + len(processed_wide)

    # 每組在 Mean_SE_AUC 內的資料列範圍
    group_ranges_se = {}

    for group in group_order:
        idx = processed_wide.index[processed_wide["Group"] == group]

        start_excel_row = raw_start_row_se + 1 + idx.min() + 1
        end_excel_row = raw_start_row_se + 1 + idx.max() + 1

        group_ranges_se[group] = (start_excel_row, end_excel_row)

    # -------------------------
    # 8-2. Mean Table：Excel公式
    # -------------------------
    mean_start_row_se = raw_last_data_row_se + 4

    ws_se.write(mean_start_row_se, 0, "Time (h)", header_fmt)

    for i, group in enumerate(group_order):
        ws_se.write(mean_start_row_se, i + 1, group, header_fmt)

    for r, time in enumerate(time_order):

        excel_row = mean_start_row_se + 1 + r
        ws_se.write(excel_row, 0, time / 60, number_fmt)

        raw_col_index = processed_wide.columns.get_loc(time)
        raw_col_letter = xl_col_to_name(raw_col_index)

        for c, group in enumerate(group_order):

            row1, row2 = group_ranges_se[group]

            formula = (
                f"=AVERAGE("
                f"{raw_col_letter}{row1}:"
                f"{raw_col_letter}{row2})"
            )

            ws_se.write_formula(
                excel_row,
                c + 1,
                formula,
                number_fmt
            )

    # -------------------------
    # 8-3. SE Table：Excel公式
    # -------------------------
    se_start_col = len(group_order) + 3

    ws_se.write(mean_start_row_se, se_start_col, "Time (h)", header_fmt)

    for i, group in enumerate(group_order):
        ws_se.write(mean_start_row_se, se_start_col + i + 1, group, header_fmt)

    for r, time in enumerate(time_order):

        excel_row = mean_start_row_se + 1 + r
        ws_se.write(excel_row, se_start_col, time / 60, number_fmt)

        raw_col_index = processed_wide.columns.get_loc(time)
        raw_col_letter = xl_col_to_name(raw_col_index)

        for c, group in enumerate(group_order):

            row1, row2 = group_ranges_se[group]

            formula = (
                f"=STDEV("
                f"{raw_col_letter}{row1}:"
                f"{raw_col_letter}{row2})"
                f"/SQRT(COUNT("
                f"{raw_col_letter}{row1}:"
                f"{raw_col_letter}{row2}))"
            )

            ws_se.write_formula(
                excel_row,
                se_start_col + c + 1,
                formula,
                number_fmt
            )

    # -------------------------
    # 8-4. AUC by Mouse：Excel公式
    # -------------------------
    auc_start_col_se = se_start_col + len(group_order) + 3

    ws_se.write(mean_start_row_se, auc_start_col_se, "Group", header_fmt)
    ws_se.write(mean_start_row_se, auc_start_col_se + 1, "Mouse_ID", header_fmt)
    ws_se.write(mean_start_row_se, auc_start_col_se + 2, "AUC", header_fmt)

    first_time_col = 2
    last_time_col = len(processed_wide.columns) - 1

    for r in range(len(processed_wide)):

        excel_row = mean_start_row_se + 1 + r
        raw_excel_row = raw_start_row_se + 1 + r + 1

        ws_se.write(excel_row, auc_start_col_se, processed_wide.iloc[r]["Group"])
        ws_se.write(excel_row, auc_start_col_se + 1, processed_wide.iloc[r]["Mouse_ID"])

        auc_parts = []

        for i in range(first_time_col, last_time_col):

            col1 = xl_col_to_name(i)
            col2 = xl_col_to_name(i + 1)

            t1 = processed_wide.columns[i]
            t2 = processed_wide.columns[i + 1]

            delta_h = (t2 - t1) / 60

            part = (
                f"(({col1}{raw_excel_row}+{col2}{raw_excel_row})/2)"
                f"*{delta_h}"
            )

            auc_parts.append(part)

        auc_formula = "=" + "+".join(auc_parts)

        ws_se.write_formula(
            excel_row,
            auc_start_col_se + 2,
            auc_formula,
            number_fmt
        )

    # -------------------------
    # 8-5. AUC Summary：Excel公式
    # -------------------------
    auc_summary_start_row_se = mean_start_row_se + len(processed_wide) + 4

    ws_se.write(auc_summary_start_row_se, auc_start_col_se, "Group", header_fmt)
    ws_se.write(auc_summary_start_row_se, auc_start_col_se + 1, "AUC Mean", header_fmt)
    ws_se.write(auc_summary_start_row_se, auc_start_col_se + 2, "AUC SE", header_fmt)

    auc_group_ranges_se = {}

    for group in group_order:
        idx = processed_wide.index[processed_wide["Group"] == group]

        start_excel_row = mean_start_row_se + 1 + idx.min() + 1
        end_excel_row = mean_start_row_se + 1 + idx.max() + 1

        auc_group_ranges_se[group] = (start_excel_row, end_excel_row)

    auc_col_letter_se = xl_col_to_name(auc_start_col_se + 2)

    for r, group in enumerate(group_order):

        excel_row = auc_summary_start_row_se + 1 + r

        row1, row2 = auc_group_ranges_se[group]

        ws_se.write(excel_row, auc_start_col_se, group)

        ws_se.write_formula(
            excel_row,
            auc_start_col_se + 1,
            f"=AVERAGE({auc_col_letter_se}{row1}:{auc_col_letter_se}{row2})",
            number_fmt
        )

        ws_se.write_formula(
            excel_row,
            auc_start_col_se + 2,
            f"=STDEV({auc_col_letter_se}{row1}:{auc_col_letter_se}{row2})/SQRT(COUNT({auc_col_letter_se}{row1}:{auc_col_letter_se}{row2}))",
            number_fmt
        )

    # -------------------------
    # 8-6. Mean ± SE 折線圖
    # -------------------------
    line_chart_se = workbook.add_chart({"type": "line"})

    first_data_row_se = mean_start_row_se + 1
    last_data_row_se = mean_start_row_se + len(time_order)

    for i, group in enumerate(group_order):

        line_chart_se.add_series({
            "name":       ["Mean_SE_AUC", mean_start_row_se, i + 1],
            "categories": ["Mean_SE_AUC", first_data_row_se, 0, last_data_row_se, 0],
            "values":     ["Mean_SE_AUC", first_data_row_se, i + 1, last_data_row_se, i + 1],
            "y_error_bars": {
                "type": "custom",
                "plus_values": [
                    "Mean_SE_AUC",
                    first_data_row_se,
                    se_start_col + i + 1,
                    last_data_row_se,
                    se_start_col + i + 1
                ],
                "minus_values": [
                    "Mean_SE_AUC",
                    first_data_row_se,
                    se_start_col + i + 1,
                    last_data_row_se,
                    se_start_col + i + 1
                ],
            },
            "marker": {
                "type": "circle",
                "size": 6,
                "border": {"color": group_colors[group]},
                "fill": {"color": group_colors[group]},
            },
            "line": {
                "color": group_colors[group],
                "width": 2
            },
        })

    line_chart_se.set_title({"name": "IPGTT Mean ± SE"})
    line_chart_se.set_x_axis({"name": "Time (h)"})
    line_chart_se.set_y_axis({"name": "Blood glucose (mg/dL)"})
    line_chart_se.set_legend({"position": "bottom"})
    line_chart_se.set_size({"width": 720, "height": 420})

    chart_start_row_se = mean_start_row_se + len(time_order) + 5
    ws_se.insert_chart(chart_start_row_se, 0, line_chart_se)

    # -------------------------
    # 8-7. AUC Mean ± SE 長條圖
    # -------------------------
    bar_chart_se = workbook.add_chart({"type": "column"})

    first_auc_summary_row_se = auc_summary_start_row_se + 1
    last_auc_summary_row_se = auc_summary_start_row_se + len(group_order)

    bar_chart_se.add_series({
        "name": "AUC Mean ± SE",

        "categories": [
            "Mean_SE_AUC",
            first_auc_summary_row_se,
            auc_start_col_se,
            last_auc_summary_row_se,
            auc_start_col_se
        ],

        "values": [
            "Mean_SE_AUC",
            first_auc_summary_row_se,
            auc_start_col_se + 1,
            last_auc_summary_row_se,
            auc_start_col_se + 1
        ],

        "y_error_bars": {
            "type": "custom",
            "plus_values": [
                "Mean_SE_AUC",
                first_auc_summary_row_se,
                auc_start_col_se + 2,
                last_auc_summary_row_se,
                auc_start_col_se + 2
            ],
            "minus_values": [
                "Mean_SE_AUC",
                first_auc_summary_row_se,
                auc_start_col_se + 2,
                last_auc_summary_row_se,
                auc_start_col_se + 2
            ]
        },

        "points": [
            {"fill": {"color": group_colors[group]}}
            for group in group_order
        ],

        "data_labels": {
            "value": True,
            "position": "outside_end"
        }
    })

    bar_chart_se.set_title({"name": "AUC Mean ± SE"})
    bar_chart_se.set_x_axis({"name": "Group"})
    bar_chart_se.set_y_axis({
        "name": "AUC (mg/dL·h)",
        "min": 0,
        "max": AUC_Y_AXIS_MAX
    })

    bar_chart_se.set_legend({"none": True})

    bar_chart_se.set_size({
        "width": 550,
        "height": 360
    })

    ws_se.insert_chart(chart_start_row_se, 8, bar_chart_se)

    ws_se.set_column(0, auc_start_col_se + 3, 13)


print("完成！")
print(f"輸出檔案：{output_path}")

os.startfile(output_path)




import os
import shutil
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import PatternFill, Font, Alignment

# 班別長條的統一顏色，對應網頁預覽的 --series-1（前後段班別性質上無差異，不用顏色區分）
DUR_SERIES_COLOR = "2A78D6"

def _get_n_shifts(row_meanings):
    n_shifts = 0
    for m in row_meanings:
        if m["type"] == "shift":
            n_shifts = max(n_shifts, m.get("index", 0) + 1)
        elif m["type"] in ["shift_start", "shift_end"]:
            n_shifts = max(n_shifts, m.get("index", 0) + 1)
        elif m["type"] == "shift_string":
            n_shifts = max(n_shifts, 2)
    return max(1, n_shifts)

def generate_xlsx(employees, output_path, template_config):
    """
    Generates a monthly XLSX workbook. Each day is a sheet containing the schedule table and Gantt chart.
    """
    display_config = template_config.get("display") or {}
    axis_start = display_config.get("axis_start", 8)
    axis_end = display_config.get("axis_end", 24)

    wb = Workbook()
    wb.remove(wb.active) # Remove default sheet

    for day in range(1, 32):
        ws = wb.create_sheet(title=f"{day}")
        
        # 1. Headers
        headers = ["員工姓名"]
        block_config = template_config["block"]
        row_meanings = block_config["row_meanings"]
        
        # Count shifts
        n_shifts = _get_n_shifts(row_meanings)
        for s_idx in range(1, n_shifts + 1):
            headers.extend([f"班別{s_idx}起", f"班別{s_idx}訖"])
            
        ws.append(headers)
        
        # Style headers
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(name="Microsoft JhengHei", size=11, bold=True, color="FFFFFF")
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # 2. Populate employee raw shifts
        time_fmt = "h:mm"
        for idx, emp in enumerate(employees):
            r = idx + 2
            row_data = [emp["name"]]
            
            day_info = emp["days"][day]
            for s_idx in range(n_shifts):
                shift = day_info["shifts"][s_idx] if s_idx < len(day_info["shifts"]) else {"start": None, "end": None}
                
                start_frac = (shift["start"] / 24) if shift["start"] is not None else None
                end_frac = (shift["end"] / 24) if shift["end"] is not None else None
                row_data.extend([start_frac, end_frac])
                
            ws.append(row_data)
            
            # Format times
            for c_idx in range(2, 2 + 2 * n_shifts):
                cell = ws.cell(row=r, column=c_idx)
                cell.number_format = time_fmt
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            ws.cell(row=r, column=1).font = Font(name="Microsoft JhengHei", size=10, bold=True)
            
        # 3. Build helper columns for Gantt chart
        start_helper_col = 2 + 2 * n_shifts
        helper_headers = ["base"]
        for s_idx in range(1, n_shifts + 1):
            helper_headers.append(f"dur{s_idx}")
            if s_idx < n_shifts:
                helper_headers.append(f"gap{s_idx}")
                
        for h_idx, h_name in enumerate(helper_headers):
            ws.cell(row=1, column=start_helper_col + h_idx, value=h_name)

        # 輔助欄（base/dur/gap）僅供甘特圖繪製使用，對使用者無意義，隱藏但保留資料。
        for h_idx in range(len(helper_headers)):
            ws.column_dimensions[_col_letter(start_helper_col + h_idx)].hidden = True

        day_start = axis_start / 24
        for idx, emp in enumerate(employees):
            r = idx + 2
            day_info = emp["days"][day]
            
            active_shifts = []
            for s in day_info["shifts"]:
                if s["start"] is not None and s["end"] is not None:
                    active_shifts.append(s)
            
            active_shifts.sort(key=lambda x: x["start"])
            
            segments = []
            if not active_shifts:
                base = day_start
                durations_and_gaps = [0.0] * (2 * n_shifts - 1)
            else:
                base = active_shifts[0]["start"] / 24
                durations_and_gaps = [(active_shifts[0]["end"] - active_shifts[0]["start"]) / 24]
                
                for s_idx in range(1, len(active_shifts)):
                    gap = (active_shifts[s_idx]["start"] - active_shifts[s_idx - 1]["end"]) / 24
                    dur = (active_shifts[s_idx]["end"] - active_shifts[s_idx]["start"]) / 24
                    durations_and_gaps.extend([gap, dur])
                
                needed_segments = 2 * n_shifts - 1
                if len(durations_and_gaps) < needed_segments:
                    durations_and_gaps.extend([0.0] * (needed_segments - len(durations_and_gaps)))
            
            ws.cell(row=r, column=start_helper_col, value=base)
            for s_idx, val in enumerate(durations_and_gaps):
                ws.cell(row=r, column=start_helper_col + 1 + s_idx, value=val)
                
        add_xlsx_gantt_chart(ws, len(employees), n_shifts, start_helper_col, axis_start, axis_end)

    wb.save(output_path)
    print(f"Generated XLSX schedule: {output_path}")

def add_xlsx_gantt_chart(ws, n_employees, n_shifts, start_helper_col, axis_start=8, axis_end=24):
    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = f"{ws.title}日 排班長條圖"
    # 輔助欄（base/dur/gap）在上面被隱藏，但圖表仍需讀取其資料繪圖：
    # Excel 圖表預設 plotVisOnly=true（只畫可見儲存格），必須關閉。
    chart.visible_cells_only = False
    chart.legend = None

    last_row = n_employees + 1

    total_series = 2 * n_shifts
    for s_idx in range(total_series):
        col = start_helper_col + s_idx
        ref = Reference(ws, min_col=col, min_row=1, max_row=last_row)
        chart.add_data(ref, titles_from_data=True)

    # set_categories() 必須在 add_data() 之後呼叫：呼叫當下 chart.series 是空的，
    # 姓名分類永遠不會被寫進任何 series，長條圖左側會變成 Excel 自動編號 1..N。
    cats = Reference(ws, min_col=1, min_row=2, max_row=last_row)
    chart.set_categories(cats)

    # 姓名軸（x_axis，類別軸）：預設由下往上排列，reverse 成由上往下。
    chart.x_axis.title = None
    chart.x_axis.scaling.orientation = "maxMin"
    # 類別軸標籤間隔預設是 Excel/LibreOffice 的「自動」，員工數一多、圖表高度
    # 顯得不夠時，渲染軟體會自動跳過部分姓名標籤（並非資料遺漏，只是沒顯示）。
    # 強制間隔為 1，確保每一位員工的姓名都會顯示。
    chart.x_axis.tickLblSkip = 1
    chart.x_axis.tickMarkSkip = 1

    # 實測發現 LibreOffice 即使 tickLblSkip=1 仍會依「圖表高度 ÷ 員工數」自行
    # 算間隔並跳過標籤，需要圖表本身夠高才會真的顯示全部姓名。以 10 人的高度
    # 為下限（避免人少時圖表過矮），不設上限、依員工數等比放大。
    chart.height = max(10, n_employees) * 0.8

    # 時間軸（y_axis，數值軸）：原本誤設定在 x_axis 上，導致真正的數值軸完全沒有
    # 格式化（顯示 0~1 的原始小數）；改回設在 y_axis，並移到圖表上方、
    # 刻度間隔比照網頁版每小時一格。
    # 用 "[h]:mm"（經過時間格式）而非 "h:mm"：後者在數值剛好等於 1.0（=24:00）
    # 時，Excel 會把它當成隔天 0:00 顯示，導致座標軸標籤回繞成 "0:00"。
    chart.y_axis.title = "時間"
    chart.y_axis.number_format = "[h]:mm"
    chart.y_axis.scaling.min = axis_start / 24
    chart.y_axis.scaling.max = axis_end / 24
    chart.y_axis.majorUnit = 1 / 24
    chart.y_axis.crosses = "max"

    chart.series[0].graphicalProperties.noFill = True

    # 前後段班別性質上無差異，故所有班別長條統一同一色（與網頁預覽 --series-1 一致），
    # 不用 Excel 預設的逐 series 自動配色。
    for s_idx in range(1, total_series):
        is_dur = (s_idx % 2 == 1)
        if is_dur:
            shift_num = (s_idx // 2) + 1
            chart.series[s_idx].tx = SeriesLabel(v=f"班別 {shift_num}")
            chart.series[s_idx].graphicalProperties.solidFill = DUR_SERIES_COLOR
        else:
            chart.series[s_idx].graphicalProperties.noFill = True

    chart.width = 24
    
    chart_col = chr(65 + start_helper_col + total_series + 1)
    ws.add_chart(chart, f"{chart_col}2")


def _col_letter(col_idx):
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


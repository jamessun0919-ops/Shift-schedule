import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import PatternFill, Font, Alignment

# 班別長條的統一顏色，對應網頁預覽的 --series-1（前後段班別性質上無差異，不用顏色區分）
DUR_SERIES_COLOR = "2A78D6"

NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "chart": "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
}

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

def _tag(ns_key, local):
    return f"{{{NS[ns_key]}}}{local}"

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
            
    chart.height = max(8, n_employees * 0.6)
    chart.width = 24
    
    chart_col = chr(65 + start_helper_col + total_series + 1)
    ws.add_chart(chart, f"{chart_col}2")


def _create_string_cell(val):
    cell = ET.Element(_tag("table", "table-cell"), {
        _tag("office", "value-type"): "string"
    })
    p = ET.SubElement(cell, _tag("text", "p"))
    p.text = str(val)
    return cell

def _create_float_cell(val):
    cell = ET.Element(_tag("table", "table-cell"), {
        _tag("office", "value-type"): "float",
        _tag("office", "value"): str(val)
    })
    p = ET.SubElement(cell, _tag("text", "p"))
    p.text = f"{val:.1f}" if isinstance(val, float) else str(val)
    return cell

def _create_empty_cell():
    return ET.Element(_tag("table", "table-cell"))

def _col_letter(col_idx):
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def generate_ods(employees, output_path, template_config, template_ods_path="白霧.ods"):
    """
    Generates a monthly ODS file by taking a template ODS, clearing day sheets 1-31, 
    and writing the parsed data and helper columns.
    It dynamically updates the corresponding chart objects.
    """
    print(f"Generating ODS using template: {template_ods_path}")
    temp_zip = output_path + ".tmp"
    
    block_config = template_config["block"]
    row_meanings = block_config["row_meanings"]
    n_shifts = _get_n_shifts(row_meanings)

    display_config = template_config.get("display") or {}
    axis_start = display_config.get("axis_start", 8)
    axis_end = display_config.get("axis_end", 24)

    # Pre-parse content.xml to identify sheet to chart object mapping
    chart_mappings = {} # day_str -> chart_object_dir_name
    with zipfile.ZipFile(template_ods_path, "r") as z:
        content_xml = z.read("content.xml")
        root = ET.fromstring(content_xml)
        
    for t in root.iter(_tag("table", "table")):
        sheet_name = t.get(_tag("table", "name"))
        if sheet_name in [str(d) for d in range(1, 32)]:
            draw_obj = t.find(f".//{_tag('draw', 'object')}")
            if draw_obj is not None:
                href = draw_obj.get(f"{{{NS['xlink']}}}href")
                if href:
                    chart_mappings[sheet_name] = href.lstrip("./")

    # Generate day by day data
    day_sheet_data = {}
    for day in range(1, 32):
        sheet_rows = []
        
        headers = ["員工姓名"]
        for s_idx in range(1, n_shifts + 1):
            headers.extend([f"班別{s_idx}起", f"班別{s_idx}訖"])
            
        start_helper_col = 2 + 2 * n_shifts
        helper_headers = ["base"]
        for s_idx in range(1, n_shifts + 1):
            helper_headers.append(f"dur{s_idx}")
            if s_idx < n_shifts:
                helper_headers.append(f"gap{s_idx}")
        headers.extend(helper_headers)
        sheet_rows.append(headers)
        
        day_start = float(axis_start)
        for emp in employees:
            row_data = [emp["name"]]
            day_info = emp["days"][day]
            
            active_shifts = []
            for s_idx in range(n_shifts):
                shift = day_info["shifts"][s_idx] if s_idx < len(day_info["shifts"]) else {"start": None, "end": None}
                row_data.extend([shift["start"], shift["end"]])
                if shift["start"] is not None and shift["end"] is not None:
                    active_shifts.append(shift)
                    
            active_shifts.sort(key=lambda x: x["start"])
            
            # Gantt segments
            if not active_shifts:
                base = day_start
                durations_and_gaps = [0.0] * (2 * n_shifts - 1)
            else:
                base = active_shifts[0]["start"]
                durations_and_gaps = [active_shifts[0]["end"] - active_shifts[0]["start"]]
                
                for s_idx in range(1, len(active_shifts)):
                    gap = active_shifts[s_idx]["start"] - active_shifts[s_idx - 1]["end"]
                    dur = active_shifts[s_idx]["end"] - active_shifts[s_idx]["start"]
                    durations_and_gaps.extend([gap, dur])
                
                needed_segments = 2 * n_shifts - 1
                if len(durations_and_gaps) < needed_segments:
                    durations_and_gaps.extend([0.0] * (needed_segments - len(durations_and_gaps)))
                    
            row_data.append(base)
            row_data.extend(durations_and_gaps)
            sheet_rows.append(row_data)
            
        day_sheet_data[str(day)] = sheet_rows

    # Copy template and write modified contents
    with zipfile.ZipFile(template_ods_path, "r") as z_in:
        with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as z_out:
            # First write mimetype uncompressed
            try:
                z_out.writestr("mimetype", z_in.read("mimetype"), zipfile.ZIP_STORED)
            except KeyError:
                pass

            for item in z_in.infolist():
                if item.filename == "mimetype":
                    continue
                
                content = z_in.read(item.filename)
                
                if item.filename == "content.xml":
                    # Rewrite the sheet cells
                    root = ET.fromstring(content)
                    for t in root.iter(_tag("table", "table")):
                        sheet_name = t.get(_tag("table", "name"))
                        if sheet_name in day_sheet_data:
                            # Rewrite rows in table
                            data_rows = day_sheet_data[sheet_name]
                            
                            # Keep draw:frame (the chart object frame) so the chart is preserved
                            draw_frame = t.find(f".//{_tag('draw', 'frame')}")
                            
                            # Remove all row children of the table
                            for child in list(t):
                                if child.tag == _tag("table", "table-row") or child.tag == _tag("table", "table-column"):
                                    t.remove(child)
                                    
                            # Add basic column definitions
                            for _ in range(35): # 35 columns
                                ET.SubElement(t, _tag("table", "table-column"))
                                
                            # Write new data rows
                            for r_idx, row_values in enumerate(data_rows):
                                row_el = ET.SubElement(t, _tag("table", "table-row"))
                                for c_idx, val in enumerate(row_values):
                                    if val is None or val == "":
                                        row_el.append(_create_empty_cell())
                                    elif isinstance(val, (int, float)):
                                        row_el.append(_create_float_cell(val))
                                    else:
                                        row_el.append(_create_string_cell(val))
                                        
                            # Append chart frame if present to keep chart rendering
                            if draw_frame is not None:
                                t.append(draw_frame)
                                
                    for prefix, uri in NS.items():
                        if prefix != "xlink":
                            ET.register_namespace(prefix, uri)
                    ET.register_namespace("xlink", NS["xlink"])
                    content = ET.tostring(root, encoding="utf-8")
                    
                # Rewrite chart objects content.xml
                elif item.filename.startswith("Object ") and item.filename.endswith("/content.xml"):
                    chart_dir = item.filename.split("/")[0]
                    matching_day = None
                    for d_str, c_dir in chart_mappings.items():
                        if c_dir == chart_dir:
                            matching_day = d_str
                            break
                            
                    if matching_day and matching_day in day_sheet_data:
                        data_rows = day_sheet_data[matching_day]
                        root = ET.fromstring(content)
                        
                        series_list = list(root.iter(_tag("chart", "series")))
                        last_row = len(data_rows)
                        
                        start_helper = 2 + 2 * n_shifts
                        
                        cats = root.find(f".//{_tag('chart', 'categories')}")
                        if cats is not None:
                            cats.set(_tag("table", "cell-range-address"), f"'{matching_day}'.A2:'{matching_day}'.A{last_row}")
                            
                        for s_idx, series in enumerate(series_list):
                            col_idx = start_helper + s_idx
                            col_let = _col_letter(col_idx)
                            
                            val_range = f"'{matching_day}'.{col_let}2:'{matching_day}'.{col_let}{last_row}"
                            label_range = f"'{matching_day}'.{col_let}1:'{matching_day}'.{col_let}1"
                            
                            series.set(_tag("chart", "values-cell-range-address"), val_range)
                            series.set(_tag("chart", "label-cell-address"), label_range)

                        # 座標軸範圍：y 軸（時間值軸，非分類軸）的最小/最大值寫在
                        # 它引用的 style:style 樣式上（純數字座標軸，無 xlsx 時間格式
                        # 回繞問題，可直接寫入 axis_end，不需要額外修正）。
                        axis_style_name = None
                        for axis_el in root.iter(_tag("chart", "axis")):
                            if axis_el.get(_tag("chart", "dimension")) == "y":
                                axis_style_name = axis_el.get(_tag("chart", "style-name"))
                                break
                        if axis_style_name:
                            for style_el in root.iter(_tag("style", "style")):
                                if style_el.get(_tag("style", "name")) == axis_style_name:
                                    props = style_el.find(_tag("style", "chart-properties"))
                                    if props is not None:
                                        props.set(_tag("chart", "minimum"), str(axis_start))
                                        props.set(_tag("chart", "maximum"), str(axis_end))
                                    break

                        local_table = root.find(f".//{_tag('table', 'table')}")
                        if local_table is not None:
                            for child in list(local_table):
                                local_table.remove(child)
                                
                            header_row_el = ET.SubElement(local_table, _tag("table", "table-row"))
                            header_row_el.append(_create_string_cell(""))
                            
                            header_row_el.append(_create_string_cell("base"))
                            for s_idx in range(1, n_shifts + 1):
                                header_row_el.append(_create_string_cell(f"dur{s_idx}"))
                                if s_idx < n_shifts:
                                    header_row_el.append(_create_string_cell(f"gap{s_idx}"))
                                    
                            for r_idx in range(1, len(data_rows)):
                                r_val = data_rows[r_idx]
                                row_el = ET.SubElement(local_table, _tag("table", "table-row"))
                                row_el.append(_create_string_cell(r_val[0]))
                                helper_vals = r_val[start_helper - 1:]
                                for val in helper_vals:
                                    row_el.append(_create_float_cell(val))
                                    
                        for prefix, uri in NS.items():
                            if prefix != "xlink":
                                ET.register_namespace(prefix, uri)
                        ET.register_namespace("xlink", NS["xlink"])
                        content = ET.tostring(root, encoding="utf-8")
                        
                z_out.writestr(item, content)
                
    os.replace(temp_zip, output_path)
    print(f"Generated ODS schedule: {output_path}")

import os
import re
import math
import datetime
from collections import Counter
from backend.parser import read_ods_rows, read_xlsx_rows

def is_date_or_day_num(val):
    if val is None:
        return False
    if isinstance(val, (datetime.datetime, datetime.date)):
        return True
    if isinstance(val, (int, float)):
        if 1.0 <= val <= 31.0 and val == int(val):
            return True
        return False
        
    val_str = str(val).strip()
    if not val_str:
        return False
        
    # Check if integer or float-like integer between 1 and 31 (representing day number)
    try:
        fval = float(val_str)
        if 1.0 <= fval <= 31.0 and fval == int(fval):
            return True
    except ValueError:
        pass
        
    # Match strings like "7月1日", "1/1", "11/30"
    if re.match(r"^\d{1,2}月\d{1,2}日$", val_str):
        return True
    if re.match(r"^\d{1,2}/\d{1,2}$", val_str):
        return True

    return False

def guess_month(rows, header_row_idx, first_day_col, cols_per_day, max_days=31):
    """
    從日期列的儲存格內容猜測班表所屬月份，僅供圖表/預覽標題顯示用（不影響解析）。
    純數字日期欄位（如只有 1、2、3...）沒有月份資訊，會回傳 None，呼叫端需自行
    fallback 為不顯示月份。刻意不從 parse_schedule 每次重新呼叫外的其他地方快取
    這個結果——同一份範本可能被套用在不同月份的檔案上，月份必須每次跟著實際上傳
    的檔案重新判斷，不能存進範本裡變成固定值。
    """
    if header_row_idx < 0 or header_row_idx >= len(rows):
        return None
    date_row = rows[header_row_idx]

    months = []
    for d in range(max_days):
        col = first_day_col + d * cols_per_day
        val = date_row.get(col)
        if val is None:
            continue
        if isinstance(val, (datetime.datetime, datetime.date)):
            months.append(val.month)
            continue
        m = re.match(r"^(\d{1,2})月\d{1,2}日$", str(val).strip())
        if m:
            months.append(int(m.group(1)))

    if not months:
        return None

    best_month, count = Counter(months).most_common(1)[0]
    if count / len(months) >= 0.5:
        return best_month
    return None

_HEADER_KEYWORDS = [
    "姓名", "員工姓名", "name", "employee", "職稱", "工號", "員工編號", "員編",
    "排班", "月份", "日期", "星期", "名字", "人數", "工時", "累計", "加總",
    "總計", "合計", "營業額", "成本", "比例", "餐期", "值班", "時數",
    "說明", "使用", "注意事項", "注意", "提醒", "行事曆", "公告", "規則", "規定"
]

def is_header_keyword(val_str):
    return any(k in val_str.lower() for k in _HEADER_KEYWORDS)

def _digit_ratio(val_str):
    alnum_chars = [c for c in val_str if c.isalnum()]
    if not alnum_chars:
        return 0.0
    digits = sum(1 for c in alnum_chars if c.isdigit())
    return digits / len(alnum_chars)

def is_id_like(val):
    """
    員編特徵：不含中文字元、字元中數字佔比達一半以上、且長度至少3（排除單一位數的當日工時等雜訊）。
    """
    if val is None:
        return False
    val_str = str(val).strip()
    if not val_str:
        return False
    if re.search(r"[一-龥]", val_str):
        return False
    alnum_len = sum(1 for c in val_str if c.isalnum())
    if alnum_len < 3:
        return False
    return _digit_ratio(val_str) >= 0.5

def find_name_rows(rows, name_col, start_idx, max_len=20):
    """
    掃描 name_col 欄位找出真正的姓名列：先用 is_name_like 做結構性篩選，
    再用「同一欄位中該字串出現次數」排除重複出現的類別標籤（如職稱），
    只有恰好出現一次的值才視為姓名，藉此取代寫死的職稱關鍵字清單。
    """
    candidates = []
    for r_idx in range(start_idx, len(rows)):
        row = rows[r_idx]
        is_empty = not any(v is not None and str(v).strip() for v in row.values())
        if is_empty:
            continue
        val = row.get(name_col)
        if val is None:
            continue
        val_str = str(val).strip()
        if not val_str or len(val_str) > max_len:
            continue
        if is_name_like(val):
            candidates.append((r_idx, val_str))

    freq = Counter(v for _, v in candidates)
    return [r_idx for r_idx, v in candidates if freq[v] == 1]

def find_block_anchors(rows, name_col, expected_rows, start_idx, name_row_offset=None):
    """
    Groups rows into "same position within the block grid" buckets, using the
    first genuine name-like cell in name_col (from start_idx onward) as a
    phase anchor -- rather than tiling rigidly from start_idx. A one-off
    leading row that isn't part of the repeating block grid (e.g. a weekday
    sub-header row, or a one-time summary row) is blank/non-name-shaped in
    name_col, so it's never picked up as a candidate and can't misalign every
    subsequent block the way rigid tiling from a fixed start point would.

    If name_row_offset is given, returns just that bucket's row indices
    (the final list of employee block anchors for a known template). If
    omitted, returns every bucket as {offset: [row_idx, ...]} so the caller
    can score which offset is the real name row vs. a category label like a
    job title (see guess_template).
    """
    raw_candidates = [r for r in range(start_idx, len(rows)) if is_name_like(rows[r].get(name_col))]
    if not raw_candidates:
        return [] if name_row_offset is not None else {}

    anchor = raw_candidates[0]
    buckets = {}
    for r in raw_candidates:
        offset = (r - anchor) % expected_rows
        buckets.setdefault(offset, []).append(r)

    if name_row_offset is not None:
        return buckets.get(name_row_offset % expected_rows, [])
    return buckets

def is_name_like(val):
    if val is None:
        return False
    val_str = str(val).strip()
    if not val_str:
        return False

    # Avoid numbers, dates, titles
    if val_str.replace(".", "", 1).isdigit():
        return False
    if is_date_or_day_num(val):
        return False

    if is_header_keyword(val_str):
        return False

    # Chinese name (2-4 characters) or alphanumeric name (digit ratio < 50%, so
    # "TEST1"/"James2" style names pass while "22040023"/"2306001" style IDs don't)
    if re.match(r"^[\u4e00-\u9fa5]{2,4}$", val_str):
        return True
    if re.match(r"^[a-zA-Z0-9\s]{2,15}$", val_str) and _digit_ratio(val_str) < 0.5:
        return True

    return False

def guess_template(file_path):
    """
    Scans the file and guesses the layout type and template structure.
    Returns a suggested template dictionary.
    """
    # 1. Identify sheet candidates and select one
    # ODS or XLSX
    if file_path.endswith(".ods"):
        from backend.parser import read_ods_rows as reader
        import zipfile
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(file_path) as z:
            root = ET.fromstring(z.read("content.xml"))
        sheet_names = []
        for t in root.iter("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table"):
            sheet_names.append(t.get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name"))
    else:
        from backend.parser import read_xlsx_rows as reader
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True)
        try:
            sheet_names = wb.sheetnames
        finally:
            wb.close()

    # Choose sheet: prefer names containing "整月", "外場", "內場", "班表"
    sheet_candidates = [n for n in sheet_names if any(k in n for k in ["整月", "外場", "內場", "班表"])]
    selected_sheet = sheet_candidates[0] if sheet_candidates else sheet_names[0]
    
    rows = reader(file_path, selected_sheet)
    if not rows:
        raise ValueError(f"工作表 {selected_sheet} 為空。")
        
    # Limit row scans
    rows = rows[:120]
    
    # 2. Find Date Row Index
    date_row_idx = None
    max_date_matches = 0
    
    for r_idx, row in enumerate(rows):
        # We need a valid date row to contain multiple unique date/number values
        # to filter out rows containing the same month number repeated (e.g. Row 1 containing all 2s)
        date_vals = [c_val for c_val in row.values() if is_date_or_day_num(c_val)]
        unique_dates = len(set(str(v).strip() for v in date_vals))
        
        # A valid date row must have at least 15 matches and at least 10 unique values
        if len(date_vals) > max_date_matches and len(date_vals) >= 15 and unique_dates >= 10:
            max_date_matches = len(date_vals)
            date_row_idx = r_idx
            
    if date_row_idx is None:
        # Fallback to row 4
        date_row_idx = 3
        
    date_row = rows[date_row_idx]
    
    # Find columns containing dates
    date_cols = sorted([col for col, val in date_row.items() if is_date_or_day_num(val)])
    if not date_cols:
        date_cols = list(range(4, 35)) # fallback
        
    first_day_col = date_cols[0]
    
    # Determine cols_per_day
    # Check spacing of first few date columns
    if len(date_cols) >= 3:
        diffs = [date_cols[i] - date_cols[i-1] for i in range(1, min(5, len(date_cols)))]
        cols_per_day = Counter(diffs).most_common(1)[0][0]
    else:
        cols_per_day = 2
        
    # 3. Find Name Column
    # Check columns 1 to 5 for rows below date_row_idx
    name_col = 2 # default Column B
    max_name_matches = 0
    for col in range(1, 6):
        matches = len(find_name_rows(rows, col, date_row_idx + 1))
        if matches > max_name_matches:
            max_name_matches = matches
            name_col = col

    # 4. Determine block size (expected_rows)
    # Collect rows containing names in name_col
    name_rows = find_name_rows(rows, name_col, date_row_idx + 1)
    
    if len(name_rows) >= 3:
        diffs = [name_rows[i] - name_rows[i-1] for i in range(1, len(name_rows))]
        # The spacing represents the block size
        expected_rows = Counter(diffs).most_common(1)[0][0]
    else:
        expected_rows = 3 # default
        
    # If the spacing is very large or weird, fall back to 1
    if expected_rows > 5:
        expected_rows = 1
        
    row_meanings = []

    # Determine which row_offset within a block actually holds the employee's name
    # (vs. a category label like job title). Job titles are a controlled vocabulary
    # shared across employees, but a title held by only one employee (e.g. only one
    # 副店長) never repeats and so can't be excluded by frequency alone. Instead, tile
    # the sheet at each candidate offset and count how many DISTINCT name-like values
    # occur there: the true name offset has one distinct value per employee (maximal),
    # while a title offset -- even with some unique titles mixed in -- has fewer
    # distinct values than employees because most titles repeat. Run uniformly for any
    # block size (including expected_rows==1, where there's only one possible offset: 0).
    offset_buckets = find_block_anchors(rows, name_col, expected_rows, date_row_idx + 1)
    name_block_offset = 0
    best_distinct = -1
    for offset, anchor_rows in offset_buckets.items():
        distinct = len(set(str(rows[r].get(name_col)).strip() for r in anchor_rows))
        if distinct > best_distinct:
            best_distinct = distinct
            name_block_offset = offset

    block_name_rows = offset_buckets.get(name_block_offset, [])

    if expected_rows == 1:
        # Check if cells in date columns contain text format shifts
        sample_val = ""
        for col in date_cols:
            val = rows[name_rows[0]].get(col) if name_rows else None
            if val and "-" in str(val):
                sample_val = str(val)
                break

        if "-" in sample_val:
            # Layout Type B
            row_meanings = [
                {"type": "shift_string", "index": 0}
            ]
        else:
            # Simple 1 shift grid
            row_meanings = [
                {"type": "shift", "index": 0, "name": "班別"}
            ]
    elif expected_rows == 2:
        # Layout Type C (start and end on different rows)
        row_meanings = [
            {"type": "shift_start", "index": 0},
            {"type": "shift_end", "index": 0}
        ]
    else:
        # Layout Type A (3-row or multi-row grid layout)
        # The block's row count (expected_rows) is not assumed to map to a fixed
        # "2 shifts + 1 metadata row" structure (that would hardcode the shift count,
        # conflicting with the requirement to support N shifts per block, e.g. 3-段班).
        # Observed real samples consistently keep the employee_id/daily-total row as the
        # LAST row of the block, so that convention is used as the split point; every
        # other row in the block becomes a sequential shift row in order.
        metadata_row_offset = expected_rows - 1

        # Vote across all blocks for which column (in the label area before first_day_col,
        # including name_col itself) holds an ID-like value in the metadata row. If no
        # column clears the confidence bar, no employee_id is guessed (left unset) rather
        # than defaulting to a possibly-wrong column.
        id_votes = {}
        block_count = 0
        for name_row_idx in block_name_rows:
            block_start = name_row_idx - name_block_offset
            meta_row_idx = block_start + metadata_row_offset
            if meta_row_idx >= len(rows):
                continue
            block_count += 1
            meta_row = rows[meta_row_idx]
            for col in range(1, first_day_col):
                if meta_row_idx == name_row_idx and col == name_col:
                    continue  # skip the exact cell the name itself was found in
                if is_id_like(meta_row.get(col)):
                    id_votes[col] = id_votes.get(col, 0) + 1

        employee_id_col = None
        if id_votes and block_count:
            best_col, best_count = max(id_votes.items(), key=lambda kv: kv[1])
            if best_count / block_count >= 0.5:
                employee_id_col = best_col

        row_meanings = []
        for row_offset in range(expected_rows):
            if row_offset == metadata_row_offset:
                meaning = {"type": "metadata", "name": "employee_id"}
                if employee_id_col is not None:
                    meaning["col"] = employee_id_col
                row_meanings.append(meaning)
            else:
                # metadata_row_offset is always the last row, so every other row_offset
                # (0..expected_rows-2) maps directly to a sequential shift index.
                row_meanings.append({"type": "shift", "index": row_offset, "name": f"班別{row_offset + 1}"})

    # Best-effort guess of which physical row holds the FIRST real employee's first
    # shift cell, for display purposes only (e.g. front-end shows "C5" instead of just
    # "C"). Anchored on the earliest confirmed name row rather than on first_day_col's
    # own content, because a genuine first employee can coincidentally have an atypical
    # first-day value (on leave, blank, ad-hoc note) that would fool a "looks like a
    # valid shift" scan into skipping past the real first block onto a later one.
    first_data_row = None
    if block_name_rows:
        first_shift_offset = 0
        for idx, meaning in enumerate(row_meanings):
            if meaning["type"] in ("shift", "shift_start", "shift_string"):
                first_shift_offset = idx
                break
        earliest_name_row = min(block_name_rows)
        target_row = earliest_name_row - name_block_offset + first_shift_offset
        if 0 <= target_row < len(rows):
            first_data_row = target_row + 1  # 1-based

    suggested_template = {
        "template_name": f"{selected_sheet} 啟發式預測範本",
        "sheet_name": selected_sheet,
        "header_row_index": date_row_idx + 1, # 1-based index
        "first_data_row": first_data_row, # 1-based；純顯示用途，不影響解析
        "mapping": {
            "name_col": name_col,
            "first_day_col": first_day_col,
            "cols_per_day": cols_per_day
        },
        "block": {
            "expected_rows": expected_rows,
            "name_row_offset": name_block_offset,
            "row_meanings": row_meanings
        }
    }

    return suggested_template

def guess_axis_range(employees, n_days=7, default_start=8, default_end=24):
    """
    Best-effort guess of the Gantt chart's time axis range, from actually parsed
    shift data (not structural guessing). Scans the first n_days across all
    employees, taking the earliest shift start (rounded down) and latest shift
    end (rounded up), for display purposes only -- caller should fall back to
    (default_start, default_end) if this raises or the data has no valid shifts.
    """
    min_start = None
    max_end = None
    for emp in employees:
        for d in range(1, n_days + 1):
            day_info = emp.get("days", {}).get(d)
            if not day_info:
                continue
            for shift in day_info.get("shifts", []):
                s, e = shift.get("start"), shift.get("end")
                # 排除 0 長度的佔位班別（例如某些範本用 0:00-0:00 代表「無排班」而非
                # 留空），否則會把座標軸誤拉到 0 點。
                if s is None or e is None or e <= s:
                    continue
                if min_start is None or s < min_start:
                    min_start = s
                if max_end is None or e > max_end:
                    max_end = e

    if min_start is None or max_end is None:
        return {"axis_start": default_start, "axis_end": default_end}

    axis_start = max(0, min(24, math.floor(min_start)))
    axis_end = max(0, min(24, math.ceil(max_end)))
    if axis_end <= axis_start:
        return {"axis_start": default_start, "axis_end": default_end}

    return {"axis_start": axis_start, "axis_end": axis_end}

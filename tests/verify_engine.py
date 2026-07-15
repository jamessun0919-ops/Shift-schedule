import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from backend.parser import parse_schedule
from backend.generator import generate_xlsx, generate_ods
from backend.heuristics import guess_template

white_mist_template = {
    "template_name": "白霧預設範本",
    "sheet_name": "1",
    "header_row_index": 5,
    "mapping": {
        "name_col": 4,
        "title_col": 5,
        "employee_id_col": None,
        "first_day_col": 6,
        "cols_per_day": 2
    },
    "block": {
        "expected_rows": 3,
        "row_meanings": [
            {"type": "shift", "index": 0, "name": "早班"},
            {"type": "shift", "index": 1, "name": "晚班"},
            {"type": "metadata", "name": "employee_id", "col": 4}
        ]
    }
}

def test_original_parsing():
    print("=== Testing Original Schedule Parsing (白霧.ods) ===")
    ods_file = "白霧.ods"
    if not os.path.exists(ods_file):
        print(f"Error: {ods_file} not found.")
        return
        
    result = parse_schedule(ods_file, white_mist_template)
    print(f"Parsing Success: {result['is_healthy']}")
    print(f"Parsed Employee Count: {len(result['employees'])}")
    print(f"Anomalies Found: {len(result['anomalies'])}")
    for a in result['anomalies']:
        print(f"  Anomaly: {a}")
        
    if result['employees']:
        emp0 = result['employees'][0]
        print(f"TEST1 Day 1 Shifts: {emp0['days'][1]['shifts']}")
        
    # Generate XLSX and ODS tests
    out_xlsx = os.path.join("tests", "output_test.xlsx")
    generate_xlsx(result['employees'], out_xlsx, white_mist_template)
    
    out_ods = os.path.join("tests", "output_test.ods")
    generate_ods(result['employees'], out_ods, white_mist_template, template_ods_path="白霧.ods")

def test_heuristics():
    print("\n=== Testing Heuristic Layout Detection ===")
    files = [
        "範例班表1.ods",
        "範例班表2.ods",
        "範例班表3.ods",
        "範例班表4.xlsx",
        "範例班表5.ods"
    ]
    
    for f in files:
        if not os.path.exists(f):
            print(f"File not found: {f}")
            continue
            
        print(f"\nAnalyzing file layout for: {f}")
        try:
            template = guess_template(f)
            print(f"Guessed Template Name: {template['template_name']}")
            print(f"  Header Row: {template['header_row_index']}")
            print(f"  Mapping: {template['mapping']}")
            print(f"  Block expected_rows: {template['block']['expected_rows']}")
            print(f"  Block row meanings: {template['block']['row_meanings']}")
            
            # Now run parsing with this guessed template
            parse_result = parse_schedule(f, template)
            print(f"  Parse Result: healthy={parse_result['is_healthy']}, employees={len(parse_result['employees'])}")
            if parse_result['employees']:
                sample_emp = parse_result['employees'][0]
                print(f"  Sample Employee: '{sample_emp['name']}'")
                print(f"  Day 1 Shifts: {sample_emp['days'][1]['shifts']}")
                
                # Write a converted test sheet to verify generator
                out_path = os.path.join("tests", f"converted_{f.split('.')[0]}.xlsx")
                generate_xlsx(parse_result['employees'], out_path, template)
                print(f"  Saved converted XLSX to: {out_path}")
        except Exception as e:
            print(f"  Failed heuristic test on {f}: {e}")

if __name__ == "__main__":
    test_original_parsing()
    test_heuristics()

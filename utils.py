from config import logger

def chinese_to_int(chinese_num):
    """Convert Chinese numerals to Arabic integers."""
    numeral_map = {
        '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
    }
    return numeral_map.get(chinese_num, 0)

def parse_district(district_name):
    """
    Parse district name into main and sub parts.
    Returns tuple of (main_district, sub_district_num) where sub_district_num is an integer.
    """
    if not district_name:  # Handle empty string case
        return ("未知區", 0)
    
    parts = district_name.split("區")
    if len(parts) < 2:  # If no "區" or incomplete format
        return (district_name, 0)
    
    main_part = parts[0] + "區"  # e.g., "二大區"
    sub_part = parts[1]          # e.g., "三"
    
    # 將子區的中文數字轉換為整數
    sub_district_num = chinese_to_int(sub_part) if sub_part in ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'] else 0
    return (main_part, sub_district_num)

# 原有的 render_attendance_table 函數已移除，因為它現在位於 render_table.py 中

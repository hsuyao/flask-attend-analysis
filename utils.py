
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
    Returns tuple of (main_district, sub_district)
    """
    if not district_name:  # Handle empty string case
        return ("未知區", "未知小區")
    
    parts = district_name.split("區")
    if len(parts) < 2:  # If no "區" or incomplete format
        return (district_name, "未知小區")
    
    main_part = parts[0] + "區"  # e.g., "一大區"
    sub_part = parts[1]          # e.g., "一"
    
    return (main_part, sub_part)

# In render_table.py
def render_attendance_table(attendance_data):
    if not attendance_data or 'attended' not in attendance_data or 'not_attended' not in attendance_data:
        return "<p>No attendance data available</p>"
    
    # Get all unique districts, handling potential empty strings
    districts = sorted(set(attendance_data['attended'].keys()).union(attendance_data['not_attended'].keys()))
    
    # Build HTML table
    html = '<table class="attendance-table">\n'
    html += '<tr><th>區</th><th>小區</th><th>參加</th><th>未參加</th></tr>\n'
    
    for district in districts:
        main_district, sub_district = parse_district(district)
        attended = ", ".join(attendance_data['attended'].get(district, []))
        not_attended = ", ".join(attendance_data['not_attended'].get(district, []))
        
        html += f'<tr><td>{main_district}</td><td>{sub_district}</td><td>{attended}</td><td>{not_attended}</td></tr>\n'
    
    html += '</table>'
    return html

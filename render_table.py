# render_table.py
from config import logger
from utils import chinese_to_int

def render_attendance_table(week_display, latest_attendance_data, all_attendance_data):
    if not latest_attendance_data:
        return """
        <div class="table-wrapper">
            <table class="excel-table">
                <tr class="title-row"><th>無出席資料</th></tr>
            </table>
        </div>
        """

    districts = sorted(set(latest_attendance_data['attended'].keys()).union(latest_attendance_data['not_attended'].keys()), 
                      key=lambda x: chinese_to_int(x[3:4]))
    max_len = max(max(len(latest_attendance_data['attended'].get(d, [])), len(latest_attendance_data['not_attended'].get(d, []))) for d in districts)
    
    # Find the previous week's data by comparing dates
    previous_week_data = None
    all_attendance_data.sort(key=lambda x: x[0])  # Sort by date
    current_week_idx = None
    for idx, (date, data, week_name) in enumerate(all_attendance_data):
        if week_name == week_display:
            current_week_idx = idx
            break
    
    if current_week_idx is not None and current_week_idx > 0:
        previous_week_data = all_attendance_data[current_week_idx - 1][1]  # Previous week's data

    sorted_attended = {}
    sorted_not_attended = {}
    for district in districts:
        attended_list = latest_attendance_data['attended'].get(district, [])
        not_attended_list = latest_attendance_data['not_attended'].get(district, [])
        
        attended_with_highlights = []
        not_attended_with_highlights = []
        if previous_week_data:
            prev_attended = previous_week_data['attended'].get(district, [])
            prev_not_attended = previous_week_data['not_attended'].get(district, [])
            
            for name in attended_list:
                display_name = name[:4] if len(name) > 4 else name
                highlight = 'highlight-green' if name in prev_not_attended else ''
                attended_with_highlights.append((name, display_name, highlight))
            
            for name in not_attended_list:
                display_name = name[:4] if len(name) > 4 else name
                highlight = 'highlight-red' if name in prev_attended else ''
                not_attended_with_highlights.append((name, display_name, highlight))
        
        else:
            attended_with_highlights = [(name, name[:4] if len(name) > 4 else name, '') for name in attended_list]
            not_attended_with_highlights = [(name, name[:4] if len(name) > 4 else name, '') for name in not_attended_list]
        
        attended_with_highlights.sort(key=lambda x: (x[2] == '', x[0]))
        not_attended_with_highlights.sort(key=lambda x: (x[2] == '', x[0]))
        
        sorted_attended[district] = attended_with_highlights
        sorted_not_attended[district] = not_attended_with_highlights
        max_len = max(max_len, max(len(attended_with_highlights), len(not_attended_with_highlights)))

    attendance_table_html = """
    <div class="table-wrapper">
        <table class="excel-table">
    """

    total_attendance_cols = len(districts) * 2

    attendance_table_html += f'<tr class="title-row">'
    attendance_table_html += f'<th colspan="{total_attendance_cols}">{week_display}</th>'
    attendance_table_html += '</tr>'

    attendance_table_html += '<tr class="header">'
    for district in districts:
        attendance_table_html += f'<th colspan="2">{district}</th>'
    attendance_table_html += '</tr>'

    attendance_table_html += '<tr class="subheader">'
    for _ in districts:
        attendance_table_html += '<th>本週到會</th><th>未到會</th>'
    attendance_table_html += '</tr>'

    for r in range(max_len):
        row_class = "even" if r % 2 == 0 else "odd"
        attendance_table_html += f'<tr class="{row_class}">'
        for district in districts:
            attended_with_highlights = sorted_attended.get(district, [])
            not_attended_with_highlights = sorted_not_attended.get(district, [])
            attended_info = attended_with_highlights[r] if r < len(attended_with_highlights) else ('', '', '')
            not_attended_info = not_attended_with_highlights[r] if r < len(not_attended_with_highlights) else ('', '', '')
            attended_display = attended_info[1]
            not_attended_display = not_attended_info[1]
            attended_class = attended_info[2]
            not_attended_class = not_attended_info[2]

            attendance_table_html += f'<td class="{attended_class}">{attended_display}</td><td class="{not_attended_class}">{not_attended_display}</td>'
        attendance_table_html += '</tr>'

    attendance_table_html += "</table></div>"

    return attendance_table_html

def render_stats_table(latest_district_counts, latest_main_district):
    if not latest_district_counts:
        return """
        <div class="table-wrapper">
            <table class="excel-table">
                <tr class="title-row"><th>無統計資料</th></tr>
            </table>
        </div>
        """

    stats_districts = sorted([d for d in latest_district_counts.keys() if d != '總計'], key=lambda x: chinese_to_int(x[3:4]))
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']

    stats_table_html = """
    <div class="table-wrapper">
        <table class="excel-table">
    """

    stats_rows = []
    row_index = 0
    for district in stats_districts:
        row_class = "even" if row_index % 2 == 0 else "odd"
        stats_rows.append((row_class, f'<td colspan="2" class="district-header">{district}</td>'))
        row_index += 1
        for age in age_categories:
            count = latest_district_counts[district]['ages'][age]
            row_class = "even" if row_index % 2 == 0 else "odd"
            stats_rows.append((row_class, f'<td class="sub-row" style="padding-left: 15px;">{age}</td><td class="sub-row">{count}</td>'))
            row_index += 1
        total = latest_district_counts[district]['total']
        row_class = "even" if row_index % 2 == 0 else "odd"
        stats_rows.append((row_class, f'<td class="sub-row" style="padding-left: 15px;">總計</td><td class="sub-row">{total}</td>'))
        row_index += 1

    main_district = latest_main_district if latest_main_district else "未知大區"
    overall_ages = {age: sum(latest_district_counts[d]['ages'][age] for d in stats_districts) for age in age_categories}
    total_attendance = latest_district_counts['總計']
    row_class = "even" if row_index % 2 == 0 else "odd"
    stats_rows.append((row_class, f'<td colspan="2" class="district-header">{main_district}</td>'))
    row_index += 1
    for age in age_categories:
        count = overall_ages[age]
        row_class = "even" if row_index % 2 == 0 else "odd"
        stats_rows.append((row_class, f'<td class="sub-row" style="padding-left: 15px;">{age}</td><td class="sub-row">{count}</td>'))
        row_index += 1
    row_class = "even" if row_index % 2 == 0 else "odd"
    stats_rows.append((row_class, f'<td class="sub-row" style="padding-left: 15px;">總計</td><td class="sub-row">{total_attendance}</td>'))

    for row_class, stats_cells in stats_rows:
        stats_table_html += f'<tr class="{row_class}">{stats_cells}</tr>'

    stats_table_html += "</table></div>"

    return stats_table_html

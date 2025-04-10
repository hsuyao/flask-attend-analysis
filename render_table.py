# render_table.py
from config import logger, state
from utils import chinese_to_int

def render_combined_table(week_display):
    combined_table_html = ""
    if not state.latest_attendance_data or not state.latest_district_counts:
        # Display a default table or message when no data is available
        combined_table_html = """
        <div class="table-wrapper">
            <table class="excel-table">
                <tr class="title-row"><th>No Data Available</th></tr>
            </table>
        </div>
        """
        return combined_table_html

    districts = sorted(set(state.latest_attendance_data['attended'].keys()).union(state.latest_attendance_data['not_attended'].keys()), 
                      key=lambda x: chinese_to_int(x[3:4]))
    max_len = max(max(len(state.latest_attendance_data['attended'].get(d, [])), len(state.latest_attendance_data['not_attended'].get(d, []))) for d in districts)
    stats_districts = sorted([d for d in state.latest_district_counts.keys() if d != '總計'], key=lambda x: chinese_to_int(x[3:4]))
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']
    
    previous_week_data = None
    if len(state.all_attendance_data) > 1:
        state.all_attendance_data.sort(key=lambda x: x[0])
        latest_date = state.all_attendance_data[-1][0]
        for date, data, week_name in reversed(state.all_attendance_data[:-1]):
            if date < latest_date:
                previous_week_data = data
                break

    sorted_attended = {}
    sorted_not_attended = {}
    for district in districts:
        attended_list = state.latest_attendance_data['attended'].get(district, [])
        not_attended_list = state.latest_attendance_data['not_attended'].get(district, [])
        
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

    combined_table_html = """
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
            count = state.latest_district_counts[district]['ages'][age]
            row_class = "even" if row_index % 2 == 0 else "odd"
            stats_rows.append((row_class, f'<td class="sub-row" style="padding-left: 15px;">{age}</td><td class="sub-row">{count}</td>'))
            row_index += 1
        total = state.latest_district_counts[district]['total']
        row_class = "even" if row_index % 2 == 0 else "odd"
        stats_rows.append((row_class, f'<td class="sub-row" style="padding-left: 15px;">總計</td><td class="sub-row">{total}</td>'))
        row_index += 1

    main_district = state.latest_main_district if state.latest_main_district else "未知大區"
    overall_ages = {age: sum(state.latest_district_counts[d]['ages'][age] for d in stats_districts) for age in age_categories}
    total_attendance = state.latest_district_counts['總計']
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

    stats_index = 0
    total_attendance_cols = len(districts) * 2

    combined_table_html += f'<tr class="title-row">'
    combined_table_html += f'<th colspan="{total_attendance_cols + 1}">{week_display}</th>'
    if stats_index < len(stats_rows):
        row_class, stats_cells = stats_rows[stats_index]
        combined_table_html += stats_cells
        stats_index += 1
    else:
        combined_table_html += '<td></td><td></td>'
    combined_table_html += '</tr>'

    combined_table_html += '<tr class="header">'
    for district in districts:
        combined_table_html += f'<th colspan="2">{district}</th>'
    combined_table_html += '<th class="separator"></th>'
    if stats_index < len(stats_rows):
        row_class, stats_cells = stats_rows[stats_index]
        combined_table_html += stats_cells
        stats_index += 1
    else:
        combined_table_html += '<td></td><td></td>'
    combined_table_html += '</tr>'

    combined_table_html += '<tr class="subheader">'
    for _ in districts:
        combined_table_html += '<th>本週到會</th><th>未到會</th>'
    combined_table_html += '<th class="separator"></th>'
    if stats_index < len(stats_rows):
        row_class, stats_cells = stats_rows[stats_index]
        combined_table_html += stats_cells
        stats_index += 1
    else:
        combined_table_html += '<td></td><td></td>'
    combined_table_html += '</tr>'

    for r in range(max_len):
        row_class = "even" if r % 2 == 0 else "odd"
        combined_table_html += f'<tr class="{row_class}">'
        for district in districts:
            attended_with_highlights = sorted_attended.get(district, [])
            not_attended_with_highlights = sorted_not_attended.get(district, [])
            attended_info = attended_with_highlights[r] if r < len(attended_with_highlights) else ('', '', '')
            not_attended_info = not_attended_with_highlights[r] if r < len(not_attended_with_highlights) else ('', '', '')
            attended_display = attended_info[1]
            not_attended_display = not_attended_info[1]
            attended_class = attended_info[2]
            not_attended_class = not_attended_info[2]

            combined_table_html += f'<td class="{attended_class}">{attended_display}</td><td class="{not_attended_class}">{not_attended_display}</td>'
        combined_table_html += '<td class="separator"></td>'
        if stats_index < len(stats_rows):
            row_class, stats_cells = stats_rows[stats_index]
            combined_table_html += stats_cells
            stats_index += 1
        else:
            combined_table_html += '<td></td><td></td>'
        combined_table_html += '</tr>'

    while stats_index < len(stats_rows):
        row_class, stats_cells = stats_rows[stats_index]
        combined_table_html += f'<tr class="{row_class}">'
        for _ in districts:
            combined_table_html += '<td></td><td></td>'
        combined_table_html += '<td class="separator"></td>'
        combined_table_html += stats_cells
        stats_index += 1
        combined_table_html += '</tr>'

    combined_table_html += "</table></div>"

    return combined_table_html

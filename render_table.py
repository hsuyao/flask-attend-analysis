from config import logger
from utils import chinese_to_int, parse_district

def render_attendance_table(week_display, latest_attendance_data, all_attendance_data, latest_district_counts, latest_main_district_counts):
    districts = sorted(set(latest_attendance_data['attended'].keys()).union(latest_attendance_data['not_attended'].keys()), 
                      key=parse_district)
    
    if not districts:
        return """
        <div class="district-section">
            <table class="excel-table">
                <tr class="title-row"><th>該週無有效數據</th></tr>
            </table>
        </div>
        """

    previous_week_data = None
    all_attendance_data.sort(key=lambda x: x[0])
    current_week_idx = next((idx for idx, (date, data, week_name) in enumerate(all_attendance_data) if week_name == week_display), None)
    
    if current_week_idx is not None and current_week_idx > 0:
        previous_week_data = all_attendance_data[current_week_idx - 1][1]

    main_districts = sorted(set(parse_district(d)[0] for d in districts), key=lambda x: chinese_to_int(x[0]))
    district_groups = {md: [d for d in districts if d.startswith(md)] for md in main_districts}

    html = ""
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']

    for main_district in main_districts:
        sub_districts = district_groups[main_district]
        if not sub_districts:
            continue

        max_len = max(max(len(latest_attendance_data['attended'].get(d, [])), len(latest_attendance_data['not_attended'].get(d, []))) for d in sub_districts)

        sorted_attended = {}
        sorted_not_attended = {}
        for district in sub_districts:
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

        # 開始大區區塊
        html += f'<div class="district-section">\n'
        html += f'<h2>{main_district} - {week_display}</h2>\n'
        html += '<div class="district-container">\n'

        # 出勤名單表
        html += '<div class="table-wrapper attendance-wrapper">\n<table class="excel-table">\n'
        total_cols = len(sub_districts) * 2
        html += f'<tr class="header"><th colspan="{total_cols}">{main_district}</th></tr>\n'
        html += '<tr class="district-row">\n'
        for district in sub_districts:
            html += f'<th colspan="2">{district}</th>'
        html += '</tr>\n'
        html += '<tr class="subheader">\n'
        for _ in sub_districts:
            html += '<th>本週到會</th><th>未到會</th>'
        html += '</tr>\n'

        for r in range(max_len):
            row_class = "even" if r % 2 == 0 else "odd"
            html += f'<tr class="{row_class}">\n'
            for district in sub_districts:
                attended_with_highlights = sorted_attended.get(district, [])
                not_attended_with_highlights = sorted_not_attended.get(district, [])
                attended_info = attended_with_highlights[r] if r < len(attended_with_highlights) else ('', '', '')
                not_attended_info = not_attended_with_highlights[r] if r < len(not_attended_with_highlights) else ('', '', '')
                attended_display = attended_info[1]
                not_attended_display = not_attended_info[1]
                attended_class = attended_info[2]
                not_attended_class = not_attended_info[2]
                html += f'<td class="{attended_class}">{attended_display}</td><td class="{not_attended_class}">{not_attended_display}</td>'
            html += '</tr>\n'
        html += '</table>\n</div>\n'

        # 統計表（總計移至標題行並染色）
        stats_districts = sorted([d for d in latest_district_counts.keys() if d != '總計'], key=parse_district)
        sub_districts_stats = [d for d in stats_districts if d.startswith(main_district)]
        if sub_districts_stats:
            html += '<div class="table-wrapper stats-wrapper">\n<table class="excel-table">\n'
            html += f'<tr class="header"><th colspan="2">{main_district} 統計</th></tr>\n'
            row_index = 0
            
            # 子區統計
            for district in sub_districts_stats:
                total = latest_district_counts[district]['total']
                html += f'<tr class="total-row"><td style="padding-left: 15px;">{district}</td><td>{total}</td></tr>\n'
                row_index += 1
                for age in age_categories:
                    count = latest_district_counts[district]['ages'][age]
                    row_class = "even" if row_index % 2 == 0 else "odd"
                    html += f'<tr class="{row_class}"><td style="padding-left: 30px;">{age}</td><td>{count}</td></tr>\n'
                    row_index += 1
            
            # 主區統計
            total = latest_main_district_counts[main_district]['total']
            html += f'<tr class="total-row"><td style="padding-left: 15px;">{main_district}</td><td>{total}</td></tr>\n'
            row_index += 1
            for age in age_categories:
                count = latest_main_district_counts[main_district]['ages'][age]
                row_class = "even" if row_index % 2 == 0 else "odd"
                html += f'<tr class="{row_class}"><td style="padding-left: 30px;">{age}</td><td>{count}</td></tr>\n'
                row_index += 1

            html += '</table>\n</div>\n'

        html += '</div>\n</div>\n'

    if not html:
        html = """
        <div class="district-section">
            <table class="excel-table">
                <tr class="title-row"><th>該週無有效數據</th></tr>
            </table>
        </div>
        """

    return html

from config import logger
from utils import chinese_to_int, parse_district
from database import get_six_month_averages
from excel_handler import get_all_latest_attendance_dates
from datetime import datetime

def render_stats_table(main_district, district_counts, main_district_counts):
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']
    stats_districts = sorted(
        [d for d in district_counts.keys() if d != '總計'],
        key=parse_district
    )
    sub_districts = [d for d in stats_districts if d.startswith(main_district)]

    if not sub_districts and main_district not in main_district_counts:
        return '<div class="table-wrapper stats-wrapper flex-item"><p>無統計資料</p></div>'

    html = '<div class="table-wrapper stats-wrapper flex-item">\n<table class="excel-table">\n'
    html += f'<tr class="header"><th></th>'
    districts = [main_district] + sub_districts
    for district in districts:
        html += f'<th>{district}</th>'
    html += '</tr>\n'

    for i, age in enumerate(age_categories):
        html += f'<tr class="{"even" if i % 2 == 0 else "odd"}"><td>{age}</td>'
        for district in districts:
            count = main_district_counts.get(district, {'ages': {age: 0}}).get('ages', {age: 0}).get(age, 0) if district == main_district else \
                    district_counts.get(district, {'ages': {age: 0}}).get('ages', {age: 0}).get(age, 0)
            html += f'<td>{count}</td>'
        html += '</tr>\n'

    html += '<tr class="total-row"><td>總計</td>'
    for district in districts:
        total = main_district_counts.get(district, {'total': 0}).get('total', 0) if district == main_district else \
                district_counts.get(district, {'total': 0}).get('total', 0)
        html += f'<td>{total}</td>'
    html += '</tr>\n</table>\n</div>\n'
    return html

def render_attendance_table(week_display, latest_attendance_data, all_attendance_data, district_counts, main_district_counts, avg_attendance_rates=None):
    if avg_attendance_rates is None:
        avg_attendance_rates = {}

    all_districts = set(latest_attendance_data['attended'].keys()).union(latest_attendance_data['not_attended'].keys())
    districts = sorted(
        [d for d in all_districts if len(parse_district(d)) > 1 and parse_district(d)[1]],
        key=parse_district
    )
    if not districts:
        return '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>無資料</th></tr></table></div>'

    previous_week_data = None
    all_attendance_data.sort(key=lambda x: x[2])  # Sort by week_display
    current_week_idx = next((idx for idx, (_, _, week_name) in enumerate(all_attendance_data) if week_name == week_display), None)

    if current_week_idx is not None and current_week_idx > 0:
        previous_week_data = all_attendance_data[current_week_idx - 1][1]

    # Use placeholder date for latest_dates; actual date not critical
    placeholder_date = datetime.now()
    main_districts = sorted(set(parse_district(d)[0] for d in districts), key=lambda x: chinese_to_int(x[0]))
    district_groups = {md: [d for d in districts if d.startswith(md)] for md in main_districts}

    html = ""
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
            combined_list = []

            if previous_week_data:
                prev_attended = previous_week_data['attended'].get(district, [])
                prev_not_attended = previous_week_data['not_attended'].get(district, [])
                combined_list.extend((name, True, name in prev_not_attended) for name in attended_list)
                combined_list.extend((name, False, name in prev_attended) for name in not_attended_list)
            else:
                combined_list.extend((name, True, False) for name in attended_list)
                combined_list.extend((name, False, False) for name in not_attended_list)

            names = [name for name, _, _ in combined_list]
            latest_dates = get_all_latest_attendance_dates(names, placeholder_date)
            logger.debug(f"District {district}: Attendance rates {[(name, avg_attendance_rates.get(name, 0.0)) for name in names]}")
            logger.debug(f"District {district}: Highlights {[(name, has_highlight) for name, _, has_highlight in combined_list]}")
            # Sort by highlight status (True first), then attendance rate (descending), then latest date
            combined_list.sort(
                key=lambda x: (
                    -int(x[2]),                           # Primary: highlight (True first)
                    -avg_attendance_rates.get(x[0], 0.0),  # Secondary: descending attendance rate
                    latest_dates.get(x[0], datetime(1970, 1, 1))  # Tertiary: latest date
                )
            )

            attended_with_highlights = []
            not_attended_with_highlights = []
            for name, is_attended, has_highlight in combined_list:
                display_name = name[:4] if len(name) > 4 else name
                highlight_class = 'highlight-green' if is_attended and has_highlight else \
                                 'highlight-red' if not is_attended and has_highlight else ''
                entry = (name, display_name, highlight_class)
                if is_attended:
                    attended_with_highlights.append(entry)
                else:
                    not_attended_with_highlights.append(entry)

            sorted_attended[district] = attended_with_highlights
            sorted_not_attended[district] = not_attended_with_highlights

        html += f'<div class="district-section">\n<h2>{main_district} - {week_display}</h2>\n<div class="district-container flex-container">\n'
        html += '<div class="table-wrapper attendance-wrapper flex-item">\n<table class="excel-table">\n'
        total_cols = len(sub_districts) * 2
        html += f'<tr class="header"><th colspan="{total_cols}">{main_district}</th></tr>\n<tr class="district-row">\n'
        for district in sub_districts:
            html += f'<th colspan="2">{district}</th>'
        html += '</tr>\n<tr class="subheader">\n'
        for _ in sub_districts:
            html += '<th>本週到會</th><th>未到會</th>'
        html += '</tr>\n'

        for r in range(max_len):
            row_class = "even" if r % 2 == 0 else "odd"
            html += f'<tr class="{row_class}">\n'
            for district in sub_districts:
                attended_info = sorted_attended.get(district, [])[r] if r < len(sorted_attended.get(district, [])) else ('', '', '')
                not_attended_info = sorted_not_attended.get(district, [])[r] if r < len(sorted_not_attended.get(district, [])) else ('', '', '')
                html += f'<td class="{attended_info[2]}">{attended_info[1]}</td><td class="{not_attended_info[2]}">{not_attended_info[1]}</td>'
            html += '</tr>\n'
        html += '</table>\n</div>\n'

        html += render_stats_table(main_district, district_counts, main_district_counts)
        html += '</div>\n</div>\n'

    if not html:
        html = '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>該週無有效資料</th></tr></table></div>'
    return html

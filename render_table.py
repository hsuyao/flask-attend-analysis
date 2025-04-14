from config import logger
from utils import chinese_to_int, parse_district
import sqlite3
from datetime import datetime
from database import DATABASE_PATH

def get_all_latest_attendance_dates(names, latest_date):
    """批量獲取多人的最近出勤日期，若無記錄則返回遠古日期"""
    if not names:
        return {}
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    placeholders = ','.join('?' for _ in names)
    cursor.execute(f'''
        SELECT name, MAX(date) FROM attendance_records
        WHERE name IN ({placeholders}) AND attended = 1 AND date <= ?
        GROUP BY name
    ''', names + [latest_date.strftime('%Y-%m-%d')])
    results = {row[0]: datetime.strptime(row[1], '%Y-%m-%d') if row[1] else datetime(1970, 1, 1) for row in cursor.fetchall()}
    conn.close()
    return {name: results.get(name, datetime(1970, 1, 1)) for name in names}

def render_stats_table(main_district, district_counts, main_district_counts):
    """生成統計表 HTML（僅包含子區統計，移除主區統計）"""
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']
    html = ""

    stats_districts = sorted([d for d in district_counts.keys() if d != '總計'], key=parse_district)
    sub_districts_stats = [d for d in stats_districts if d.startswith(main_district)]
    if sub_districts_stats:
        html += '<div class="table-wrapper stats-wrapper flex-item">\n<table class="excel-table">\n'
        html += f'<tr class="header"><th colspan="2">{main_district} 統計</th></tr>\n'
        row_index = 0
        
        for district in sub_districts_stats:
            total = district_counts[district]['total']
            html += f'<tr class="total-row"><td style="padding-left: 15px;">{district}</td><td>{total}</td></tr>\n'
            row_index += 1
            for age in age_categories:
                count = district_counts[district]['ages'][age]
                row_class = "even" if row_index % 2 == 0 else "odd"
                html += f'<tr class="{row_class}"><td style="padding-left: 30px;">{age}</td><td>{count}</td></tr>\n'
                row_index += 1
        
        html += '</table>\n</div>\n'

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
        return """
        <div class="district-section">
            <table class="excel-table">
                <tr class="title-row"><th>無資料</th></tr>
            </table>
        </div>
        """

    previous_week_data = None
    all_attendance_data.sort(key=lambda x: x[0])
    current_week_idx = next((idx for idx, (date, data, week_name) in enumerate(all_attendance_data) if week_name == week_display), None)
    
    if current_week_idx is not None and current_week_idx > 0:
        previous_week_data = all_attendance_data[current_week_idx - 1][1]
        logger.debug(f"找到前一週資料，週次 {week_display}: {previous_week_data}")
    else:
        logger.warning(f"沒有前一週資料可用，週次: {week_display}。高亮顯示將被停用。")
    
    latest_date = all_attendance_data[-1][0] if all_attendance_data else datetime.now()

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
            
            combined_list = []
            if previous_week_data:
                prev_attended = previous_week_data['attended'].get(district, [])
                prev_not_attended = previous_week_data['not_attended'].get(district, [])
                
                logger.debug(f"地區: {district}, 前一週出勤: {prev_attended}, 前一週未出勤: {prev_not_attended}")
                
                for name in attended_list:
                    has_highlight = name in prev_not_attended
                    combined_list.append((name, True, has_highlight))
                    logger.debug(f"姓名: {name}, 地區: {district}, 出勤: 是, 是否高亮: {has_highlight}")
                
                for name in not_attended_list:
                    has_highlight = name in prev_attended
                    combined_list.append((name, False, has_highlight))
                    logger.debug(f"姓名: {name}, 地區: {district}, 出勤: 否, 是否高亮: {has_highlight}")
            else:
                combined_list = [(name, True, False) for name in attended_list] + \
                               [(name, False, False) for name in not_attended_list]
                logger.debug("無前一週資料，所有高亮狀態設為否")
            
            names = attended_list + not_attended_list
            latest_dates = get_all_latest_attendance_dates(names, latest_date)
            
            combined_list.sort(key=lambda x: (-int(x[2]), latest_dates[x[0]]), reverse=True)
            
            logger.debug(f"{district} 排序後的合併名單: {[(name, is_attended, has_highlight) for name, is_attended, has_highlight in combined_list]}")
            
            highlighted_attended = []
            highlighted_not_attended = []
            non_highlighted_attended = []
            non_highlighted_not_attended = []
            
            for name, is_attended, has_highlight in combined_list:
                display_name = name[:4] if len(name) > 4 else name
                highlight_class = 'highlight-green' if is_attended and has_highlight else \
                                'highlight-red' if not is_attended and has_highlight else ''
                entry = (name, display_name, highlight_class)
                if has_highlight:
                    if is_attended:
                        highlighted_attended.append(entry)
                    else:
                        highlighted_not_attended.append(entry)
                else:
                    if is_attended:
                        non_highlighted_attended.append(entry)
                    else:
                        non_highlighted_not_attended.append(entry)
            
            attended_with_highlights = highlighted_attended + non_highlighted_attended
            not_attended_with_highlights = highlighted_not_attended + non_highlighted_not_attended
            
            logger.debug(f"{district} 出勤高亮名單: {attended_with_highlights}")
            logger.debug(f"{district} 未出勤高亮名單: {not_attended_with_highlights}")
            
            sorted_attended[district] = attended_with_highlights
            sorted_not_attended[district] = not_attended_with_highlights

        html += f'<div class="district-section">\n'
        html += f'<h2>{main_district} - {week_display}</h2>\n'
        html += '<div class="district-container flex-container">\n'

        # 出勤表（左側）
        html += '<div class="table-wrapper attendance-wrapper flex-item">\n<table class="excel-table">\n'
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

        # 統計表（右側）
        html += render_stats_table(main_district, district_counts, main_district_counts)

        html += '</div>\n</div>\n'

    if not html:
        html = """
        <div class="district-section">
            <table class="excel-table">
                <tr class="title-row"><th>該週無有效資料</th></tr>
            </table>
        </div>
        """

    return html

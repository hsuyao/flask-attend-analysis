from config import logger, db, DB_OFFLINE, COLLECTION_NAME
from utils import chinese_to_int, parse_district, parse_week_display
from database import (
    get_six_month_averages,
    get_all_latest_attendance_dates,
    get_six_month_trimmed_mean_by_event,
    get_six_month_trimmed_mean_by_age_group,
)
from datetime import datetime
from datetime import timedelta

def render_stats_table(main_district, district_counts, main_district_counts, event_name="未指定活動", is_history_page=False, event_totals=None):
    age_categories = ['青職以上', '大專', '中學', '小學', '學齡前']
    stats_districts = sorted(
        [d for d in district_counts.keys() if d != '總計'],
        key=parse_district
    )
    sub_districts = [d for d in stats_districts if d.startswith(main_district)]

    if not sub_districts and main_district not in main_district_counts:
        return '<div class="table-wrapper stats-wrapper flex-item"><p>無統計資料</p></div>'

    # Calculate number of districts for dynamic width
    districts = [main_district] + sub_districts
    num_districts = len(districts)

    # Set CSS custom property for dynamic width
    html = f'<div class="table-wrapper stats-wrapper flex-item" style="--num-districts: {num_districts};">\n<table class="excel-table">\n'
    html += f'<tr class="header"><th>{event_name}</th>'
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
    html += '</tr>\n'

    if is_history_page and event_totals is not None:
        logger.debug(f"Rendering event totals for main_district: {main_district}, event_totals: {event_totals}")
        # Add spacer row for 0.25 line spacing
        html += f'<tr class="spacer"><td colspan="{num_districts + 1}"></td></tr>\n'
        
        # Add rows for 禱告, 晨興, 小排 if they have data
        for event in ["禱告", "晨興", "小排"]:
            total = event_totals.get(event, {}).get("total", 0)
            if total > 0:
                html += f'<tr class="event-row"><td>{event}</td>'
                for district in districts:
                    count = total if district == main_district else \
                            event_totals[event]["districts"].get(district, 0)
                    html += f'<td>{count}</td>'
                html += '</tr>\n'
            else:
                logger.debug(f"Skipping event {event} for {main_district}: total count is {total}")

    html += '</table>\n</div>\n'
    return html

def render_attendance_table(week_display, latest_attendance_data, all_attendance_data, district_counts, main_district_counts, avg_attendance_rates=None, event_name="未指定活動", is_history_page=False, event_totals=None):
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
    all_attendance_data.sort(key=lambda x: parse_week_display(x[2]))  # Sort by parsed week_display
    current_week_idx = next((idx for idx, (_, _, week_name, _) in enumerate(all_attendance_data) if week_name == week_display), None)

    if current_week_idx is not None and current_week_idx > 0:
        previous_week_data = all_attendance_data[current_week_idx - 1][1]
        logger.debug(f"Previous week data for {week_display}: {previous_week_data}")

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
        num_sub_districts = len(sub_districts)  # Number of sub-districts
        total_cols = num_sub_districts * 2  # Each sub-district has 2 columns
        # Calculate number of districts for stats table
        stats_districts = sorted(
            [d for d in district_counts.keys() if d != '總計' and d.startswith(main_district)],
            key=parse_district
        )
        num_districts = len([main_district] + stats_districts)

        for district in sub_districts:
            attended_list = latest_attendance_data['attended'].get(district, [])
            not_attended_list = latest_attendance_data['not_attended'].get(district, [])
            combined_list = []

            if previous_week_data:
                prev_attended = previous_week_data['attended'].get(district, [])
                prev_not_attended = previous_week_data['not_attended'].get(district, [])
                logger.debug(f"District {district}: prev_attended={prev_attended}, prev_not_attended={prev_not_attended}")
                combined_list.extend((name, True, name in prev_not_attended) for name in attended_list)
                combined_list.extend((name, False, name in prev_attended) for name in not_attended_list)
            else:
                combined_list.extend((name, True, False) for name in attended_list)
                combined_list.extend((name, False, False) for name in not_attended_list)

            names = [name for name, _, _ in combined_list]

            # ───────────────────────────────────────────────
            # 離線模式：避免觸發 MongoDB，直接給空字串
            # ───────────────────────────────────────────────
            if DB_OFFLINE:
                latest_dates = {n: '' for n in names}
            else:
                latest_dates = get_all_latest_attendance_dates(
                    names, placeholder_date
                )
            logger.debug(f"District {district}: Attendance rates {[(name, avg_attendance_rates.get(name, 0.0)) for name in names]}")
            logger.debug(f"District {district}: Highlights {[(name, has_highlight) for name, _, has_highlight in combined_list]}")
            # Sort by highlight status (True first), then attendance rate (descending), then latest date (parsed)
            combined_list.sort(
                key=lambda x: (
                    -int(x[2]),                           # Primary: highlight (True first)
                    -avg_attendance_rates.get(x[0], 0.0),  # Secondary: descending attendance rate
                    parse_week_display(latest_dates.get(x[0], ''))  # Tertiary: parsed week_display
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

        html += f'<div class="district-section">\n<h2>{main_district} - {week_display}</h2>\n<div class="district-container flex-container" style="--num-sub-districts: {num_sub_districts}; --num-districts: {num_districts};">\n'
        html += f'<div class="table-wrapper attendance-wrapper flex-item" style="--num-sub-districts: {num_sub_districts};">\n<table class="excel-table">\n'
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

        html += render_stats_table(main_district, district_counts, main_district_counts, event_name, is_history_page, event_totals)
        html += '</div>\n</div>\n'

    if not html:
        html = '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>該週無有效資料</th></tr></table></div>'
    return html


def render_average_attendance_table(main_district, end_date, trimmed_mean_data_list, age_group_data):
    """
    Build the HTML table that shows each person’s 6-month average attendance rate
    for the given `main_district`, using `end_date` as the right edge of the window.

    Parameters
    ----------
    main_district : str
        The target main district (e.g. "二大區").
    end_date : datetime.datetime
        The end date chosen by the user; the query looks back 180 days.
    trimmed_mean_data_list : list[dict]
        Output of `get_six_month_trimmed_mean_by_event`, already filtered
        for the same main district and date window.
    age_group_data : dict
        Output of `get_six_month_trimmed_mean_by_age_group` for the same
        main district and date window.

    Returns
    -------
    str
        Rendered HTML snippet.
    """
    # ------------------------------------------------------------------
    # 1. Compute the 6-month window [start_date, end_date]
    # ------------------------------------------------------------------
    six_months_ago = end_date - timedelta(days=180)
    start_date_str = six_months_ago.strftime('%Y-%m-%d')
    end_date_str   = end_date.strftime('%Y-%m-%d')

    # ------------------------------------------------------------------
    # 2. Query MongoDB for each person’s average attendance over the window
    # ------------------------------------------------------------------
    pipeline = [
        {"$match": {
            "district":  {"$regex": f"^{main_district}"},
            "event_name": "主日",
            "date":      {"$gte": start_date_str, "$lte": end_date_str}
        }},
        {"$group": {
            "_id": {"district": "$district", "name": "$name"},
            "attendance_rate": {
                "$avg": {"$cond": [{"$eq": ["$attended", 1]}, 1, 0]}
            }
        }}
    ]
    name_records = list(db[COLLECTION_NAME].aggregate(pipeline))

    # Collect sub-districts under the main district and sort them
    districts = sorted(
        {r["_id"]["district"]
         for r in name_records
         if r["_id"]["district"].startswith(main_district)},
        key=parse_district
    )
    if not districts:
        return ('<div class="district-section"><table class="excel-table">'
                '<tr class="title-row"><th>No data</th></tr></table></div>')

    # Map each sub-district to a sorted list of (name, rate) pairs
    district_names = {d: [] for d in districts}
    for rec in name_records:
        d = rec["_id"]["district"]
        if d in district_names:
            district_names[d].append((rec["_id"]["name"], rec["attendance_rate"]))
    for d in district_names:
        district_names[d].sort(key=lambda x: -x[1])          # descending

    # ------------------------------------------------------------------
    # 3. Build the left-hand table (names + rates)
    # ------------------------------------------------------------------
    max_len            = max(len(lst) for lst in district_names.values())
    num_sub_districts  = len(districts)
    total_cols         = num_sub_districts * 2

    html = (
        '<div class="district-section">\n'
        f'<h2>{main_district} — 6-Month Avg Attendance (through {end_date_str})</h2>\n'
        f'<div class="district-container flex-container" '
        f'style="--num-sub-districts:{num_sub_districts}; --num-districts:{num_sub_districts + 1};">\n'
        f'<div class="table-wrapper attendance-wrapper flex-item" '
        f'style="--num-sub-districts:{num_sub_districts};">\n'
        '<table class="excel-table">\n'
        f'<tr class="header"><th colspan="{total_cols}">{main_district}</th></tr>\n'
        '<tr class="district-row">\n'
    )
    for d in districts:
        html += f'<th colspan="2">{d}</th>'
    html += '</tr>\n<tr class="subheader">\n'
    html += ''.join('<th>姓名</th><th>到會率</th>' for _ in districts)
    html += '</tr>\n'

    for r in range(max_len):
        row_cls = "even" if r % 2 == 0 else "odd"
        html += f'<tr class="{row_cls}">\n'
        for d in districts:
            name, rate = district_names[d][r] if r < len(district_names[d]) else ('', 0.0)
            display    = name[:4] if len(name) > 4 else name
            html += f'<td>{display}</td><td>{rate:.0%}</td>'
        html += '</tr>\n'
    html += '</table>\n</div>\n'

    # ------------------------------------------------------------------
    # 4. Build the right-hand stats table (trimmed means for each event)
    # ------------------------------------------------------------------
    num_districts = num_sub_districts + 2  # main + subs + label
    html += (
        f'<div class="table-wrapper stats-wrapper flex-item" '
        f'style="--num-districts:{num_districts};">\n'
        '<table class="excel-table">\n'
        '<tr class="header"><th style="min-width:2em">Avg</th>'
        f'<th style="min-width:1em">{main_district}</th>'
    )
    for d in districts:
        html += f'<th style="min-width:1em">{d}</th>'
    html += '</tr>\n'

    row_index = 0
    # Each row represents one event: 主日, 禱告, 晨興, 小排…
    for data in trimmed_mean_data_list:
        if not any(data["districts"].values()) and not data["counts"].get(main_district):
            continue
        row_cls = "even" if row_index % 2 == 0 else "odd"
        html += (
            f'<tr class="{row_cls}"><td style="min-width:2em">{data["event_name"]}</td>'
            f'<td style="min-width:1em">{data["counts"].get(main_district, 0)}</td>'
        )
        for d in districts:
            html += f'<td style="min-width:1em">{data["districts"].get(d, 0)}</td>'
        html += '</tr>\n'
        row_index += 1

    age_categories = ['青職以上', '大專', '中學', '小學', '學齡前']
    for age in age_categories:
        row_cls = "even" if row_index % 2 == 0 else "odd"
        html += (
            f'<tr class="{row_cls}"><td style="min-width:2em">{age}</td>'
            f'<td style="min-width:1em">{age_group_data.get(age, {}).get(main_district, 0)}</td>'
        )
        for d in districts:
            html += f'<td style="min-width:1em">{age_group_data.get(age, {}).get(d, 0)}</td>'
        html += '</tr>\n'
        row_index += 1

    html += '</table>\n</div>\n'   # end stats-wrapper
    html += '</div>\n</div>\n'     # end container / section
    return html

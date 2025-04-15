import os
import subprocess
import tempfile
from io import BytesIO
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import sqlite3
from config import logger, START_COLUMN
from utils import chinese_to_int, parse_district
from database import get_six_month_average, DATABASE_PATH

def convert_xls_to_xlsx(file_stream):
    logger.info("正在將 .xls 轉換為 .xlsx 使用 soffice")
    file_stream.seek(0)
    file_content = file_stream.read()

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xls') as temp_xls:
        temp_xls.write(file_content)
        temp_xls_path = temp_xls.name

    temp_xlsx_path = temp_xls_path.replace('.xls', '.xlsx')

    try:
        result = subprocess.run([
            'soffice',
            '--headless',
            '--convert-to',
            'xlsx',
            temp_xls_path,
            '--outdir',
            os.path.dirname(temp_xls_path)
        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        logger.info(f"成功將 {temp_xls_path} 轉換為 {temp_xlsx_path}")

        if not os.path.exists(temp_xlsx_path):
            logger.error("轉換後的 .xlsx 檔案未找到")
            raise Exception("轉換失敗：未找到輸出檔案")

        with open(temp_xlsx_path, 'rb') as temp_xlsx:
            output_stream = BytesIO(temp_xlsx.read())
        return output_stream

    except subprocess.CalledProcessError as e:
        logger.error(f"無法將 .xls 轉換為 .xlsx: {e.stderr.decode()}")
        raise Exception(f"無法將 .xls 轉換為 .xlsx: {e.stderr.decode()}")
    except Exception as e:
        logger.error(f"轉換過程中發生未知錯誤: {str(e)}")
        raise
    finally:
        if os.path.exists(temp_xls_path):
            os.remove(temp_xls_path)
        if os.path.exists(temp_xlsx_path):
            os.remove(temp_xlsx_path)

def classify_attendance(sheet, week_col, week_date):
    main_district = None
    logger.debug(f"正在為週次欄位分類出勤: {week_col}")
    attended = {}
    not_attended = {}
    district_counts = {}
    main_district_counts = {}
    records = []  # 收集数据库记录
    youth_above = {'年長', '中壯', '青壯', '青職'}
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']
    max_row = sheet.max_row
    
    for row in range(3, max_row + 1):
        main_district_value = str(sheet.cell(row, 1).value or "").strip()
        sub_district = str(sheet.cell(row, 2).value or "").strip()
        district = f"{main_district_value}{sub_district}"
        name = sheet.cell(row, 4).value
        age = str(sheet.cell(row, 6).value or "").strip()
        if not name or not district.startswith(main_district_value):
            continue
        if main_district is None and main_district_value:
            main_district = main_district_value
            logger.debug(f"設定主區名稱為: {main_district}")
        attendance = sheet.cell(row, week_col + 1).value
        effective_age = '青職以上' if age in youth_above or not age else age
        if effective_age not in age_categories:
            logger.warning(f"無法識別年齡 '{age}' 對於 {name} 在 {district}，預設為 '青職以上'")
            effective_age = '青職以上'
        
        # 收集记录用于批量插入
        records.append((name, week_date.strftime('%Y-%m-%d'), 1 if attendance == 1 else 0, district, effective_age))
        
        if attendance == 1:
            if district not in attended:
                attended[district] = []
            attended[district].append(name)
            if district not in district_counts:
                district_counts[district] = {'total': 0, 'ages': {age: 0 for age in age_categories}}
            if main_district_value not in main_district_counts:
                main_district_counts[main_district_value] = {'total': 0, 'ages': {age: 0 for age in age_categories}}
            district_counts[district]['total'] += 1
            main_district_counts[main_district_value]['total'] += 1
            district_counts[district]['ages'][effective_age] += 1
            main_district_counts[main_district_value]['ages'][effective_age] += 1
        else:
            if district not in not_attended:
                not_attended[district] = []
            not_attended[district].append(name)
    
    total_attendance = sum(d['total'] for d in district_counts.values())
    district_counts['總計'] = total_attendance
    return attended, not_attended, district_counts, main_district, main_district_counts, records

def get_all_latest_attendance_dates(names, latest_date, cursor):
    """批量獲取多人的最近出勤日期，若無記錄則返回遠古日期"""
    if not names:
        return {}
    placeholders = ','.join('?' for _ in names)
    cursor.execute(f'''
        SELECT name, MAX(date) FROM attendance_records
        WHERE name IN ({placeholders}) AND attended = 1 AND date <= ?
        GROUP BY name
    ''', list(names) + [latest_date.strftime('%Y-%m-%d')])
    results = {row[0]: datetime.strptime(row[1], '%Y-%m-%d') if row[1] else datetime(1970, 1, 1) for row in cursor.fetchall()}
    return {name: results.get(name, datetime(1970, 1, 1)) for name in names}

def write_summary(new_sheet, attended, not_attended, latest_date, previous_week_data=None):
    logger.debug(f"正在寫入總覽，出席: {attended}, 未出席: {not_attended}")
    districts = sorted(set(attended.keys()).union(not_attended.keys()), key=lambda x: (chinese_to_int(x[0]), chinese_to_int(x[3:4])))
    row = 1

    header_fill = PatternFill(start_color="107C10", end_color="107C10", fill_type="solid")
    subheader_fill = PatternFill(start_color="5DBB63", end_color="5DBB63", fill_type="solid")
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    red_fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for i, district in enumerate(districts):
        cell1 = new_sheet.cell(row, i * 3 + 1)
        cell2 = new_sheet.cell(row, i * 3 + 2)
        cell3 = new_sheet.cell(row, i * 3 + 3)
        cell1.value = district
        cell2.value = district
        cell3.value = district
        cell1.fill = header_fill
        cell2.fill = header_fill
        cell3.fill = header_fill
        cell1.font = header_font
        cell2.font = header_font
        cell3.font = header_font
        cell1.alignment = Alignment(horizontal='center')
        cell2.alignment = Alignment(horizontal='center')
        cell3.alignment = Alignment(horizontal='center')

        sub_cell1 = new_sheet.cell(row + 1, i * 3 + 1)
        sub_cell2 = new_sheet.cell(row + 1, i * 3 + 2)
        sub_cell3 = new_sheet.cell(row + 1, i * 3 + 3)
        sub_cell1.value = "本週到會"
        sub_cell2.value = "未到會"
        sub_cell3.value = "半年平均出勤率"
        sub_cell1.fill = subheader_fill
        sub_cell2.fill = subheader_fill
        sub_cell3.fill = subheader_fill
        sub_cell1.font = header_font
        sub_cell2.font = header_font
        sub_cell3.font = header_font
        sub_cell1.alignment = Alignment(horizontal='center')
        sub_cell2.alignment = Alignment(horizontal='center')
        sub_cell3.alignment = Alignment(horizontal='center')

    max_len = max(max(len(attended.get(d, [])), len(not_attended.get(d, []))) for d in districts)

    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    try:
        for i, district in enumerate(districts):
            attended_list = attended.get(district, [])
            not_attended_list = not_attended.get(district, [])

            # 合併名單，記錄底色
            combined_list = []
            if previous_week_data:
                prev_attended = previous_week_data['attended'].get(district, [])
                prev_not_attended = previous_week_data['not_attended'].get(district, [])
                for name in attended_list:
                    has_highlight = name in prev_not_attended
                    combined_list.append((name, True, has_highlight))
                for name in not_attended_list:
                    has_highlight = name in prev_attended
                    combined_list.append((name, False, has_highlight))
            else:
                combined_list = [(name, True, False) for name in attended_list] + \
                               [(name, False, False) for name in not_attended_list]

            # 批量查詢最近出勤日期
            names = attended_list + not_attended_list
            latest_dates = get_all_latest_attendance_dates(names, latest_date, cursor)

            # 排序：底色優先 + 最近出勤日期降序
            combined_list.sort(key=lambda x: (-int(x[2]), latest_dates[x[0]]), reverse=True)

            # 分列：將有底色的姓名放在最前
            highlighted_attended = []
            highlighted_not_attended = []
            non_highlighted_attended = []
            non_highlighted_not_attended = []

            for name, is_attended, has_highlight in combined_list:
                if has_highlight:
                    if is_attended:
                        highlighted_attended.append((name, has_highlight))
                    else:
                        highlighted_not_attended.append((name, has_highlight))
                else:
                    if is_attended:
                        non_highlighted_attended.append((name, has_highlight))
                    else:
                        non_highlighted_not_attended.append((name, has_highlight))

            attended_with_highlights = highlighted_attended + non_highlighted_attended
            not_attended_with_highlights = highlighted_not_attended + non_highlighted_not_attended

            for r in range(max_len):
                # 出勤列
                if r < len(attended_with_highlights):
                    name, has_highlight = attended_with_highlights[r]
                    cell = new_sheet.cell(r + 3, i * 3 + 1)
                    cell.value = name
                    if has_highlight:
                        cell.fill = green_fill
                # 未出勤列
                if r < len(not_attended_with_highlights):
                    name, has_highlight = not_attended_with_highlights[r]
                    cell = new_sheet.cell(r + 3, i * 3 + 2)
                    cell.value = name
                    if has_highlight:
                        cell.fill = red_fill
                    # 半年平均出勤率（僅未出勤的人需要額外計算）
                    avg_rate = get_six_month_average(name, latest_date)
                    new_sheet.cell(r + 3, i * 3 + 3).value = f"{avg_rate:.2%}"
                # 出勤的人的出勤率（在未出勤列之後補充）
                if r < len(attended_with_highlights):
                    name, _ = attended_with_highlights[r]
                    avg_rate = get_six_month_average(name, latest_date)
                    new_sheet.cell(r + 3, i * 3 + 3).value = f"{avg_rate:.2%}"
    finally:
        conn.close()

    logger.debug("總覽寫入成功")

def process_excel(file_stream, file_extension):
    file_stream.seek(0)
    file_content = file_stream.read()
    buffered_stream = BytesIO(file_content)
    logger.info(f"正在處理檔案，副檔名: {file_extension}，大小: {len(file_content)} 位元組")

    if file_extension == '.xls':
        logger.info("檢測到 .xls 檔案，正在轉換為 .xlsx")
        buffered_stream = convert_xls_to_xlsx(file_stream)
        file_extension = '.xlsx'

    try:
        workbook = openpyxl.load_workbook(buffered_stream)
    except Exception as e:
        logger.error(f"無法載入工作簿: {str(e)}")
        raise

    input_sheet = workbook.active
    logger.debug(f"已載入工作表: {input_sheet.title}，列數: {input_sheet.max_row}，欄數: {input_sheet.max_column}")

    week_cols = []
    current_month = "2025年1月"
    for col in range(START_COLUMN, input_sheet.max_column + 1):
        month_header = str(input_sheet.cell(1, col + 1).value or "")
        week_header = str(input_sheet.cell(2, col + 1).value or "")
        if "年" in month_header and "月" in month_header:
            current_month = month_header.strip()
        if "週" in week_header:
            week_cols.append((col, week_header, current_month))

    logger.info(f"檢測到週次欄位與月份: {week_cols}")

    if not week_cols:
        logger.warning("未檢測到週次欄位；輸出將缺少分析工作表")

    all_attendance_data = []
    latest_date = datetime(1970, 1, 1)
    latest_attended = None
    latest_not_attended = None
    latest_week = None
    latest_districts = None
    latest_main_district = None
    latest_main_district_counts = None
    all_records = []
    all_names = set()

    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    try:
        for col, week_name, month_prefix in week_cols:
            logger.info(f"正在處理週次: {week_name} 在 {month_prefix}")
            year = int(month_prefix.split("年")[0])
            month_part = month_prefix.split("年")[1]
            week_str = week_name.replace("第", "").replace("週", "")
            week_num = chinese_to_int(week_str)
            month_num = int(month_part.replace("月", ""))
            current_date = datetime(year, month_num, min(week_num * 7, 28))
            
            attended, not_attended, district_counts, main_district, main_district_counts, records = classify_attendance(input_sheet, col, current_date)
            all_records.extend(records)
            if main_district and not latest_main_district:
                latest_main_district = main_district

            for district in attended:
                all_names.update(attended[district])
            for district in not_attended:
                all_names.update(not_attended[district])

            if not any(attended.values()):
                logger.info(f"{week_name} 在 {month_prefix} 無出席者，跳過工作表建立和資料包含")
                continue

            all_attendance_data.append((current_date, {'attended': attended, 'not_attended': not_attended}, f"{month_prefix}{week_name}"))

            if current_date > latest_date:
                latest_date = current_date
                latest_attended = attended
                latest_not_attended = not_attended
                latest_week = f"{month_prefix}{week_name}"
                latest_districts = district_counts
                latest_main_district_counts = main_district_counts

        # 批量插入主记录
        if all_records:
            batch_size = 1000
            for i in range(0, len(all_records), batch_size):
                batch = all_records[i:i + batch_size]
                cursor.executemany('''
                    INSERT OR REPLACE INTO attendance_records (name, date, attended, district, age_group)
                    VALUES (?, ?, ?, ?, ?)
                ''', batch)
                conn.commit()
                logger.debug(f"插入批次 {i // batch_size + 1}，记录数 {len(batch)}")
            logger.info(f"批量插入 {len(all_records)} 条主记录")

        # 补全缺失记录
        if latest_date > datetime(1970, 1, 1):
            six_months_ago = latest_date - timedelta(days=180)
            cursor.execute('''
                SELECT name, date FROM attendance_records
                WHERE date >= ? AND date <= ?
            ''', (six_months_ago.strftime('%Y-%m-%d'), latest_date.strftime('%Y-%m-%d')))
            existing_records = {(row[0], row[1]) for row in cursor.fetchall()}
            
            weeks = []
            current_week = six_months_ago
            while current_week <= latest_date:
                weeks.append(current_week.strftime('%Y-%m-%d'))
                current_week += timedelta(days=7)
            
            missing_records = []
            for name in all_names:
                district = next((d for d in attended if name in attended[d]), None) or \
                          next((d for d in not_attended if name in not_attended[d]), None) or "未知區"
                age_group = "未知"
                for week in weeks:
                    if (name, week) not in existing_records:
                        missing_records.append((name, week, 0, district, age_group))
            
            if missing_records:
                batch_size = 1000
                for i in range(0, len(missing_records), batch_size):
                    batch = missing_records[i:i + batch_size]
                    cursor.executemany('''
                        INSERT OR REPLACE INTO attendance_records (name, date, attended, district, age_group)
                        VALUES (?, ?, ?, ?, ?)
                    ''', batch)
                    conn.commit()
                    logger.debug(f"插入缺失记录批次 {i // batch_size + 1}，记录数 {len(batch)}")
                logger.info(f"批量插入 {len(missing_records)} 条缺失记录")

        if not all_attendance_data:
            logger.warning("檔案中未找到有出席者的週次")
            return {
                'latest_analytic_date': None,
                'latest_attendance_data': None,
                'latest_week_display': None,
                'latest_district_counts': None,
                'latest_main_district': None,
                'latest_main_district_counts': None,
                'all_attendance_data': []
            }

    finally:
        conn.close()

    logger.info("檔案處理完成，未生成 Excel 檔案（延遲到下載時）")

    return {
        'latest_analytic_date': latest_date.strftime("%Y年%m月%d日") if latest_date > datetime(1970, 1, 1) else None,
        'latest_attendance_data': {'attended': latest_attended, 'not_attended': latest_not_attended} if latest_attended else None,
        'latest_week_display': latest_week,
        'latest_district_counts': latest_districts,
        'latest_main_district': latest_main_district,
        'latest_main_district_counts': latest_main_district_counts,
        'all_attendance_data': all_attendance_data
    }

def generate_excel(all_attendance_data):
    """動態生成 Excel 檔案，包含所有週次的總覽工作表"""
    workbook = openpyxl.Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    for date, data, week_name in all_attendance_data:
        year = date.year
        month = date.month
        new_sheet_name = f"{year}年{month}月{week_name.split('月')[1]} 主日"
        
        if new_sheet_name in workbook.sheetnames:
            logger.error(f"檢測到重複的工作表名稱: {new_sheet_name}")
            raise ValueError(f"工作表名稱 '{new_sheet_name}' 已存在")
        
        new_sheet = workbook.create_sheet(new_sheet_name)
        logger.debug(f"已建立新工作表: {new_sheet_name}")

        previous_week_data = None
        current_week_idx = next(idx for idx, (d, _, w) in enumerate(all_attendance_data) if d == date)
        if current_week_idx > 0:
            previous_week_data = all_attendance_data[current_week_idx - 1][1]

        write_summary(new_sheet, data['attended'], data['not_attended'], date, previous_week_data)

    output_stream = BytesIO()
    workbook.save(output_stream)
    output_stream.seek(0)
    return output_stream

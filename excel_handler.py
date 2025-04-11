import os
import subprocess
import tempfile
from io import BytesIO
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from config import logger, START_COLUMN
from utils import chinese_to_int

def convert_xls_to_xlsx(file_stream):
    """Convert .xls to .xlsx using soffice command."""
    logger.info("Converting .xls to .xlsx using soffice")
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
        logger.info(f"Successfully converted {temp_xls_path} to {temp_xlsx_path}")

        if not os.path.exists(temp_xlsx_path):
            logger.error("Converted .xlsx file not found after soffice conversion")
            raise Exception("Conversion failed: Output file not found")

        with open(temp_xlsx_path, 'rb') as temp_xlsx:
            output_stream = BytesIO(temp_xlsx.read())
        return output_stream

    except subprocess.CalledProcessError as e:
        logger.error(f"Failed to convert .xls to .xlsx: {e.stderr.decode()}")
        raise Exception(f"Failed to convert .xls to .xlsx: {e.stderr.decode()}")
    except Exception as e:
        logger.error(f"Unexpected error during conversion: {str(e)}")
        raise
    finally:
        if os.path.exists(temp_xls_path):
            os.remove(temp_xls_path)
        if os.path.exists(temp_xlsx_path):
            os.remove(temp_xlsx_path)

def classify_attendance(sheet, week_col):
    main_district = None
    logger.debug(f"Classifying attendance for week column: {week_col}")
    attended = {}
    not_attended = {}
    district_counts = {}
    main_district_counts = {}
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
            logger.debug(f"Set main district name to: {main_district}")
        attendance = sheet.cell(row, week_col + 1).value
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
            effective_age = '青職以上' if age in youth_above or not age else age
            if effective_age not in age_categories:
                logger.warning(f"Unrecognized age '{age}' for {name} in {district}, defaulting to '青職以上'")
                effective_age = '青職以上'
            district_counts[district]['ages'][effective_age] += 1
            main_district_counts[main_district_value]['ages'][effective_age] += 1
        else:
            if district not in not_attended:
                not_attended[district] = []
            not_attended[district].append(name)
    total_attendance = sum(d['total'] for d in district_counts.values())
    district_counts['總計'] = total_attendance
    return attended, not_attended, district_counts, main_district, main_district_counts

def write_summary(new_sheet, attended, not_attended):
    logger.debug(f"Writing summary with attended: {attended}, not_attended: {not_attended}")
    districts = sorted(set(attended.keys()).union(not_attended.keys()), key=lambda x: chinese_to_int(x[3:4]))
    row = 1

    header_fill = PatternFill(start_color="107C10", end_color="107C10", fill_type="solid")
    subheader_fill = PatternFill(start_color="5DBB63", end_color="5DBB63", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for i, district in enumerate(districts):
        cell1 = new_sheet.cell(row, i * 2 + 1)
        cell2 = new_sheet.cell(row, i * 2 + 2)
        cell1.value = district
        cell2.value = district
        cell1.fill = header_fill
        cell2.fill = header_fill
        cell1.font = header_font
        cell2.font = header_font
        cell1.alignment = Alignment(horizontal='center')
        cell2.alignment = Alignment(horizontal='center')

        sub_cell1 = new_sheet.cell(row + 1, i * 2 + 1)
        sub_cell2 = new_sheet.cell(row + 1, i * 2 + 2)
        sub_cell1.value = "本週到會"
        sub_cell2.value = "未到會"
        sub_cell1.fill = subheader_fill
        sub_cell2.fill = subheader_fill
        sub_cell1.font = header_font
        sub_cell2.font = header_font
        sub_cell1.alignment = Alignment(horizontal='center')
        sub_cell2.alignment = Alignment(horizontal='center')

    max_len = max(max(len(attended.get(d, [])), len(not_attended.get(d, []))) for d in districts)
    for r in range(max_len):
        for i, district in enumerate(districts):
            attended_list = attended.get(district, [])
            not_attended_list = not_attended.get(district, [])
            if r < len(attended_list):
                new_sheet.cell(r + 3, i * 2 + 1).value = attended_list[r]
            if r < len(not_attended_list):
                new_sheet.cell(r + 3, i * 2 + 2).value = not_attended_list[r]

    logger.debug("Summary written successfully")

def process_excel(file_stream, file_extension):
    file_stream.seek(0)
    file_content = file_stream.read()
    buffered_stream = BytesIO(file_content)
    logger.info(f"Processing file with extension: {file_extension}, Size: {len(file_content)} bytes")

    if file_extension == '.xls':
        logger.info("Detected .xls file, converting to .xlsx")
        buffered_stream = convert_xls_to_xlsx(file_stream)
        file_extension = '.xlsx'

    try:
        workbook = openpyxl.load_workbook(buffered_stream)
    except Exception as e:
        logger.error(f"Failed to load workbook: {str(e)}")
        raise

    input_sheet = workbook.active
    logger.debug(f"Loaded sheet: {input_sheet.title}, Rows: {input_sheet.max_row}, Columns: {input_sheet.max_column}")

    week_cols = []
    current_month = "2025年1月"
    for col in range(START_COLUMN, input_sheet.max_column + 1):
        month_header = str(input_sheet.cell(1, col + 1).value or "")
        week_header = str(input_sheet.cell(2, col + 1).value or "")
        if "年" in month_header and "月" in month_header:
            current_month = month_header.strip()
        if "週" in week_header:
            week_cols.append((col, week_header, current_month))

    logger.info(f"Detected week columns with months: {week_cols}")

    if not week_cols:
        logger.warning("No week columns detected; output will lack analytic sheets")

    all_attendance_data = []
    latest_date = None
    latest_attended = None
    latest_not_attended = None
    latest_week = None
    latest_districts = None
    latest_main_district = None
    latest_main_district_counts = None
    for col, week_name, month_prefix in week_cols:
        logger.info(f"Processing week: {week_name} in {month_prefix}")
        attended, not_attended, district_counts, main_district, main_district_counts = classify_attendance(input_sheet, col)
        if main_district and not latest_main_district:
            latest_main_district = main_district

        if not any(attended.values()):
            logger.info(f"No attendees for {week_name} in {month_prefix}, skipping sheet creation and data inclusion")
            continue  # 跳過無人出席的週，不加入 all_attendance_data

        # 提取年份並生成唯一的工作表名稱
        year = int(month_prefix.split("年")[0])
        month_part = month_prefix.split("年")[1]
        week_str = week_name.replace("第", "").replace("週", "")
        week_num = chinese_to_int(week_str)
        month_num = int(month_part.replace("月", ""))
        current_date = datetime(year, month_num, min(week_num * 7, 28))
        new_sheet_name = f"{year}年{month_part}{week_name} 主日"

        all_attendance_data.append((current_date, {'attended': attended, 'not_attended': not_attended}, f"{month_prefix}{week_name}"))

        if latest_date is None or current_date > latest_date:
            latest_date = current_date
            latest_attended = attended
            latest_not_attended = not_attended
            latest_week = f"{month_prefix}{week_name}"
            latest_districts = district_counts
            latest_main_district_counts = main_district_counts

        if new_sheet_name in workbook.sheetnames:
            logger.error(f"Duplicate sheet name detected: {new_sheet_name}")
            raise ValueError(f"Sheet name '{new_sheet_name}' already exists")

        new_sheet = workbook.create_sheet(new_sheet_name)
        logger.debug(f"Created new sheet: {new_sheet_name}")
        write_summary(new_sheet, attended, not_attended)

    if not all_attendance_data:
        logger.warning("No weeks with attendees found in the file")
        return {
            'output_stream': BytesIO(),  # 返回空的輸出流
            'latest_analytic_date': None,
            'latest_attendance_data': None,
            'latest_week_display': None,
            'latest_district_counts': None,
            'latest_main_district': None,
            'latest_main_district_counts': None,
            'all_attendance_data': []
        }

    output_stream = BytesIO()
    workbook.save(output_stream)
    output_stream.seek(0)
    logger.info("File processing completed successfully")

    return {
        'output_stream': output_stream,
        'latest_analytic_date': latest_date.strftime("%Y年%m月%d日") if latest_date else None,
        'latest_attendance_data': {'attended': latest_attended, 'not_attended': latest_not_attended} if latest_attended else None,
        'latest_week_display': latest_week,
        'latest_district_counts': latest_districts,
        'latest_main_district': latest_main_district,
        'latest_main_district_counts': latest_main_district_counts,
        'all_attendance_data': all_attendance_data
    }

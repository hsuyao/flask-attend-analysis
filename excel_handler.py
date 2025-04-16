import os
import subprocess
import tempfile
from io import BytesIO
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from config import logger, START_COLUMN, db, COLLECTION_NAME
from utils import chinese_to_int, parse_district
from database import get_six_month_averages
from pymongo import UpdateOne, IndexModel, ASCENDING
import time

# Define batch size for writes
BATCH_SIZE = 500

def ensure_indexes():
    """Ensure necessary indexes for performance"""
    indexes = [
        IndexModel([("name", ASCENDING), ("date", ASCENDING), ("attended", ASCENDING)], name="name_date_attended_idx"),
        IndexModel([("date", ASCENDING), ("name", ASCENDING)], name="date_name_idx")
    ]
    try:
        db[COLLECTION_NAME].create_indexes(indexes)
        logger.info("Created/verified indexes for attendance_records")
    except Exception as e:
        logger.error(f"Failed to create indexes: {str(e)}")

def convert_xls_to_xlsx(file_stream):
    logger.info("Converting .xls to .xlsx using soffice")
    file_stream.seek(0)
    file_content = file_stream.read()

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xls') as temp_xls:
        temp_xls.write(file_content)
        temp_xls_path = temp_xls.name

    temp_xlsx_path = temp_xls_path.replace('.xls', '.xlsx')

    try:
        result = subprocess.run([
            'soffice', '--headless', '--convert-to', 'xlsx',
            temp_xls_path, '--outdir', os.path.dirname(temp_xls_path)
        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        logger.info(f"Converted {temp_xls_path} to {temp_xlsx_path}")

        if not os.path.exists(temp_xlsx_path):
            logger.error("Converted .xlsx file not found")
            raise Exception("Conversion failed: output file not found")

        with open(temp_xlsx_path, 'rb') as temp_xlsx:
            output_stream = BytesIO(temp_xlsx.read())
        return output_stream

    except subprocess.CalledProcessError as e:
        logger.error(f"Failed to convert .xls to .xlsx: {e.stderr.decode()}")
        raise Exception(f"Conversion failed: {e.stderr.decode()}")
    finally:
        if os.path.exists(temp_xls_path):
            os.remove(temp_xls_path)
        if os.path.exists(temp_xlsx_path):
            os.remove(temp_xlsx_path)

def classify_attendance(sheet, week_col, week_date):
    logger.info(f"Classifying attendance for column {week_col}")
    attended = {}
    not_attended = {}
    district_counts = {}
    main_district_counts = {}
    records = []
    youth_above = {'年長', '中壯', '青壯', '青職'}
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']
    max_row = sheet.max_row
    main_district = None

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

        attendance = sheet.cell(row, week_col + 1).value
        effective_age = '青職以上' if age in youth_above or not age else age
        if effective_age not in age_categories:
            effective_age = '青職以上'

        records.append({
            "name": name,
            "date": week_date.strftime('%Y-%m-%d'),
            "attended": 1 if attendance == 1 else 0,
            "district": district,
            "age_group": effective_age
        })

        if attendance == 1:
            attended.setdefault(district, []).append(name)
            district_counts.setdefault(district, {'total': 0, 'ages': {age: 0 for age in age_categories}})
            main_district_counts.setdefault(main_district_value, {'total': 0, 'ages': {age: 0 for age in age_categories}})
            district_counts[district]['total'] += 1
            main_district_counts[main_district_value]['total'] += 1
            district_counts[district]['ages'][effective_age] += 1
            main_district_counts[main_district_value]['ages'][effective_age] += 1
        else:
            not_attended.setdefault(district, []).append(name)

    total_attendance = sum(d['total'] for d in district_counts.values())
    district_counts['總計'] = total_attendance
    return attended, not_attended, district_counts, main_district, main_district_counts, records

def get_all_latest_attendance_dates(names, latest_date, cache=None):
    """Batch fetch latest attendance dates with caching"""
    if not names:
        return {}
    if cache is None:
        cache = {}
    
    # Return cached results if available
    names_to_query = [name for name in names if name not in cache]
    if not names_to_query:
        logger.debug(f"Cache hit for all {len(names)} names")
        return {name: cache[name] for name in names}

    start_time = time.time()
    try:
        pipeline = [
            {"$match": {
                "name": {"$in": names_to_query},
                "attended": 1,
                "date": {"$lte": latest_date.strftime('%Y-%m-%d')}
            }},
            {"$sort": {"date": -1}},
            {"$group": {
                "_id": "$name",
                "max_date": {"$first": "$date"}
            }}
        ]
        results = db[COLLECTION_NAME].aggregate(pipeline, hint="name_date_attended_idx")
        latest_dates = {name: datetime(1970, 1, 1) for name in names}
        for doc in results:
            if doc["max_date"]:
                date = datetime.strptime(doc["max_date"], '%Y-%m-%d')
                latest_dates[doc["_id"]] = date
                cache[doc["_id"]] = date
        
        elapsed = time.time() - start_time
        logger.debug(f"Retrieved latest dates for {len(names_to_query)} names in {elapsed:.2f}s")
        return latest_dates
    except Exception as e:
        logger.error(f"Failed to get latest attendance dates: {str(e)}")
        return {name: datetime(1970, 1, 1) for name in names}

def write_summary(new_sheet, attended, not_attended, latest_date, previous_week_data=None):
    logger.info("Writing Excel summary")
    districts = sorted(set(attended.keys()).union(not_attended.keys()), key=lambda x: (chinese_to_int(x[0]), chinese_to_int(x[3:4])))
    row = 1

    header_fill = PatternFill(start_color="107C10", end_color="107C10", fill_type="solid")
    subheader_fill = PatternFill(start_color="5DBB63", end_color="5DBB63", fill_type="solid")
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    red_fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for i, district in enumerate(districts):
        for col, value in enumerate([district, district, district], i * 3 + 1):
            cell = new_sheet.cell(row, col)
            cell.value = value
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

        for col, value in enumerate(["本週到會", "未到會", "半年平均出勤率"], i * 3 + 1):
            cell = new_sheet.cell(row + 1, col)
            cell.value = value
            cell.fill = subheader_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

    max_len = max(max(len(attended.get(d, [])), len(not_attended.get(d, []))) for d in districts)
    all_names = set()
    for district in districts:
        all_names.update(attended.get(district, []) + not_attended.get(district, []))
    avg_rates = get_six_month_averages(list(all_names), latest_date)

    for i, district in enumerate(districts):
        attended_list = attended.get(district, [])
        not_attended_list = not_attended.get(district, [])
        combined_list = []
        if previous_week_data:
            prev_attended = previous_week_data['attended'].get(district, [])
            prev_not_attended = previous_week_data['not_attended'].get(district, [])
            combined_list.extend((name, True, name in prev_not_attended) for name in attended_list)
            combined_list.extend((name, False, name in prev_attended) for name in not_attended_list)
        else:
            combined_list.extend((name, True, False) for name in attended_list)
            combined_list.extend((name, False, False) for name in not_attended_list)

        latest_dates = get_all_latest_attendance_dates([name for name, _, _ in combined_list], latest_date)
        combined_list.sort(key=lambda x: (-int(x[2]), latest_dates.get(x[0], datetime(1970, 1, 1))), reverse=True)

        for r, (name, is_attended, has_highlight) in enumerate(combined_list[:max_len]):
            col_offset = i * 3 + 1
            if is_attended:
                cell = new_sheet.cell(r + 3, col_offset)
                cell.value = name
                if has_highlight:
                    cell.fill = green_fill
            else:
                cell = new_sheet.cell(r + 3, col_offset + 1)
                cell.value = name
                if has_highlight:
                    cell.fill = red_fill
            cell_rate = new_sheet.cell(r + 3, col_offset + 2)
            cell_rate.value = f"{avg_rates.get(name, 0.0):.2%}"

def process_excel(file_stream, file_extension):
    logger.info(f"Processing Excel file, extension: {file_extension}")
    start_time = time.time()
    
    # Ensure indexes at startup
    ensure_indexes()

    file_stream.seek(0)
    file_content = file_stream.read()
    buffered_stream = BytesIO(file_content)

    if file_extension == '.xls':
        buffered_stream = convert_xls_to_xlsx(file_stream)
        file_extension = '.xlsx'

    try:
        workbook = openpyxl.load_workbook(buffered_stream)
    except Exception as e:
        logger.error(f"Failed to load workbook: {str(e)}")
        raise

    input_sheet = workbook.active
    week_cols = []
    current_month = "2025年1月"
    for col in range(START_COLUMN, input_sheet.max_column + 1):
        month_header = str(input_sheet.cell(1, col + 1).value or "").strip()
        week_header = str(input_sheet.cell(2, col + 1).value or "").strip()
        if "年" in month_header and "月" in month_header:
            current_month = month_header
        if "週" in week_header:
            week_cols.append((col, week_header, current_month))

    logger.info(f"Detected week columns: {week_cols}")
    if not week_cols:
        logger.warning("No week columns detected")
        return {
            'latest_analytic_date': None,
            'latest_attendance_data': None,
            'latest_week_display': None,
            'latest_district_counts': None,
            'latest_main_district': None,
            'latest_main_district_counts': None,
            'all_attendance_data': []
        }

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
    date_cache = {}  # Cache for latest attendance dates

    for col, week_name, month_prefix in week_cols:
        logger.info(f"Processing week: {week_name} in {month_prefix}")
        year = int(month_prefix.split("年")[0])
        month_num = int(month_prefix.split("年")[1].replace("月", ""))
        week_num = chinese_to_int(week_name.replace("第", "").replace("週", ""))
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
            logger.info(f"No attendees for {week_name} in {month_prefix}")
            continue

        all_attendance_data.append((current_date, {'attended': attended, 'not_attended': not_attended}, f"{month_prefix}{week_name}"))

        if current_date > latest_date:
            latest_date = current_date
            latest_attended = attended
            latest_not_attended = not_attended
            latest_week = f"{month_prefix}{week_name}"
            latest_districts = district_counts
            latest_main_district_counts = main_district_counts

    # Write records to database
    if all_records and all_names:
        bulk_ops = []
        existing_keys = set()
        write_start = time.time()
        
        # Check existing records efficiently
        min_date = min(r["date"] for r in all_records)
        max_date = max(r["date"] for r in all_records)
        try:
            cursor = db[COLLECTION_NAME].find(
                {
                    "date": {"$gte": min_date, "$lte": max_date},
                    "name": {"$in": list(all_names)}
                },
                {"name": 1, "date": 1, "_id": 0},
                hint="date_name_idx"
            )
            existing_keys = set((doc["name"], doc["date"]) for doc in cursor)
        except Exception as e:
            logger.error(f"Failed to query existing records: {str(e)}")
            existing_keys = set()  # Fallback to empty set

        # Only include new or changed records
        for r in all_records:
            key = (r["name"], r["date"])
            if key not in existing_keys:
                bulk_ops.append(UpdateOne(
                    {"name": r["name"], "date": r["date"]},
                    {"$set": r},
                    upsert=True
                ))

        if bulk_ops:
            try:
                for i in range(0, len(bulk_ops), BATCH_SIZE):
                    batch = bulk_ops[i:i + BATCH_SIZE]
                    db[COLLECTION_NAME].bulk_write(batch, ordered=False)
                    logger.info(f"Bulk wrote {len(batch)} records")
            except Exception as e:
                logger.error(f"Failed to bulk write records: {str(e)}")
        
        write_elapsed = time.time() - write_start
        logger.info(f"Database write completed in {write_elapsed:.2f}s")

    # Supplement missing records for the latest week only
    if latest_date > datetime(1970, 1, 1) and all_names:
        supplement_start = time.time()
        latest_week_str = latest_date.strftime("%Y-%m-%d")
        existing_keys = set()
        try:
            cursor = db[COLLECTION_NAME].find(
                {"date": latest_week_str, "name": {"$in": list(all_names)}},
                {"name": 1, "date": 1, "_id": 0},
                hint="date_name_idx"
            )
            existing_keys = set((doc["name"], doc["date"]) for doc in cursor)
        except Exception as e:
            logger.error(f"Failed to query latest week records: {str(e)}")

        bulk_ops = []
        for name in all_names:
            if (name, latest_week_str) not in existing_keys:
                district = next((d for d in attended if name in attended.get(d, [])), None) or \
                          next((d for d in not_attended if name in not_attended.get(d, [])), None) or "未知區"
                bulk_ops.append(UpdateOne(
                    {"name": name, "date": latest_week_str},
                    {"$set": {
                        "name": name,
                        "date": latest_week_str,
                        "attended": 0,
                        "district": district,
                        "age_group": "未知"
                    }},
                    upsert=True
                ))

        if bulk_ops:
            try:
                for i in range(0, len(bulk_ops), BATCH_SIZE):
                    batch = bulk_ops[i:i + BATCH_SIZE]
                    db[COLLECTION_NAME].bulk_write(batch, ordered=False)
                    logger.info(f"Bulk wrote {len(batch)} missing records for {latest_week_str}")
            except Exception as e:
                logger.error(f"Failed to bulk write missing records: {str(e)}")

        supplement_elapsed = time.time() - supplement_start
        logger.info(f"Missing records supplement completed in {supplement_elapsed:.2f}s")

    if not all_attendance_data:
        logger.warning("No valid attendance data processed")
        return {
            'latest_analytic_date': None,
            'latest_attendance_data': None,
            'latest_week_display': None,
            'latest_district_counts': None,
            'latest_main_district': None,
            'latest_main_district_counts': None,
            'all_attendance_data': []
        }

    total_elapsed = time.time() - start_time
    logger.info(f"Excel processing completed in {total_elapsed:.2f}s")
    return {
        'latest_analytic_date': latest_date.strftime("%Y年%m月%d日") if latest_date > datetime(1970, 1, 1) else None,
        'latest_attendance_data': {'attended': latest_attended, 'not_attended': latest_not_attended} if latest_attended else None,
        'latest_week_display': latest_week,
        'latest_district_counts': latest_districts,
        'latest_main_district': latest_main_district,
        'latest_main_district_counts': latest_main_district_counts,
        'all_attendance_data': all_attendance_data,
        'date_cache': date_cache  # Pass cache for rendering
    }

def generate_excel(all_attendance_data):
    logger.info("Generating Excel file")
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    for date, data, week_name in all_attendance_data:
        year = date.year
        month = date.month
        new_sheet_name = f"{year}年{month}月{week_name.split('月')[1]} 主日"

        if new_sheet_name in workbook.sheetnames:
            logger.error(f"Duplicate sheet name: {new_sheet_name}")
            raise ValueError(f"Sheet name '{new_sheet_name}' exists")

        new_sheet = workbook.create_sheet(new_sheet_name)
        previous_week_data = None
        current_week_idx = next(idx for idx, (d, _, _) in enumerate(all_attendance_data) if d == date)
        if current_week_idx > 0:
            previous_week_data = all_attendance_data[current_week_idx - 1][1]

        write_summary(new_sheet, data['attended'], data['not_attended'], date, previous_week_data)

    output_stream = BytesIO()
    workbook.save(output_stream)
    output_stream.seek(0)
    return output_stream

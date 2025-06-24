from config import db, DB_OFFLINE, COLLECTION_NAME
import logging
from pymongo.errors import DuplicateKeyError
from pymongo import UpdateOne
from datetime import datetime, timedelta
from utils import parse_week_display

logger = logging.getLogger(__name__)

def init_database():
    # Check for existing indexes
    existing_indexes = db[COLLECTION_NAME].index_information()
    index_name = "name_week_event_idx"
    
    if index_name in existing_indexes:
        logger.info(f"Index {index_name} already exists, skipping creation")
    else:
        # Check for duplicate records before creating unique index
        try:
            duplicates = db[COLLECTION_NAME].aggregate([
                {"$group": {
                    "_id": {"name": "$name", "week_display": "$week_display", "event_name": "$event_name"},
                    "count": {"$sum": 1},
                    "ids": {"$push": "$_id"}
                }},
                {"$match": {"count": {"$gt": 1}}}
            ])

            duplicate_found = False
            for dup in duplicates:
                duplicate_found = True
                ids = dup["ids"][1:]  # Keep the first record, remove others
                logger.warning(f"Found duplicate records for {dup['_id']}, removing {len(ids)} duplicates")
                db[COLLECTION_NAME].delete_many({"_id": {"$in": ids}})

            if duplicate_found:
                logger.info("Duplicate records cleaned up")

            # Create unique index for name, week_display, and event_name
            keys = [("name", 1), ("week_display", 1), ("event_name", 1)]
            index = {"unique": True, "name": index_name}
            db[COLLECTION_NAME].create_index(keys, **index)
            logger.info(f"Created unique index {index_name}")
        except DuplicateKeyError as e:
            logger.error(f"Failed to create index due to duplicate key: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Failed to initialize database: {str(e)}")
            raise

    # Set default event_name for existing records without event_name
    try:
        result = db[COLLECTION_NAME].update_many(
            {"event_name": {"$exists": False}},
            {"$set": {"event_name": "未指定活動"}}
        )
        logger.info(f"Updated {result.modified_count} records with default event_name")
    except Exception as e:
        logger.error(f"Failed to set default event_name: {str(e)}")

def bulk_write(operations):
    # Perform bulk write operations (InsertOne, UpdateOne, DeleteMany)
    try:
        if not operations:
            logger.info("No operations to execute in bulk write")
            return None
        
        result = db[COLLECTION_NAME].bulk_write(operations, ordered=False)
        logger.info(f"Bulk write completed: {result.bulk_api_result}, "
                   f"inserted={result.inserted_count}, modified={result.modified_count}, "
                   f"deleted={result.deleted_count}, matched={result.matched_count}")
        return result
    except Exception as e:
        logger.error(f"Failed to perform bulk write: {str(e)}")
        raise

def find_existing(names, week_display, event_names=None):
    # Find existing records for given names, week_display, and optional event_names
    try:
        # Ensure names is a list of strings
        if isinstance(names, set):
            names = list(names)
        if not isinstance(names, list):
            raise ValueError(f"names must be a list, got {type(names)}")
        if not all(isinstance(name, str) for name in names):
            raise ValueError("All names must be strings")
        
        # Ensure week_display is a list
        if isinstance(week_display, str):
            week_display = [week_display]
        if not isinstance(week_display, list):
            raise ValueError(f"week_display must be a list, got {type(week_display)}")

        query = {
            "name": {"$in": names},
            "week_display": {"$in": week_display}
        }
        if event_names:
            query["event_name"] = {"$in": event_names}

        records = db[COLLECTION_NAME].find(query)
        return [(r["name"], r["week_display"], r["event_name"]) for r in records]
    except Exception as e:
        logger.error(f"Failed to find existing records: {str(e)}")
        raise

def get_week_attendance_count(week_display, district, event_name):
    # Get the total number of attended records for a specific week_display, district, and event_name
    try:
        logger.debug(f"Querying attendance count for week_display: {week_display}, district: {district}, event_name: {event_name}")
        result = db[COLLECTION_NAME].aggregate([
            {"$match": {
                "week_display": week_display,
                "district": district,
                "event_name": event_name,
                "attended": 1
            }},
            {"$group": {
                "_id": None,
                "total_attended": {"$sum": 1}
            }}
        ])
        result = list(result)
        total_attended = result[0]["total_attended"] if result else 0
        logger.debug(f"Attendance count for {week_display}, {district}, {event_name}: {total_attended}")
        return total_attended
    except Exception as e:
        logger.error(f"Failed to get attendance count for {week_display}, {district}, {event_name}: {str(e)}")
        return 0

def get_all_latest_attendance_dates(names=None, placeholder_date=None):
    """Return latest week_display for every name, or list of all week_displays.
       - 連線正常 → 查 Mongo
       - 離線模式 → 回預設空結果，不觸發 DB
    """
    # ---------- 離線早退 ----------
    if DB_OFFLINE:
        if names:
            # dict: {name: None}
            if isinstance(names, set):
                names = list(names)
            return {n: None for n in names}
        # names is None → 回空 list
        return []

    # ---------- 正常連線 ----------
    try:
        pipeline = []
        if names:
            if isinstance(names, set):
                names = list(names)
            if not isinstance(names, list):
                raise ValueError(f"names must be a list, got {type(names)}")
            pipeline.append({"$match": {"name": {"$in": names}}})
        
        if placeholder_date:
            # Filter records before or on placeholder_date (approximated by week_display)
            year = placeholder_date.year
            month = placeholder_date.month
            week_num = (placeholder_date.day - 1) // 7 + 1
            date_filter = f"{year}年{month:02d}月第{week_num}週"
            pipeline.append({"$match": {"week_display": {"$lte": date_filter}}})

        pipeline.extend([
            {"$group": {
                "_id": "$name",
                "latest_week_display": {"$max": "$week_display"}
            }},
            {"$project": {
                "name": "$_id",
                "week_display": "$latest_week_display",
                "_id": 0
            }}
        ])

        results = db[COLLECTION_NAME].aggregate(pipeline)
        if names:
            # Return dictionary for specified names
            latest_dates = {r["name"]: r["week_display"] for r in results}
            return {name: latest_dates.get(name, None) for name in names}
        else:
            # Return all latest week_display values, sorted by parsed week_display
            weeks = set(r["week_display"] for r in results)
            return sorted(weeks, key=parse_week_display)
    except Exception as e:
        logger.error(f"Failed to get latest attendance dates: {str(e)}")
        raise

def get_six_month_averages(names, week_display):
    # Calculate six-month attendance averages based on week_display
    try:
        # Ensure names is a list
        if isinstance(names, set):
            names = list(names)
        if not isinstance(names, list):
            raise ValueError(f"names must be a list, got {type(names)}")
        
        # Parse week_display to get year, month, and week
        parsed_week = parse_week_display(week_display)
        if parsed_week == (0, 0, 0):
            logger.warning(f"Invalid week_display: {week_display}")
            return {}
        
        year, month, week = parsed_week
        # Estimate the start of the week (approximate)
        week_date = datetime(year, month, 1) + timedelta(days=(week - 1) * 7)
        six_months_ago = week_date - timedelta(days=180)  # Roughly 26 weeks
        
        # Convert six_months_ago to week_display format
        start_year = six_months_ago.year
        start_month = six_months_ago.month
        start_week_num = (six_months_ago.day - 1) // 7 + 1
        week_start = f"{start_year}年{start_month:02d}月第{start_week_num}週"
        
        pipeline = [
            {"$match": {
                "name": {"$in": names},
                "week_display": {"$gte": week_start, "$lte": week_display},
                "event_name": "主日"  # Only consider 主日 events
            }},
            {"$group": {
                "_id": "$name",
                "attendance_rate": {"$avg": {"$cond": [{"$eq": ["$attended", 1]}, 1, 0]}}
            }}
        ]
        results = db[COLLECTION_NAME].aggregate(pipeline)
        avg_rates = {r["_id"]: r["attendance_rate"] for r in results}
        
        # Ensure all names have a rate (0 if no records)
        return {name: avg_rates.get(name, 0.0) for name in names}
    except Exception as e:
        logger.error(f"Failed to calculate six-month averages for week_display {week_display}: {str(e)}")
        return {}

def get_event_name(week_display):
    """
    Return the event name for the given week_display.
    Preference order: 主日 → 其它(第一筆) → 未指定活動
    """
    try:
        # ① 最優先：主日
        record = db[COLLECTION_NAME].find_one(
            {"week_display": week_display, "event_name": "主日"},
            {"event_name": 1}
        )
        if record:
            return "主日"

        # ② 次選：該週第一筆文件的 event_name
        record = db[COLLECTION_NAME].find_one(
            {"week_display": week_display},
            {"event_name": 1}
        )
        if record and record.get("event_name"):
            return record["event_name"]

        # ③ 都沒有 → 預設
        return "未指定活動"

    except Exception as e:
        logger.error(f"Failed to get event name for {week_display}: {str(e)}")
        return "未指定活動"


def get_event_totals(week_display, main_district):
    # Get total attendance counts for '禱告', '小排', '晨興' for a specific week_display and main_district
    try:
        logger.debug(f"Querying event totals for week_display: {week_display}, main_district: {main_district}")
        pipeline = [
            {"$match": {
                "week_display": week_display,
                "district": {"$regex": f"^{main_district}"},
                "attended": 1,
                "event_name": {"$in": ["禱告", "小排", "晨興"]}
            }},
            {"$group": {
                "_id": {
                    "event_name": "$event_name",
                    "district": "$district"
                },
                "count": {"$sum": 1}
            }},
            {"$group": {
                "_id": "$_id.event_name",
                "district_counts": {
                    "$push": {
                        "district": "$_id.district",
                        "count": "$count"
                    }
                },
                "total_count": {"$sum": "$count"}
            }}
        ]
        results = list(db[COLLECTION_NAME].aggregate(pipeline))
        logger.debug(f"Raw query results for {week_display}, {main_district}: {results}")
        
        event_totals = {
            "禱告": {"total": 0, "districts": {}},
            "小排": {"total": 0, "districts": {}},
            "晨興": {"total": 0, "districts": {}}
        }
        
        for result in results:
            event_name = result["_id"]
            if event_name in event_totals:
                event_totals[event_name]["total"] = result["total_count"]
                for dc in result["district_counts"]:
                    event_totals[event_name]["districts"][dc["district"]] = dc["count"]
        
        logger.info(f"Event totals for {week_display}, {main_district}: {event_totals}")
        return event_totals
    except Exception as e:
        logger.error(f"Failed to get event totals for {week_display}, {main_district}: {str(e)}")
        return {
            "禱告": {"total": 0, "districts": {}},
            "小排": {"total": 0, "districts": {}},
            "晨興": {"total": 0, "districts": {}}
        }

def get_six_month_trimmed_mean(main_district, end_date):
    # Calculate trimmed mean attendance for '主日' over six months, excluding top and bottom 10%
    return get_six_month_trimmed_mean_by_event(main_district, end_date, event_name="主日")

def get_six_month_trimmed_mean_by_event(main_district, end_date, event_name):
    # Calculate trimmed mean attendance per district for a specific event over six months
    try:
        logger.debug(f"Calculating trimmed mean for {main_district}, event: {event_name}, up to {end_date}")
        six_months_ago = end_date - timedelta(days=180)
        year = six_months_ago.year
        month = six_months_ago.month
        week_num = (six_months_ago.day - 1) // 7 + 1
        week_start = f"{year}年{month:02d}月第{week_num}週"
        
        # Get all districts under the main district
        districts = db[COLLECTION_NAME].distinct("district", {
            "district": {"$regex": f"^{main_district}"},
            "event_name": event_name
        })
        
        result = {
            "districts": {},
            "main_district": main_district,
            "counts": {},
            "event_name": event_name
        }
        
        for district in districts:
            # Get attendance counts per week
            pipeline = [
                {"$match": {
                    "district": district,
                    "week_display": {"$gte": week_start},
                    "event_name": event_name,
                    "attended": 1
                }},
                {"$group": {
                    "_id": "$week_display",
                    "count": {"$sum": 1}
                }}
            ]
            weekly_counts = list(db[COLLECTION_NAME].aggregate(pipeline))
            counts = [r["count"] for r in weekly_counts]
            
            if not counts:
                result["districts"][district] = 0
                continue
            
            # Sort counts and trim top/bottom 10%
            counts.sort()
            n = len(counts)
            trim = int(n * 0.1)
            trimmed_counts = counts[trim:n-trim] if n > 2 * trim else counts
            
            # Calculate mean of trimmed counts and round to integer
            mean = round(sum(trimmed_counts) / len(trimmed_counts)) if trimmed_counts else 0
            result["districts"][district] = mean
        
        # Aggregate for main district
        pipeline = [
            {"$match": {
                "district": {"$regex": f"^{main_district}"},
                "week_display": {"$gte": week_start},
                "event_name": event_name,
                "attended": 1
            }},
            {"$group": {
                "_id": "$week_display",
                "count": {"$sum": 1}
            }}
        ]
        weekly_counts = list(db[COLLECTION_NAME].aggregate(pipeline))
        counts = [r["count"] for r in weekly_counts]
        
        if counts:
            counts.sort()
            n = len(counts)
            trim = int(n * 0.1)
            trimmed_counts = counts[trim:n-trim] if n > 2 * trim else counts
            mean = round(sum(trimmed_counts) / len(trimmed_counts)) if trimmed_counts else 0
            result["counts"][main_district] = mean
        else:
            result["counts"][main_district] = 0
        
        logger.info(f"Trimmed mean for {main_district}, event: {event_name}: {result}")
        return result
    except Exception as e:
        logger.error(f"Failed to calculate trimmed mean for {main_district}, event: {event_name}: {str(e)}")
        return {"districts": {}, "main_district": main_district, "counts": {}, "event_name": event_name}

def get_six_month_trimmed_mean_by_age_group(main_district, end_date):
    """Calculate trimmed mean attendance per age group for the last six months."""
    try:
        age_categories = ['青職以上', '大專', '中學', '小學', '學齡前']

        six_months_ago = end_date - timedelta(days=180)
        year = six_months_ago.year
        month = six_months_ago.month
        week_num = (six_months_ago.day - 1) // 7 + 1
        week_start = f"{year}年{month:02d}月第{week_num}週"

        districts = db[COLLECTION_NAME].distinct(
            "district",
            {"district": {"$regex": f"^{main_district}"}, "event_name": "主日"}
        )

        def _trimmed_mean(values):
            if not values:
                return 0
            values.sort()
            n = len(values)
            trim = int(n * 0.1)
            trimmed = values[trim:n - trim] if n > 2 * trim else values
            return round(sum(trimmed) / len(trimmed)) if trimmed else 0

        # Initialize result structure
        result = {age: {} for age in age_categories}

        for district in districts:
            pipeline = [
                {"$match": {
                    "district": district,
                    "week_display": {"$gte": week_start},
                    "event_name": "主日",
                    "attended": 1
                }},
                {"$group": {
                    "_id": {"age_group": "$age_group", "week_display": "$week_display"},
                    "count": {"$sum": 1}
                }},
                {"$group": {
                    "_id": "$_id.age_group",
                    "weekly": {"$push": "$count"}
                }}
            ]
            age_results = list(db[COLLECTION_NAME].aggregate(pipeline))
            for doc in age_results:
                age = doc["_id"] or '青職以上'
                if age not in age_categories:
                    age = '青職以上'
                result[age][district] = _trimmed_mean(doc.get("weekly", []))

        pipeline = [
            {"$match": {
                "district": {"$regex": f"^{main_district}"},
                "week_display": {"$gte": week_start},
                "event_name": "主日",
                "attended": 1
            }},
            {"$group": {
                "_id": {"age_group": "$age_group", "week_display": "$week_display"},
                "count": {"$sum": 1}
            }},
            {"$group": {
                "_id": "$_id.age_group",
                "weekly": {"$push": "$count"}
            }}
        ]
        age_results = list(db[COLLECTION_NAME].aggregate(pipeline))
        for doc in age_results:
            age = doc["_id"] or '青職以上'
            if age not in age_categories:
                age = '青職以上'
            result[age][main_district] = _trimmed_mean(doc.get("weekly", []))

        logger.info(
            f"Trimmed mean by age for {main_district}: {result}"
        )
        return result
    except Exception as e:
        logger.error(
            f"Failed to calculate trimmed mean by age for {main_district}: {str(e)}"
        )
        return {age: {} for age in ['青職以上', '大專', '中學', '小學', '學齡前']}

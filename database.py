from config import db, COLLECTION_NAME
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

def bulk_write(records):
    # Perform bulk write operations to insert or update records
    try:
        operations = []
        for record in records:
            if not record.get("name") or not record.get("week_display") or not record.get("event_name"):
                logger.warning(f"Skipping invalid record: {record}")
                continue
            # Use upsert to update existing record or insert new one
            operations.append(UpdateOne(
                {
                    "name": record["name"],
                    "week_display": record["week_display"],
                    "event_name": record["event_name"]
                },
                {"$set": record},
                upsert=True
            ))
        
        if operations:
            result = db[COLLECTION_NAME].bulk_write(operations)
            logger.info(f"Bulk write completed: {result.bulk_api_result}")
            return result
        else:
            logger.info("No valid records to write")
            return None
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

def get_all_latest_attendance_dates(names=None, placeholder_date=None):
    # Get latest week_display for each name, optionally filtered by names and date
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

def get_six_month_averages(names, end_date):
    # Calculate six-month attendance averages
    try:
        # Ensure names is a list
        if isinstance(names, set):
            names = list(names)
        
        # Calculate six months ago
        six_months_ago = end_date - timedelta(days=180)
        
        # Convert end_date to week_display format (e.g., "2025年4月第一週")
        year = six_months_ago.year
        month = six_months_ago.month
        week_num = (six_months_ago.day - 1) // 7 + 1
        week_start = f"{year}年{month:02d}月第{week_num}週"
        
        pipeline = [
            {"$match": {
                "name": {"$in": names},
                "week_display": {"$gte": week_start}
            }},
            {"$group": {
                "_id": "$name",
                "attendance_rate": {"$avg": {"$cond": [{"$eq": ["$attended", 1]}, 1, 0]}}
            }}
        ]
        results = db[COLLECTION_NAME].aggregate(pipeline)
        return {r["_id"]: r["attendance_rate"] for r in results}
    except Exception as e:
        logger.error(f"Failed to calculate six-month averages: {str(e)}")
        return {}

def get_event_name(week_display):
    # Get the event name associated with a specific week_display
    try:
        logger.debug(f"Querying event name for week_display: {week_display}")
        record = db[COLLECTION_NAME].find_one(
            {"week_display": week_display},
            {"event_name": 1}
        )
        if not record:
            logger.warning(f"No record found for week_display: {week_display}")
            return "未指定活動"
        event_name = record.get("event_name", "未指定活動")
        logger.debug(f"Retrieved event name for {week_display}: {event_name}")
        return event_name
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

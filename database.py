from config import db, COLLECTION_NAME, DB_TYPE
from datetime import datetime, timedelta
import time
import logging

logger = logging.getLogger(__name__)

def init_database():
    """Initialize database (create collection and indexes)"""
    try:
        indexes = [
            {"key": {"name": 1, "date": 1}, "unique": True, "name": "name_date_idx"},
            {"key": {"date": 1, "name": 1}, "name": "date_name_idx"}
        ]
        for index in indexes:
            db[COLLECTION_NAME].create_index(**index)
        logger.info("Initialized MongoDB with indexes for attendance_records")
        return db
    except Exception as e:
        logger.error(f"Failed to initialize database: {str(e)}")
        raise

def get_six_month_averages(names, latest_date):
    """Calculate six-month attendance averages"""
    if not names:
        return {}

    start_time = time.time()
    result = {}
    
    six_months_ago = (latest_date - timedelta(days=180)).strftime("%Y-%m-%d")
    latest_date_str = latest_date.strftime("%Y-%m-%d")
    
    batch_size = 100
    for i in range(0, len(names), batch_size):
        batch_names = names[i:i + batch_size]
        try:
            pipeline = [
                {"$match": {
                    "name": {"$in": batch_names},
                    "date": {"$gte": six_months_ago, "$lte": latest_date_str}
                }},
                {"$group": {
                    "_id": "$name",
                    "attendance_rate": {"$avg": "$attended"}
                }},
                {"$project": {
                    "_id": 1,
                    "attendance_rate": {"$ifNull": ["$attendance_rate", 0.0]}
                }}
            ]
            cursor = db[COLLECTION_NAME].aggregate(pipeline)
            for doc in cursor:
                result[doc["_id"]] = doc["attendance_rate"]
        except Exception as e:
            logger.error(f"Failed to get averages for batch {i//batch_size + 1}: {str(e)}")

    for name in names:
        if name not in result:
            result[name] = 0.0

    elapsed = time.time() - start_time
    logger.info(f"Calculated averages for {len(names)} names in {elapsed:.2f}s")
    return result

def bulk_write(records):
    """Bulk write records to database"""
    if not records:
        return
    start_time = time.time()
    try:
        batch_size = 5000
        for i in range(0, len(records), batch_size):
            batch = records[i:i + batch_size]
            try:
                db[COLLECTION_NAME].insert_many(batch, ordered=False)
                logger.info(f"Inserted {len(batch)} MongoDB records")
            except Exception as e:
                logger.warning(f"Insert failed for batch {i//batch_size + 1}: {str(e)}")
                for r in batch:
                    db[COLLECTION_NAME].update_one(
                        {"name": r["name"], "date": r["date"]},
                        {"$set": r},
                        upsert=True
                    )
                logger.info(f"Upserted {len(batch)} MongoDB records")
    except Exception as e:
        logger.error(f"Failed to bulk write records: {str(e)}")
        raise
    elapsed = time.time() - start_time
    logger.info(f"Database write completed in {elapsed:.2f}s")

def find_existing(min_date, max_date, names):
    """Find existing (name, date) pairs"""
    start_time = time.time()
    existing_keys = set()
    try:
        cursor = db[COLLECTION_NAME].find(
            {
                "date": {"$gte": min_date, "$lte": max_date},
                "name": {"$in": list(names)}
            },
            {"name": 1, "date": 1, "_id": 0}
        )
        existing_keys = set((doc["name"], doc["date"]) for doc in cursor)
    except Exception as e:
        logger.error(f"Failed to find existing records: {str(e)}")
    elapsed = time.time() - start_time
    logger.debug(f"Found existing records in {elapsed:.2f}s")
    return existing_keys

def get_all_latest_attendance_dates(names, latest_date, cache=None):
    """Get latest attendance dates for names"""
    if not names:
        return {}
    if cache is None:
        cache = {}
    
    names_to_query = [name for name in names if name not in cache]
    if not names_to_query:
        logger.debug(f"Cache hit for all {len(names)} names")
        return {name: cache[name] for name in names}

    start_time = time.time()
    latest_dates = {name: datetime(1970, 1, 1) for name in names}
    
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
        results = db[COLLECTION_NAME].aggregate(pipeline)
        for doc in results:
            if doc["max_date"]:
                date = datetime.strptime(doc["max_date"], '%Y-%m-%d')
                latest_dates[doc["_id"]] = date
                cache[doc["_id"]] = date
    except Exception as e:
        logger.error(f"Failed to get latest attendance dates: {str(e)}")
    
    elapsed = time.time() - start_time
    logger.debug(f"Retrieved latest dates for {len(names_to_query)} names in {elapsed:.2f}s")
    return latest_dates

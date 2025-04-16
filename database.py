from config import db, COLLECTION_NAME, logger, MONGO_URI
from datetime import datetime, timedelta
from pymongo import MongoClient, IndexModel, ASCENDING
import time

def init_database():
    """Initialize MongoDB connection and ensure indexes"""
    try:
        # Use existing db from config
        collection = db[COLLECTION_NAME]
        # Ensure indexes
        indexes = [
            IndexModel([("name", ASCENDING), ("date", ASCENDING), ("attended", ASCENDING)], name="name_date_attended_idx"),
            IndexModel([("date", ASCENDING), ("name", ASCENDING)], name="date_name_idx")
        ]
        collection.create_indexes(indexes)
        logger.info("Initialized database and verified indexes for attendance_records")
        return db
    except Exception as e:
        logger.error(f"Failed to initialize database: {str(e)}")
        raise

def get_six_month_averages(names, latest_date):
    """Calculate six-month attendance averages in batches"""
    if not names:
        return {}

    start_time = time.time()
    batch_size = 100
    result = {}
    
    for i in range(0, len(names), batch_size):
        batch_names = names[i:i + batch_size]
        try:
            pipeline = [
                {"$match": {
                    "name": {"$in": batch_names},
                    "date": {"$gte": (latest_date - timedelta(days=180)).strftime("%Y-%m-%d"),
                             "$lte": latest_date.strftime("%Y-%m-%d")}
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
            cursor = db[COLLECTION_NAME].aggregate(pipeline, hint="name_date_attended_idx")
            for doc in cursor:
                result[doc["_id"]] = doc["attendance_rate"]
        except Exception as e:
            logger.error(f"Failed to get averages for batch {i//batch_size + 1}: {str(e)}")
    
    # Fill missing names with 0.0
    for name in names:
        if name not in result:
            result[name] = 0.0

    elapsed = time.time() - start_time
    logger.info(f"Calculated averages for {len(names)} names in {elapsed:.2f}s")
    return result

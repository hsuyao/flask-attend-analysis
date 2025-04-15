from config import logger, db, COLLECTION_NAME
from datetime import datetime, timedelta
from pymongo import ASCENDING

def init_database():
    """Initialize MongoDB collection with optimized indexes"""
    try:
        collections = db.list_collection_names()
        if COLLECTION_NAME not in collections:
            db.create_collection(COLLECTION_NAME)
            logger.info(f"Created collection {COLLECTION_NAME}")
        else:
            logger.info(f"Collection {COLLECTION_NAME} exists")

        # Create optimized compound index
        db[COLLECTION_NAME].create_index(
            [("date", ASCENDING), ("name", ASCENDING), ("attended", ASCENDING)],
            name="date_name_attended_idx"
        )
        logger.info(f"Indexes optimized for {COLLECTION_NAME}")
    except Exception as e:
        logger.error(f"Failed to initialize database: {str(e)}")
        raise

def get_six_month_averages(names, latest_date):
    """Calculate six-month attendance rates for multiple names"""
    try:
        six_months_ago = latest_date - timedelta(days=180)
        pipeline = [
            {
                "$match": {
                    "name": {"$in": list(names)},
                    "date": {
                        "$gte": six_months_ago.strftime("%Y-%m-%d"),
                        "$lte": latest_date.strftime("%Y-%m-%d")
                    }
                }
            },
            {
                "$group": {
                    "_id": "$name",
                    "total": {"$sum": 1},
                    "attended": {"$sum": {"$cond": [{"$eq": ["$attended", 1]}, 1, 0]}}
                }
            },
            {
                "$project": {
                    "rate": {"$divide": ["$attended", {"$max": ["$total", 1]}]}
                }
            }
        ]
        results = db[COLLECTION_NAME].aggregate(pipeline)
        return {doc["_id"]: doc["rate"] for doc in results}
    except Exception as e:
        logger.error(f"Failed to calculate six-month averages: {str(e)}")
        return {name: 0.0 for name in names}

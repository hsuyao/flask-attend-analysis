from config import db, COLLECTION_NAME, logger, DB_TYPE
from datetime import datetime, timedelta
from pymongo import IndexModel, ASCENDING, UpdateOne
import time
import sqlite3

def init_database():
    """Initialize database (create table/collection and indexes)"""
    try:
        if DB_TYPE == "mongodb":
            # MongoDB: Create indexes
            indexes = [
                IndexModel([("name", ASCENDING), ("date", ASCENDING), ("attended", ASCENDING)], name="name_date_attended_idx"),
                IndexModel([("date", ASCENDING), ("name", ASCENDING)], name="date_name_idx")
            ]
            db[COLLECTION_NAME].create_indexes(indexes)
            logger.info("Initialized MongoDB with indexes for attendance_records")
        elif DB_TYPE == "sqlite":
            # SQLite: Create table and indexes
            cursor = db.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS attendance_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    date TEXT NOT NULL,
                    attended INTEGER NOT NULL,
                    district TEXT,
                    age_group TEXT,
                    UNIQUE(name, date)
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_date_name ON attendance_records (date, name)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_name_date_attended ON attendance_records (name, date, attended)")
            db.commit()
            logger.info("Initialized SQLite table and indexes for attendance_records")
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
    
    if DB_TYPE == "mongodb":
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
                cursor = db[COLLECTION_NAME].aggregate(pipeline, hint="name_date_attended_idx")
                for doc in cursor:
                    result[doc["_id"]] = doc["attendance_rate"]
            except Exception as e:
                logger.error(f"Failed to get averages for batch {i//batch_size + 1}: {str(e)}")
    elif DB_TYPE == "sqlite":
        try:
            cursor = db.cursor()
            query = """
                SELECT name, AVG(attended) as attendance_rate
                FROM attendance_records
                WHERE name IN ({}) AND date >= ? AND date <= ?
                GROUP BY name
            """
            # Prepare placeholders for names
            placeholders = ",".join("?" for _ in names)
            cursor.execute(query.format(placeholders), names + [six_months_ago, latest_date_str])
            for row in cursor.fetchall():
                result[row["name"]] = float(row["attendance_rate"]) if row["attendance_rate"] is not None else 0.0
        except Exception as e:
            logger.error(f"Failed to get averages for SQLite: {str(e)}")

    # Fill missing names with 0.0
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
        if DB_TYPE == "mongodb":
            bulk_ops = [
                UpdateOne(
                    {"name": r["name"], "date": r["date"]},
                    {"$set": r},
                    upsert=True
                ) for r in records
            ]
            for i in range(0, len(bulk_ops), 500):
                batch = bulk_ops[i:i + 500]
                db[COLLECTION_NAME].bulk_write(batch, ordered=False)
                logger.info(f"Bulk wrote {len(batch)} MongoDB records")
        elif DB_TYPE == "sqlite":
            cursor = db.cursor()
            query = """
                INSERT OR REPLACE INTO attendance_records (name, date, attended, district, age_group)
                VALUES (?, ?, ?, ?, ?)
            """
            cursor.executemany(query, [
                (r["name"], r["date"], r["attended"], r.get("district", "未知區"), r.get("age_group", "未知"))
                for r in records
            ])
            db.commit()
            logger.info(f"Bulk wrote {len(records)} SQLite records")
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
        if DB_TYPE == "mongodb":
            cursor = db[COLLECTION_NAME].find(
                {
                    "date": {"$gte": min_date, "$lte": max_date},
                    "name": {"$in": list(names)}
                },
                {"name": 1, "date": 1, "_id": 0},
                hint="date_name_idx"
            )
            existing_keys = set((doc["name"], doc["date"]) for doc in cursor)
        elif DB_TYPE == "sqlite":
            cursor = db.cursor()
            query = """
                SELECT name, date
                FROM attendance_records
                WHERE date >= ? AND date <= ? AND name IN ({})
            """
            placeholders = ",".join("?" for _ in names)
            cursor.execute(query.format(placeholders), [min_date, max_date] + list(names))
            existing_keys = set((row["name"], row["date"]) for row in cursor.fetchall())
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
        if DB_TYPE == "mongodb":
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
            for doc in results:
                if doc["max_date"]:
                    date = datetime.strptime(doc["max_date"], '%Y-%m-%d')
                    latest_dates[doc["_id"]] = date
                    cache[doc["_id"]] = date
        elif DB_TYPE == "sqlite":
            cursor = db.cursor()
            query = """
                SELECT name, MAX(date) as max_date
                FROM attendance_records
                WHERE name IN ({}) AND attended = 1 AND date <= ?
                GROUP BY name
            """
            placeholders = ",".join("?" for _ in names_to_query)
            cursor.execute(query.format(placeholders), names_to_query + [latest_date.strftime('%Y-%m-%d')])
            for row in cursor.fetchall():
                if row["max_date"]:
                    date = datetime.strptime(row["max_date"], '%Y-%m-%d')
                    latest_dates[row["name"]] = date
                    cache[row["name"]] = date
    except Exception as e:
        logger.error(f"Failed to get latest attendance dates: {str(e)}")
    
    elapsed = time.time() - start_time
    logger.debug(f"Retrieved latest dates for {len(names_to_query)} names in {elapsed:.2f}s")
    return latest_dates

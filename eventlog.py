import logging
from datetime import datetime
from config import db, DB_OFFLINE

logger = logging.getLogger(__name__)

COLLECTION_NAME = "event_log"
ARCHIVE_COLLECTION_NAME = "event_log_archive"
MAX_ENTRIES = 100000


def init_eventlog_collection():
    """Ensure event_log collections exist and have indexes on ``ts``."""
    if DB_OFFLINE:
        logger.warning("離線模式：跳過 event_log 初始化")
        return

    if COLLECTION_NAME not in db.list_collection_names():
        try:
            db.create_collection(COLLECTION_NAME)
            logger.info("Created event_log collection")
        except Exception as e:
            logger.error(f"Failed to create event_log collection: {e}")

    try:
        db[COLLECTION_NAME].create_index("ts")
    except Exception as e:
        logger.error(f"Failed to create index on event_log.ts: {e}")

    if ARCHIVE_COLLECTION_NAME not in db.list_collection_names():
        try:
            db.create_collection(ARCHIVE_COLLECTION_NAME)
            logger.info("Created event_log_archive collection")
        except Exception as e:
            logger.error(f"Failed to create event_log_archive collection: {e}")

    try:
        db[ARCHIVE_COLLECTION_NAME].create_index("ts")
    except Exception as e:
        logger.error(f"Failed to create index on event_log_archive.ts: {e}")


def _archive_old_entries(docs):
    """Move old log entries to the archive collection."""
    try:
        if not docs:
            return
        for d in docs:
            d.pop("_id", None)
        db[ARCHIVE_COLLECTION_NAME].insert_many(docs)
    except Exception as e:
        logger.error(f"Failed to archive event_log entries: {e}")


def _enforce_limit():
    """Ensure only MAX_ENTRIES newest documents are kept."""
    try:
        cnt = db[COLLECTION_NAME].estimated_document_count()
        if cnt > MAX_ENTRIES:
            skip = cnt - MAX_ENTRIES
            cursor = db[COLLECTION_NAME].find().sort("ts", 1).limit(skip)
            old_docs = list(cursor)
            if old_docs:
                _archive_old_entries(old_docs)
                ids = [d["_id"] for d in old_docs if "_id" in d]
                db[COLLECTION_NAME].delete_many({"_id": {"$in": ids}})
    except Exception as e:
        logger.error(f"Failed to enforce event_log limit: {e}")


def log_event(action, username=None, details=None):
    """Write an event log entry."""
    if DB_OFFLINE:
        return
    doc = {
        "action": action,
        "username": username,
        "details": details,
        "ts": datetime.utcnow(),
    }
    try:
        db[COLLECTION_NAME].insert_one(doc)
        _enforce_limit()
    except Exception as e:
        logger.error(f"Failed to log event: {e}")


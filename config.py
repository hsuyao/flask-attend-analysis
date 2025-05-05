import os
import logging
from pymongo import MongoClient

# Configure logging
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# Environment variables
DB_TYPE = os.getenv("DB_TYPE", "mongodb")  # Default to mongodb
MONGO_URI = os.getenv("MONGO_URI")
COLLECTION_NAME = "attendance_records"
START_COLUMN = 7

# Initialize database connection
if DB_TYPE == "mongodb":
    if not MONGO_URI:
        logger.error("MONGO_URI is required for MongoDB")
        raise ValueError("MONGO_URI is not set")
    try:
        client = MongoClient(MONGO_URI)
        db = client["attendance_db"]
        logger.info("Successfully connected to MongoDB Atlas")
    except Exception as e:
        logger.error(f"Failed to connect to MongoDB: {str(e)}")
        raise
else:
    logger.error(f"Invalid DB_TYPE: {DB_TYPE}")
    raise ValueError(f"Unsupported DB_TYPE: {DB_TYPE}")
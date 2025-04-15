import os
import logging
from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Constants
START_COLUMN = 8
BATCH_SIZE = 500

# MongoDB Atlas Configuration
MONGO_URI_DEFAULT = "mongodb+srv://attendance-analysis:aaaaaaaaa@cluster0.1c6d4dt.mongodb.net/attendance-analysis?retryWrites=true&w=majority&appName=Cluster0"
MONGO_URI = os.getenv("MONGO_URI", MONGO_URI_DEFAULT)
DATABASE_NAME = "attendance-analysis"
COLLECTION_NAME = "attendance_records"

# Validate MongoDB URI
if not MONGO_URI.startswith(("mongodb://", "mongodb+srv://")):
    logger.error(f"Invalid MONGO_URI: {MONGO_URI}. Must start with 'mongodb://' or 'mongodb+srv://'")
    raise ValueError("Invalid MongoDB URI scheme")

# Initialize MongoDB client
try:
    mongo_client = MongoClient(MONGO_URI, server_api=ServerApi('1'))
    mongo_client.admin.command('ping')
    logger.info("Successfully connected to MongoDB Atlas")
except Exception as e:
    logger.error(f"Failed to connect to MongoDB Atlas: {str(e)}")
    raise

db = mongo_client[DATABASE_NAME]
collection = db[COLLECTION_NAME]

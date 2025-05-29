import os
import logging
from pymongo import MongoClient, errors

# ──────────────────────────────────────────────────────────────
#  logging & env
# ──────────────────────────────────────────────────────────────
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

DB_TYPE = os.getenv("DB_TYPE", "mongodb")
MONGO_URI = os.getenv("MONGO_URI")
COLLECTION_NAME = "attendance_records"
START_COLUMN = 7

# ──────────────────────────────────────────────────────────────
#  Offline stub
# ──────────────────────────────────────────────────────────────
class _OfflineCollection:
    def __init__(self, name): self.name = name
    def __getattr__(self, _):
        def _stub(*a, **kw):
            raise RuntimeError("資料庫離線，無法存取")
        return _stub

class _OfflineDB:
    def __getitem__(self, name): return _OfflineCollection(name)
    def command(self, *a, **kw): raise RuntimeError("資料庫離線")

# ──────────────────────────────────────────────────────────────
#  Connect helper
# ──────────────────────────────────────────────────────────────
def _connect_mongo(uri: str):
    if not uri:
        raise ValueError("MONGO_URI 未設定")
    try:
        client = MongoClient(uri, serverSelectionTimeoutMS=3000, connect=True)
        client.admin.command("ping")          # 強制 I/O
        logger.info("✅  MongoDB connected")
        return client["attendance_db"], False  # False = 非離線
    except (errors.PyMongoError, Exception) as e:
        logger.warning(f"⚠️  MongoDB unreachable: {e}")
        return _OfflineDB(), True             # True  = 離線

# ──────────────────────────────────────────────────────────────
#  Export objects
# ──────────────────────────────────────────────────────────────
if DB_TYPE != "mongodb":
    raise ValueError(f"Unsupported DB_TYPE: {DB_TYPE}")

db, DB_OFFLINE = _connect_mongo(MONGO_URI)
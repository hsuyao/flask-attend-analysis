# admin_routes.py  –  Blueprint for DB admin UI
from flask import Blueprint, render_template, request, send_file, jsonify, \
                  current_app, session
from io import StringIO, BytesIO
from datetime import datetime, timedelta
from bson import ObjectId
from pymongo import UpdateOne
import csv, re, os, math

from config import db, DB_OFFLINE, COLLECTION_NAME
from user import update_user_role, block_user, unblock_user, delete_user
from utils import parse_week_display   # 若後續用不到可移除
from eventlog import log_event

# --- helper: avoid circular import ---
def is_admin():
    return 'user' in session and session['user'].get('role') == 'admin'
# -------------------------------------

admin_bp = Blueprint("admin_bp", __name__, url_prefix="/admin")

# ------------------------------------------------------------------------------
# Admin home page
# ------------------------------------------------------------------------------
@admin_bp.route("/")
def admin_home():
    if DB_OFFLINE:
        return "Database offline", 503
    if not is_admin():
        return "Forbidden", 403
    # distinct values for dropdowns
    districts = sorted(db[COLLECTION_NAME].distinct("district"))
    events    = sorted(db[COLLECTION_NAME].distinct("event_name"))
    return render_template("admin.html",
                           version=current_app.config.get("VERSION", "dev"),
                           districts=districts,
                           events=events)

# ------------------------------------------------------------------------------
#  使用者管理主頁面
# ------------------------------------------------------------------------------
@admin_bp.route("/users")
def admin_users():
    if DB_OFFLINE:
        return "Database offline", 503
    if not is_admin():
        return "Forbidden", 403
    return render_template("admin_users.html", version=current_app.config.get("VERSION", "dev"))

# ------------------------------------------------------------------------------
#  使用者列表 (AJAX)
# ------------------------------------------------------------------------------
@admin_bp.route("/users/data")
def admin_users_data():
    if DB_OFFLINE:
        return jsonify({"error": "db offline"}), 503
    if not is_admin():
        return jsonify({"error": "forbidden"}), 403

    users = list(db["users"].find({}, {"password":0}))
    # serialize
    data = []
    for u in users:
        data.append({
            "username": u["username"],
            "email": u.get("email",""),
            "role": u.get("role",""),
            "blocked": u.get("blocked", False)
        })
    return jsonify(data)

# ------------------------------------------------------------------------------
#  更新使用者角色
# ------------------------------------------------------------------------------
@admin_bp.route("/users/update_role", methods=["POST"])
def admin_update_role():
    if not is_admin():
        return jsonify({"error":"forbidden"}), 403
    username = request.json.get("username")
    new_role = request.json.get("role")
    ok = update_user_role(username, new_role)
    if ok:
        log_event("admin_update_role", session.get('user', {}).get('username'), details={"target": username, "role": new_role})
    return jsonify({"success": ok})

# ------------------------------------------------------------------------------
#  封鎖 / 解封
# ------------------------------------------------------------------------------
@admin_bp.route("/users/block", methods=["POST"])
def admin_block():
    if not is_admin():
        return jsonify({"error":"forbidden"}), 403
    username = request.json.get("username")
    ok = block_user(username)
    if ok:
        log_event("admin_block", session.get('user', {}).get('username'), details={"target": username})
    return jsonify({"success": ok})

@admin_bp.route("/users/unblock", methods=["POST"])
def admin_unblock():
    if not is_admin():
        return jsonify({"error":"forbidden"}), 403
    username = request.json.get("username")
    ok = unblock_user(username)
    if ok:
        log_event("admin_unblock", session.get('user', {}).get('username'), details={"target": username})
    return jsonify({"success": ok})

# ------------------------------------------------------------------------------
#  刪除使用者
# ------------------------------------------------------------------------------
@admin_bp.route("/users/delete", methods=["POST"])
def admin_delete_user():
    if not is_admin():
        return jsonify({"error":"forbidden"}), 403
    username = request.json.get("username")
    ok = delete_user(username)
    if ok:
        log_event("admin_delete_user", session.get('user', {}).get('username'), details={"target": username})
    return jsonify({"success": ok})

# ------------------------------------------------------------------------------
# REST - data list  (pagination + multi-filter)
# ------------------------------------------------------------------------------
@admin_bp.route("/data")
def admin_data():
    if not is_admin():
        return jsonify({"error": "forbidden"}), 403

    # pagination / sort
    page      = int(request.args.get("page", 1))
    per_page  = int(request.args.get("per_page", 20))
    sort_by   = request.args.get("sort", "week_display")
    order     = int(request.args.get("order", -1))        # -1 desc, 1 asc

    # filters
    name_kw   = request.args.get("name", "").strip()
    start_dt  = request.args.get("start", "").strip()     # YYYY-MM-DD
    end_dt    = request.args.get("end", "").strip()
    district  = request.args.get("district", "").strip()
    event     = request.args.get("event", "").strip()

    query = {}
    if name_kw:
        query["name"] = {"$regex": re.escape(name_kw)}
    if start_dt or end_dt:
        query["date"] = {}
        if start_dt:
            query["date"]["$gte"] = start_dt
        if end_dt:
            query["date"]["$lte"] = end_dt
    if district:
        query["district"] = district
    if event:
        query["event_name"] = event

    cursor = db[COLLECTION_NAME].find(query).sort(sort_by, order)
    total  = db[COLLECTION_NAME].count_documents(query)
    cursor = cursor.skip((page - 1) * per_page).limit(per_page)

    records = [{
        "id": str(doc["_id"]),
        "name": doc["name"],
        "district": doc["district"],
        "week_display": doc["week_display"],
        "event_name": doc.get("event_name", ""),
        "attended": doc["attended"],
        "date": doc.get("date", "")
    } for doc in cursor]

    return jsonify({"total": total, "records": records})

# ------------------------------------------------------------------------------
# Delete single record
# ------------------------------------------------------------------------------
@admin_bp.route("/delete/<doc_id>", methods=["DELETE"])
def admin_delete(doc_id):
    if not is_admin():
        return jsonify({"error": "forbidden"}), 403
    try:
        res = db[COLLECTION_NAME].delete_one({"_id": ObjectId(doc_id)})
        return jsonify({"deleted": res.deleted_count})
    except Exception as e:
        current_app.logger.error(f"Delete failed: {e}")
        return jsonify({"error": str(e)}), 500

@admin_bp.route("/delete_batch", methods=["POST"])
def admin_delete_batch():
    if not is_admin():
        return jsonify({"error": "forbidden"}), 403

    ids = request.json.get("ids", [])
    if not ids:
        return jsonify({"error": "no ids"}), 400

    try:
        obj_ids = [ObjectId(i) for i in ids]
        res = db[COLLECTION_NAME].delete_many({"_id": {"$in": obj_ids}})
        return jsonify({"deleted": res.deleted_count})
    except Exception as e:
        current_app.logger.error(f"Batch delete failed: {e}")
        return jsonify({"error": str(e)}), 500

# ------------------------------------------------------------------------------
# Export CSV
# ------------------------------------------------------------------------------
@admin_bp.route("/export")
def admin_export():
    if not is_admin():
        return "Forbidden", 403

    sio = StringIO()
    writer = csv.writer(sio)
    writer.writerow(["name","district","week_display","event_name",
                     "attended","age_group","date"])
    for doc in db[COLLECTION_NAME].find():
        writer.writerow([doc.get(k,"") for k in
                         ("name","district","week_display","event_name",
                          "attended","age_group","date")])

    mem = BytesIO(sio.getvalue().encode("utf-8-sig"))
    mem.seek(0)
    fname = f"attendance_{datetime.now():%Y%m%d_%H%M%S}.csv"
    return send_file(mem, as_attachment=True,
                     download_name=fname, mimetype="text/csv")

# ------------------------------------------------------------------------------
# Import CSV (upsert)
# ------------------------------------------------------------------------------
@admin_bp.route("/import", methods=["POST"])
def admin_import():
    if not is_admin():
        return jsonify({"error": "forbidden"}), 403
    f = request.files.get("file")
    if not f or not f.filename.endswith(".csv"):
        return jsonify({"error": "Please upload .csv"}), 400

    reader = csv.DictReader(StringIO(f.stream.read().decode("utf-8-sig")))
    ops, count = [], 0
    for row in reader:
        row["attended"] = int(row.get("attended", 0))
        filter_ = {"name": row["name"],
                   "week_display": row["week_display"],
                   "event_name": row.get("event_name","")}
        ops.append(UpdateOne(filter_, {"$set": row}, upsert=True))
        count += 1
        if len(ops) == 1000:
            db[COLLECTION_NAME].bulk_write(ops, ordered=False); ops=[]
    if ops:
        db[COLLECTION_NAME].bulk_write(ops, ordered=False)
    return jsonify({"imported": count})

@admin_bp.route("/db_status")
def admin_db_status():
    if not is_admin():
        return jsonify({"error": "forbidden"}), 403

    # MongoDB dbStats
    stats = db.command("dbStats")
    # 使用 dataSize + indexSize 以符合 MongoDB 網站的用量顯示
    used_bytes = stats.get("dataSize", 0) + stats.get("indexSize", 0)
    used_mb    = used_bytes / (1024 * 1024)
    quota_mb   = float(os.getenv("DB_QUOTA_MB", 512))          # 預設 512 MB，可用環境變數覆蓋
    remain_mb  = max(0.0, quota_mb - used_mb)

    return jsonify({
        "used_mb":   round(used_mb, 1),
        "quota_mb":   quota_mb,
        "remain_mb":  round(remain_mb, 1),
        "collections": stats["collections"],
        "objects":     stats["objects"]
    })

# ------------------------------------------------------------------------------
#  Event log viewer
# ------------------------------------------------------------------------------
@admin_bp.route("/logs")
def admin_logs():
    if DB_OFFLINE:
        return "Database offline", 503
    if not is_admin():
        return "Forbidden", 403
    log_event("view_event_logs", session.get('user', {}).get('username'))
    return render_template("admin_logs.html", version=current_app.config.get("VERSION", "dev"))


@admin_bp.route("/logs/data")
def admin_logs_data():
    if DB_OFFLINE:
        return jsonify({"error": "db offline"}), 503
    if not is_admin():
        return jsonify({"error": "forbidden"}), 403

    page     = int(request.args.get("page", 1))
    per_page = int(request.args.get("per_page", 20))
    sort_by  = request.args.get("sort", "ts")
    order    = int(request.args.get("order", -1))

    cursor = db["event_log"].find().sort(sort_by, order)
    total  = db["event_log"].count_documents({})
    cursor = cursor.skip((page - 1) * per_page).limit(per_page)

    records = []
    for doc in cursor:
        ts = doc.get("ts")
        if isinstance(ts, datetime):
            ts = ts.strftime("%Y-%m-%d %H:%M:%S")
        records.append({
            "ts": ts,
            "action": doc.get("action", ""),
            "username": doc.get("username", ""),
            "details": doc.get("details")
        })

    return jsonify({"total": total, "records": records})

# admin_routes.py  –  Blueprint for DB admin UI
from flask import Blueprint, render_template, request, send_file, jsonify, \
                  current_app, session
from io import StringIO, BytesIO
from datetime import datetime, timedelta
from bson import ObjectId
from pymongo import UpdateOne
import csv, re

from config import db, COLLECTION_NAME
from utils import parse_week_display   # 若後續用不到可移除

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


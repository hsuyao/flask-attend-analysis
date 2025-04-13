import sqlite3
from datetime import datetime, timedelta
from config import logger

DATABASE_PATH = '/app/attendance.db'

def init_database():
    """初始化数据库和表，并添加唯一索引"""
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS attendance_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            date TEXT NOT NULL,
            attended INTEGER NOT NULL,
            district TEXT NOT NULL,
            age_group TEXT NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE UNIQUE INDEX IF NOT EXISTS idx_name_date 
        ON attendance_records (name, date)
    ''')
    conn.commit()
    conn.close()
    logger.info("Database initialized with unique index")

def add_attendance_record(name, date, attended, district, age_group):
    """添加或更新出勤记录"""
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT OR REPLACE INTO attendance_records (name, date, attended, district, age_group)
        VALUES (?, ?, ?, ?, ?)
    ''', (name, date.strftime('%Y-%m-%d'), attended, district, age_group))
    conn.commit()
    conn.close()

def get_six_month_average(name, latest_date):
    """计算某人过去 6 个月的移动平均出勤率"""
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    six_months_ago = latest_date - timedelta(days=180)
    cursor.execute('''
        SELECT attended FROM attendance_records
        WHERE name = ? AND date >= ? AND date <= ?
        ORDER BY date ASC
    ''', (name, six_months_ago.strftime('%Y-%m-%d'), latest_date.strftime('%Y-%m-%d')))
    records = cursor.fetchall()
    conn.close()
    
    total_weeks = (latest_date - six_months_ago).days // 7 + 1
    if not records:
        return 0.0
    
    total_attended = sum(record[0] for record in records)
    return total_attended / total_weeks

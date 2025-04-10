# utils.py
from config import logger

def chinese_to_int(chinese_num):
    """Convert Chinese numerals to Arabic integers."""
    numeral_map = {
        '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
    }
    return numeral_map.get(chinese_num, 0)

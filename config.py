# config.py
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Constants
START_COLUMN = 8

class State:
    def __init__(self):
        self.latest_analytic_date = None
        self.latest_attendance_data = None  # {'attended': {}, 'not_attended': {}}
        self.latest_file_stream = None
        self.latest_week_display = None
        self.latest_district_counts = None  # {'district': {'total': count, 'ages': {'age': count}}, '總計': total}
        self.latest_main_district = None  # Main district name
        self.all_attendance_data = []  # List of (date, {'attended': {}, 'not_attended': {}}, week_display) for all weeks

# Create a global state instance
state = State()

from pywinauto import application
from win32gui import GetWindowText, GetForegroundWindow
from threading import Thread
import sys
import subprocess
from datetime import timedelta, datetime
from time import sleep, strptime
from json import load, dump
import os

DEFAULT_RBS_TIMEOUT_SECONDS = 1200


class CalculateRBSTime:
    def __init__(self):
        self.json_data = {}
        self.total_seconds_elapsed = 0
        self.json_path = 'RBS_Attendance.json'
        self.rbs_exe_name = 'CDViewer.exe'
        self.start_date = None

    def create_json(self):
        if not os.path.exists(self.json_path):
            with open(self.json_path, 'w') as f:
                dump(self.json_data, f, indent=4)
        else:
            with open(self.json_path, 'r') as f:
                self.json_data = load(f)

    @staticmethod
    def get_current_date():
        return datetime.today().strftime('%d-%m-%Y')

    def check_previous_time(self):
        self.start_date = self.get_current_date()
        if self.start_date in self.json_data:
            time_till_now = self.json_data[self.start_date]
            if time_till_now:
                timestamp = strptime(time_till_now, '%H:%M:%S')
                self.total_seconds_elapsed = self.timestamp_to_seconds(timestamp)
            else:
                time_till_now = "None"
            print('Previous Time for Date {} - {}\n'.format(self.start_date, time_till_now))
        else:
            print('No Previous Records for Date - {} found\n'.format(self.start_date))

    @staticmethod
    def timestamp_to_seconds(timestamp):
        return timedelta(hours=timestamp.tm_hour, minutes=timestamp.tm_min, seconds=timestamp.tm_sec).total_seconds()

    @staticmethod
    def seconds_to_timestamp(seconds):
        return timedelta(seconds=seconds)

    def write_to_json(self):
        current_date = self.get_current_date()
        if current_date != self.start_date:
            self.total_seconds_elapsed = 0
            self.start_date = current_date
        self.json_data[current_date] = str(self.seconds_to_timestamp(self.total_seconds_elapsed))
        with open(self.json_path, 'w+') as f:
            dump(self.json_data, f, indent=4)

    def check_if_rbs_in_memory(self):
        time_elapsed = 0
        while True:
            try:
                _ = subprocess.check_output('tasklist | findstr {}'.format(self.rbs_exe_name), shell=True).decode().strip()
                self.total_seconds_elapsed += 1
            except subprocess.CalledProcessError:
                pass

            sleep(1)
            time_elapsed += 1
            if time_elapsed % 5 == 0:
                print('\n-> Total time since RBS in memory : {}\n'.format(self.seconds_to_timestamp(self.total_seconds_elapsed)))
            if self.total_seconds_elapsed and self.total_seconds_elapsed % 5 == 0:
                self.write_to_json()

    def main(self):
        self.create_json()
        self.check_previous_time()
        self.check_if_rbs_in_memory()


class RBS:
    def __init__(self):
        self.current_window = None
        self.rbs_app = 'Desktop Viewer'

        self.current_window = None
        self.rbs_in_focus = False
        self.start_rbs_countdown = False
        self.total_seconds_since_last_focus = 0
        self.current_rbs_timeout = DEFAULT_RBS_TIMEOUT_SECONDS
        self.is_rbs_opened = False
        self.calculate_rbs_time = CalculateRBSTime()

    def check_current_window(self):
        while True:
            try:
                self.current_window = GetWindowText(GetForegroundWindow())
                if self.rbs_app in self.current_window:  # Whenever RBS comes into focus
                    print('RBS is in focus...')
                    self.total_seconds_since_last_focus = 0
                    self.start_rbs_countdown = False
                    self.current_rbs_timeout = DEFAULT_RBS_TIMEOUT_SECONDS
                else:  # Start RBS countdown in case not in focus
                    print('RBS not in focus...')
                    self.start_rbs_countdown = True
                sleep(1)
            except KeyboardInterrupt:
                break

    def try_to_open_rbs(self):
        self.open_rbs()
        rbs_status = self.check_if_rbs_opened()
        if rbs_status:
            print('RBS already opened...')
            self.start_rbs_countdown = False
            self.current_rbs_timeout = DEFAULT_RBS_TIMEOUT_SECONDS
        else:
            self.current_rbs_timeout = 5

    def open_rbs_after_limit(self):
        while self.start_rbs_countdown:
            self.total_seconds_since_last_focus += 1
            print('Time since RBS Last Focus : {} Seconds'.format(self.total_seconds_since_last_focus))
            if self.total_seconds_since_last_focus == self.current_rbs_timeout:
                self.try_to_open_rbs()
                self.total_seconds_since_last_focus = 0
            sleep(1)

    def watch_rbs(self):
        while True:
            if self.start_rbs_countdown:
                self.open_rbs_after_limit()
            sleep(1)

    def check_if_rbs_opened(self):
        sleep(1)
        self.rbs_in_focus = False
        current_window = GetWindowText(GetForegroundWindow())
        if self.rbs_app in current_window:
            self.rbs_in_focus = True
        return self.rbs_in_focus

    def open_rbs(self):
        app = application.Application()
        try:
            print('Opening : RBS')
            app.connect(title_re=".*{}.*".format(self.rbs_app))

            app_dialog = app.top_window()

            app_dialog.minimize()
            app_dialog.restore()
        except Exception as e:
            print('Exception : {}'.format(e))
            print('RBS Could not be opened...!!!')

    def start_threads(self):
        Thread(target=self.calculate_rbs_time.main, daemon=True).start()
        Thread(target=self.check_current_window, daemon=True).start()
        self.watch_rbs()

    def parse_args(self):
        global DEFAULT_RBS_TIMEOUT_SECONDS
        args = sys.argv
        if len(args) == 1:
            pass
        elif len(args) == 2:
            try:
                DEFAULT_RBS_TIMEOUT_SECONDS = int(args[1])
            except Exception as e:
                print('Exception : {}'.format(e))
                exit(1)
        else:
            print('Improper Command line arguments Supplied...!!!')
            exit(2)
        self.current_rbs_timeout = DEFAULT_RBS_TIMEOUT_SECONDS

    def main(self):
        self.parse_args()

        print('\nRAT - RBS Attendance Tracker')
        sleep(1)
        print('Timeout Set to : {} Seconds...'.format(self.current_rbs_timeout))
        self.start_threads()


if __name__ == '__main__':
    RBS().main()

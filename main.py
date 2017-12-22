import ctypes
import datetime
import time
import win32api
import win32gui
import json
import atexit
import threading

isAlive = True

class WorkLog:
    def __init__(self):
        self.logfile = open('worklog.json','r+');
        try:
            self.worklog = json.load(self.logfile)
        except json.JSONDecodeError:
            self.worklog = []

    def save(self):
        self.logfile.seek(0)
        self.logfile.write(json.dumps(self.worklog))
        self.logfile.truncate()

    def __append_new_work(self, work):
        current_time = get_current_time()
        self.worklog.append({'start' : current_time , 'end' : current_time, 'name' : work})

    def log(self, work):
        print(work)
        if len(self.worklog) > 0:
            lastWork = self.worklog[-1]
            if (lastWork['name'] == work):
                lastWork['end'] = get_current_time()
            else:
                self.__append_new_work(work)
        else:
             self.__append_new_work(work)


track_log = WorkLog()

def get_current_time():
    return str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))


def get_active_window():
    return win32gui.GetWindowText(win32gui.GetForegroundWindow())

def is_workstation_locked():
    user32 = ctypes.windll.User32
    OpenDesktop = user32.OpenDesktopW
    SwitchDesktop = user32.SwitchDesktop
    DESKTOP_SWITCHDESKTOP = 0x0100

    hDesktop = OpenDesktop (u"Default", 0, False, DESKTOP_SWITCHDESKTOP)
    unlocked_state = SwitchDesktop (hDesktop)
    return not unlocked_state

def parse_window_name(window_name):
    parameters = window_name.split(' - ')
    return str(parameters)

@atexit.register
def save_track_log():
    global track_log
    track_log.save()
    print("Saving log...")

def kill():
    print("Killing...")
    global isAlive
    isAlive = False

def save_track_thread():
    while True:
        save_track_log()
        time.sleep(1*60)

def track_time_thread():
    global track_log
    while True:
        if is_workstation_locked():
            track_log.log('WORKSTATION LOCKED')
        else:
            track_log.log(parse_window_name(get_active_window()))
        time.sleep(15)

def main():

    track_timer = threading.Thread(target=track_time_thread)
    track_timer.daemon = True
    track_timer.start()

    save_timer = threading.Thread(target=save_track_thread)
    save_timer.daemon = True
    save_timer.start()

    while isAlive:
        time.sleep(1)

if __name__ == '__main__':
    try:
        main()
    except (KeyboardInterrupt):
        kill()


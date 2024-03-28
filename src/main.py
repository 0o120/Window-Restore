import win32gui
import win32api
import win32con
import ctypes
from ctypes import POINTER, windll, Structure, cast, CFUNCTYPE, c_int, c_uint, c_void_p, c_bool
from ctypes.wintypes import HANDLE, DWORD
from cguid import GUID
import time
import json

from threading import Thread
import PySimpleGUI as sg
from psgtray import SystemTray
import os.path
from os import getcwd
from sys import exit
from config import APP_NAME, APP_ICON, APP_ICON_SYS_TRAY


### DEFAULT CONFIG ###

CONFIG = {
    'mute_on_screen_off': True,
    'seconds_until_sleep': 60 * 5,
    'seconds_restore_check': 8,
    'restore_target_success_checks': 2,
    'debug': True,
    'config_file': None
}

##############


# def _message(self, code, flags, **kwargs):
#     """Sends a message the the systray icon.

#     This method adds ``cbSize``, ``hWnd``, ``hId`` and ``uFlags`` to the
#     message data.

#     :param int message: The message to send. This should be one of the
#         ``NIM_*`` constants.

#     :param int flags: The value of ``NOTIFYICONDATAW::uFlags``.

#     :param kwargs: Data for the :class:`NOTIFYICONDATAW` object.
#     """
#     win32.Shell_NotifyIcon(code, win32.NOTIFYICONDATAW(
#         cbSize=ctypes.sizeof(win32.NOTIFYICONDATAW),
#         hWnd=self._hwnd,
#         hID=id(self),
#         uFlags=flags,
#         **kwargs))
    
# def _notify(self, message, title=None):
#     self._message(
#         win32.NIM_MODIFY,
#         win32.NIF_INFO,
#         szInfo=message,
#         szInfoTitle=title or self.title or '')


restart_monitor_thread = False
kill_monitor_thread = False

def get_file_path(filename):
    paths = [os.path.dirname(__file__), getcwd(), '..', '', '.']
    for p in paths:
        filepath = os.path.abspath(os.path.join(p, filename))
        if os.path.exists(filepath):
            return filepath
    return filename

def load_config():
    CONFIG['config_file'] = CONFIG['config_file'] or get_file_path('config.json')
    if os.path.exists(CONFIG['config_file']):
        with open(CONFIG['config_file']) as f:
            try:
                CONFIG.update(json.load(f))
            except:
                pass
    else:
        CONFIG['config_file'] = None

def save_config():
    CONFIG['config_file'] = CONFIG['config_file'] or get_file_path('config.json')
    with open(CONFIG['config_file'], 'w') as f:
        json.dump(CONFIG, f, indent=4)

def change_system_tray_menu(self, menu):
        self.close()
        # menu_items = self._convert_psg_menu_to_tray(menu[1])
        self.__init__(menu=menu, tooltip=self.tooltip, single_click_events=self.single_click_events_enabled, window=self.window, key=self.key)

setattr(SystemTray, 'change_system_tray_menu', change_system_tray_menu)


def system_tray():
    global restart_monitor_thread, kill_monitor_thread
    layout = [
        [
            sg.Column([
                [sg.Text('Idle Time to Sleep (Minutes):')],
                [sg.Text('Restore Check Duration (Seconds):')],
                [sg.Text('Restore Target Success Checks:')],
                [sg.Text('Mute System Audio on Sleep:')],
            ], size=(None, None), scrollable=False, expand_x=True, expand_y=True),
            sg.Column([
                [sg.Input('30', size=(4, 1), tooltip='Enter minutes until computer goes to sleep', key='-MINUTES-TO-SLEEP-')],
                [sg.Input('8', size=(4, 1), tooltip='Seconds to keep checking for out of place windows after PC wake up', key='-SECONDS-CHECKING-')],
                [sg.Input('3', size=(4, 1), tooltip='How many checks until restore is considered successful', key='-SUCCESS-CHECKS-')],
                [sg.Checkbox('', default=True, tooltip='Mute audio when computer goes to sleep', key='-MUTE-AUDIO-ON-SLEEP-')],
            ], size=(None, None), scrollable=False, expand_x=True)
        ],

        # [sg.Multiline(size=(60,10), reroute_stdout=False, reroute_cprint=True, write_only=True, key='-OUT-')],
        [sg.Button('Save & Close', expand_x=True), sg.Button('Close', expand_x=True)]
    ]
    window = sg.Window(f'{APP_NAME} - Settings', layout, size=(340, 180), finalize=True, enable_close_attempted_event=True, icon=get_file_path(APP_ICON))
    menu = ['Settings', '---', 'Enable', '!Disabled', '---', 'Exit']
    tooltip = f'{APP_NAME}'
    tray = SystemTray(['', menu], single_click_events=False, window=window, tooltip=tooltip, icon=get_file_path(APP_ICON_SYS_TRAY))
    # tray.show_message('System Tray', f'{(os.path.join(os.path.dirname(__file__), "icon.ico"))}')
    tray.show_icon()

    window['-MINUTES-TO-SLEEP-'].update(CONFIG['seconds_until_sleep'] // 60 if CONFIG['seconds_until_sleep'] >= 60 else round(CONFIG['seconds_until_sleep'] / 60, 2))
    window['-SECONDS-CHECKING-'].update(CONFIG['seconds_restore_check'])
    window['-SUCCESS-CHECKS-'].update(CONFIG['restore_target_success_checks'])
    window['-MUTE-AUDIO-ON-SLEEP-'].update(CONFIG['mute_on_screen_off'])


    while True:

        event, values = window.read()
    
        if event == tray.key:
            # sg.cprint(f'System Tray Event = ', values[event], c='white on red')
            event = values[event]

        if event in (sg.WIN_CLOSED, 'Exit'):
            kill_monitor_thread = True
            break
        
        if event in ('Settings', sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED):
            window['-MINUTES-TO-SLEEP-'].update(CONFIG['seconds_until_sleep'] // 60 if CONFIG['seconds_until_sleep'] >= 60 else round(CONFIG['seconds_until_sleep'] / 60, 2))
            window['-SECONDS-CHECKING-'].update(CONFIG['seconds_restore_check'])
            window['-SUCCESS-CHECKS-'].update(CONFIG['restore_target_success_checks'])
            window['-MUTE-AUDIO-ON-SLEEP-'].update(CONFIG['mute_on_screen_off'])
            window.un_hide()
            window.bring_to_front()
        elif event == 'Enable':
            menu[menu.index('!Disabled')] = 'Disable'
            menu[menu.index('Enable')] = '!Enabled'
            tray.change_system_tray_menu(['', menu])
        elif event == 'Disable':
            menu[menu.index('!Enabled')] = 'Enable'
            menu[menu.index('Disable')] = '!Disabled'
            tray.change_system_tray_menu(['', menu])
        elif event == 'Change Tooltip':
            tray.set_tooltip(values['-IN-'])
        elif event in (sg.WIN_CLOSE_ATTEMPTED_EVENT, 'Close'):
            window.hide()
            tray.show_icon()
        elif event == 'Save & Close':
            CONFIG['seconds_until_sleep'] = int(float(window['-MINUTES-TO-SLEEP-'].get()) * 60)
            CONFIG['seconds_restore_check'] = int(window['-SECONDS-CHECKING-'].get())
            CONFIG['restore_target_success_checks'] = int(window['-SUCCESS-CHECKS-'].get())
            CONFIG['mute_on_screen_off'] = window['-MUTE-AUDIO-ON-SLEEP-'].get()
            save_config()
            restart_monitor_thread = False
            window.hide()
            tray.show_icon()
            tray.show_message(title=APP_NAME, message='Running in background...')


    tray.close()
    window.close()


def windows_set_mute(state):
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    CLSCTX_ALL = 7
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    volume.SetMute(state, None)


def get_window_placements():

    windows = {}

    def window(hwnd, _data):
    
        window_is_visible = win32gui.IsWindowVisible(hwnd)
        window_is_enabled = win32gui.IsWindowEnabled(hwnd)
        window_title_length = win32gui.GetWindowTextLength(hwnd)
        if not all([window_is_visible, window_is_enabled, window_title_length > 0]):
            return

        w_title = win32gui.GetWindowText(hwnd)
        w_placement = win32gui.GetWindowPlacement(hwnd)
        # w_length, w_flags, w_ptMinPosition, w_ptMaxPosition, w_rcNormalPosition = w_placement
        win_x1, win_y1, win_x2, win_y2 = win32gui.GetWindowRect(hwnd)
        win_width, win_height = win_x2-win_x1, win_y2-win_y1
        windowPos = (hwnd, 0, win_x1, win_y1, win_width, win_height, 0x0200)
        windows[hwnd] = windowPos
        # if CONFIG['debug']:
        #     print(hwnd, w_placement, w_title)
        # if 'Microsoft​ Edge' in w_title:
        #     print('Waiting...')
        #     time.sleep(5)
            # win32gui.SetWindowPlacement(hwnd, w_placement)
            # win32gui.SetWindowPos(hwnd, 0, win_x1, win_y1, win_width, win_height, 0x0200)
    
    win32gui.EnumThreadWindows(0, window, 0)
    
    return windows

# from pprint import pprint as pp
# pp(get_window_placements())
# breakpoint()

PBT_POWERSETTINGCHANGE = 0x8013
GUID_CONSOLE_DISPLAY_STATE = '{6FE69556-704A-47A0-8F24-C28D936FDA47}'
GUID_ACDC_POWER_SOURCE = '{5D3E9A59-E9D5-4B00-A6BD-FF34FF516548}'
GUID_BATTERY_PERCENTAGE_REMAINING = '{A7AD8041-B45A-4CAE-87A3-EECBB468A9E1}'
GUID_MONITOR_POWER_ON = '{02731015-4510-4526-99E6-E5A17EBD1AEA}'
GUID_SYSTEM_AWAYMODE = '{98A7F580-01F7-48AA-9C0F-44352C29E5C0}'


class POWERBROADCAST_SETTING(Structure):
    _fields_ = [("PowerSetting", GUID),
                ("DataLength", DWORD),
                ("Data", DWORD)]

def wndproc(hwnd, msg, wparam, lparam):
    print('msg', msg)
    if msg == win32con.WM_POWERBROADCAST:
        if wparam == win32con.PBT_APMPOWERSTATUSCHANGE:
            print('Power status has changed')
        if wparam == win32con.PBT_APMRESUMEAUTOMATIC:
            print('System resume')
        if wparam == win32con.PBT_APMRESUMESUSPEND:
            print('System resume by user input')
        if wparam == win32con.PBT_APMSUSPEND:
            print('System suspend')
        if wparam == PBT_POWERSETTINGCHANGE:
            print('Power setting changed...')
            settings = cast(lparam, POINTER(POWERBROADCAST_SETTING)).contents
            power_setting = str(settings.PowerSetting)
            data_length = settings.DataLength
            data = settings.Data
            if power_setting == GUID_CONSOLE_DISPLAY_STATE:
                if data == 0: print('Display off')
                if data == 1: print('Display on')
                if data == 2: print('Display dimmed')
            elif power_setting == GUID_ACDC_POWER_SOURCE:
                if data == 0: print('AC power')
                if data == 1: print('Battery power')
                if data == 2: print('Short term power')
            elif power_setting == GUID_BATTERY_PERCENTAGE_REMAINING:
                print('battery remaining: %s' % data)
            elif power_setting == GUID_MONITOR_POWER_ON:
                if data == 0: print('Monitor off')
                if data == 1: print('Monitor on')
            elif power_setting == GUID_SYSTEM_AWAYMODE:
                if data == 0: print('Exiting away mode')
                if data == 1: print('Entering away mode')
            else:
                print('unknown GUID')
        return True

    return False

def init_power_monitor():
    print("*** STARTING ***")
    hinst = win32api.GetModuleHandle(None)
    wndclass = win32gui.WNDCLASS()
    wndclass.hInstance = hinst
    wndclass.lpszClassName = "testWindowClass"
    # CMPFUNC = CFUNCTYPE(c_bool, c_int, c_uint, c_uint, c_void_p)
    # wndproc_pointer = CMPFUNC(wndproc)
    wndclass.lpfnWndProc = {win32con.WM_POWERBROADCAST : wndproc, win32con.PBT_APMRESUMESUSPEND: wndproc}
    try:
        myWindowClass = win32gui.RegisterClass(wndclass)
        hwnd = win32gui.CreateWindowEx(win32con.WS_EX_LEFT,
                                     myWindowClass, 
                                     "testMsgWindow", 
                                     0, 
                                     0, 
                                     0, 
                                     win32con.CW_USEDEFAULT, 
                                     win32con.CW_USEDEFAULT, 
                                     0, 
                                     0, 
                                     hinst, 
                                     None)
    except Exception as e:
        print("Exception: %s" % str(e))

    if hwnd is None:
        print("hwnd is none!")
    else:
        print("hwnd: %s" % hwnd)

    guids_info = {
                    'GUID_MONITOR_POWER_ON' : GUID_MONITOR_POWER_ON,
                    'GUID_SYSTEM_AWAYMODE' : GUID_SYSTEM_AWAYMODE,
                    'GUID_CONSOLE_DISPLAY_STATE' : GUID_CONSOLE_DISPLAY_STATE,
                    'GUID_ACDC_POWER_SOURCE' : GUID_ACDC_POWER_SOURCE,
                    'GUID_BATTERY_PERCENTAGE_REMAINING' : GUID_BATTERY_PERCENTAGE_REMAINING
                 }
    for name, guid_info in guids_info.items():
        result = windll.user32.RegisterPowerSettingNotification(HANDLE(hwnd), GUID(guid_info), DWORD(0))
        print('registering', name)
        print('result:', hex(result))
        print('lastError:', win32api.GetLastError())
        print()

def get_number_of_monitors():
    user32 = ctypes.windll.user32
    return user32.GetSystemMetrics(80)

def get_virtual_screen_size():
    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79
    user32 = ctypes.windll.user32
    screensize = user32.GetSystemMetrics(SM_XVIRTUALSCREEN), user32.GetSystemMetrics(SM_YVIRTUALSCREEN), user32.GetSystemMetrics(SM_CXVIRTUALSCREEN), user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
    return screensize

def get_cursor_position():
    return win32api.GetCursorPos()

def set_cursor_position(x, y):
    win32api.SetCursorPos((x, y))

def turn_off_screen():
    return win32gui.SendMessageTimeout(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con.SC_MONITORPOWER, 2, 0, 5)

def turn_on_screen():
    cursor_position = get_cursor_position()
    set_cursor_position(0, 0)
    set_cursor_position(*cursor_position)
    return win32gui.SendMessageTimeout(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con.SC_MONITORPOWER, -1, 0, 5)

def getIdleTime():
    return (win32api.GetTickCount() - win32api.GetLastInputInfo()) / 1000.0

def getUptime():
    return win32api.GetTickCount() / 1000.0


class ScreenState:
    UNKNOWN = -1
    OFF = 1
    ON = 2
    VALS = {
        UNKNOWN: 'UNKNOWN',
        OFF: 'OFF',
        ON: 'ON'
    }


screen_state = ScreenState.UNKNOWN
screen_state_text = None
screen_state_last = ScreenState.UNKNOWN
time_screen_on_last_called = None
time_screen_off_last_called = None
seconds_idle_time = None
last_cursor_position = None

last_known_settings = {
    'window_placements': None,
    'last_cursor_position': None,
    'num_monitors': None,
    'virtual_screen_size': None
}

def debug_display():
    global screen_state_last
    time.sleep(1)
    while True:
        if kill_monitor_thread:
            return
        if screen_state != screen_state_last:
            print('screen_state_text:', screen_state_text)
            print('screen_state:', screen_state)
            print('seconds_idle_time:', seconds_idle_time)
            print('time_screen_off_last_called:', time_screen_off_last_called)
            print('seconds_until_sleep:', CONFIG['seconds_until_sleep'])

            print('get_number_of_monitors:', get_number_of_monitors())
            print('get_virtual_screen_size:', get_virtual_screen_size())
            print('get_cursor_position:', get_cursor_position())

            print('getIdleTime:', getIdleTime())
            print('getUptime:', getUptime())
            screen_state_last = screen_state
            print('---')
        time.sleep(2)


def before_screen_off():
    global last_known_settings
    last_known_settings.update({
        'window_placements': get_window_placements(),
        'last_cursor_position': get_cursor_position(),
        'num_monitors': get_number_of_monitors(),
        'virtual_screen_size': get_virtual_screen_size()
    })
    if CONFIG['mute_on_screen_off']:
        windows_set_mute(True)


def after_screen_off():
    pass

def before_screen_on():
    pass

def after_screen_on():
    """
    
        When restoring windows, compare against current window placements
        to avoid restoring windows that have been moved by the user.
    """
    time.sleep(1)
    hwnd_restored = {}
    monitors_restored = False
    monitors_restored_time = None
    cursor_restored = False
    cursor_restores = 0
    window_placements = last_known_settings['window_placements']
    is_initial_check = True
    success_checks = 0
    if CONFIG['mute_on_screen_off']:
        windows_set_mute(False)
    if window_placements:
        time_updated = False
        time_start = time.time()
        while True:
            failed_restores = 0
            cursor_position = get_cursor_position()
            num_monitors = get_number_of_monitors()
            virtual_screen_size = get_virtual_screen_size()
            if num_monitors == last_known_settings['num_monitors'] and virtual_screen_size == last_known_settings['virtual_screen_size']:
                monitors_restored = True
                monitors_restored_time = monitors_restored_time or time.time()
            if monitors_restored and not time_updated:
                time_start = time.time()
                time_updated = True
            monitors_restored_seconds = 0 if not monitors_restored_time else time.time() - monitors_restored_time
            if monitors_restored and monitors_restored_seconds > 1:
                for hwnd, w_placement in get_window_placements().items():
                    if hwnd not in hwnd_restored:
                        hwnd_restored[hwnd] = 0
                    if hwnd in window_placements:
                        # if hwnd_restored[hwnd] > 0 and window_placements[hwnd] != w_placement:
                        #     failed_restores += 1
                            # continue
                        if window_placements[hwnd] != w_placement:
                            if hwnd_restored[hwnd] > 0:
                                failed_restores += 1
                            hwnd_restored[hwnd] += 1
                            restored_position = window_placements[hwnd]
                            try:
                                win32gui.SetWindowPos(*restored_position)
                                print('Restored window placement:', hwnd, restored_position)
                            except:
                                pass
                        
                if not cursor_restored and cursor_position != last_known_settings['last_cursor_position'] and cursor_restores < 2:
                    set_cursor_position(*last_known_settings['last_cursor_position'])
                    cursor_restores += 1
                    cursor_restored = True
                if get_cursor_position() != last_known_settings['last_cursor_position']:
                    cursor_restored = False
            time.sleep(.50 if monitors_restored else .10)
            print('Checking for window placements...')
            print(last_known_settings['last_cursor_position'])

            seconds_since_check_started = time.time() - time_start

            if not is_initial_check and failed_restores == 0:
                success_checks += 1

            if failed_restores > 0 and seconds_since_check_started < (CONFIG['seconds_restore_check'] * 2):
                print('FOUND FAILS')
                success_checks -= 1
                continue


            if success_checks > CONFIG['restore_target_success_checks'] and not is_initial_check and monitors_restored and failed_restores == 0 and seconds_since_check_started > (CONFIG['seconds_restore_check'] // 2):
                print('FINISHED2')
                break
            
            if monitors_restored and seconds_since_check_started > CONFIG['seconds_restore_check']:
                print('FINISHED')
                break
            if not monitors_restored and seconds_since_check_started > (CONFIG['seconds_restore_check'] * 2):
                break

            is_initial_check = False



def seconds_since_screen_off():
    tnow = time.time()
    return tnow - (time_screen_off_last_called or tnow)




def screen_monitor():
    global kill_monitor_thread, restart_monitor_thread, screen_state, screen_state_text, time_screen_off_last_called, time_screen_on_last_called, seconds_idle_time
    load_config()
    while True:

        if kill_monitor_thread:
            return

        if restart_monitor_thread:
            restart_monitor_thread = False
            return Thread(target=screen_monitor).start()
        
        seconds_idle_time = getIdleTime()
        print(seconds_idle_time, CONFIG['seconds_until_sleep'])

        if seconds_idle_time < 2 and time_screen_off_last_called is None:
            screen_state_text, screen_state = ScreenState.VALS[ScreenState.ON], ScreenState.ON

        if seconds_idle_time < 2 and screen_state not in [ScreenState.ON, ScreenState.UNKNOWN]:
            before_screen_on()
            screen_state_text, screen_state = ScreenState.VALS[ScreenState.ON], ScreenState.ON
            time_screen_on_last_called = time.time()
            after_screen_on()

        if seconds_idle_time > CONFIG['seconds_until_sleep'] and screen_state in [ScreenState.ON, ScreenState.UNKNOWN]:
            before_screen_off()
            turn_off_screen()
            screen_state_text, screen_state = ScreenState.VALS[ScreenState.OFF], ScreenState.OFF
            time_screen_off_last_called = time.time()
            after_screen_off()
        
        time.sleep(2)

if __name__ == "__main__":
    
    # init_power_monitor()

    if CONFIG['debug']:
        Thread(target=debug_display).start()
    
    Thread(target=screen_monitor).start()
    
    system_tray()

    exit(0)
    

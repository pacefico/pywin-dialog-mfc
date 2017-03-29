"""
pywin 32 Native Window Sample
ref: http://auralbits.blogspot.com.br/2014/02/a-native-win32-application-in-python.html
"""

import os
import struct
import win32gui_struct
import win32api
import win32con
import winerror
import win32gui
import time
import logging
from sys import argv

from threading import Lock

logger = None


win32gui.InitCommonControls()
hInstance = win32gui.dllhandle

# Load the RichEdit 4.1 control
win32api.LoadLibrary('MSFTEDIT.dll')

IDC_EDIT = 1024

class DemoWindow(object):
    class_atom = None
    class_name = "Pywin32DialogDemo"
    title = "pywin32 Dialog Demo"

    def __init__(self, log_file_name=""):
        self.set_logger(log_file_name)
        message_map = {
            win32con.WM_SIZE: self.on_size,
            win32con.WM_COMMAND: self.on_command,
            win32con.WM_NOTIFY: self.on_notify,
            win32con.WM_INITDIALOG: self.on_init_dialog,
            win32con.WM_CLOSE: self.on_close,
            win32con.WM_DESTROY: self.on_destroy,
            win32con.WM_KEYDOWN: self.on_keydown,
            win32con.WM_KEYUP: self.on_keyup,
            win32con.WM_GETTEXT: self.on_gettext
        }

        self._register_wnd_class()
        template = self._get_dialog_template()

        # Create the window via CreateDialogBoxIndirect - it can then
        # work as a "normal" window, once a message loop is established.
        win32gui.CreateDialogIndirect(hInstance, template, 0, message_map)

    def set_logger(self, name):
        if name != "":
            if os.path.exists(name):
                os.remove(name)
            log_file_name = name
        else:
            log_file_tpl = "mywindow_{}.log"
            count = 1
            log_file_name = log_file_tpl.format(count)
            while True:
                if not os.path.exists(log_file_name):
                    break
                count += 1
                log_file_name = log_file_tpl.format(count)

        global logger
        logger = logging.getLogger('mywindow')
        print("log file: {}".format(log_file_name))

        hdlr = logging.FileHandler(log_file_name)
        formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
        hdlr.setFormatter(formatter)
        logger.addHandler(hdlr)
        logger.setLevel(logging.INFO)

    @classmethod
    def _register_wnd_class(cls):
        if cls.class_atom:
            return

        message_map = {}
        wc = win32gui.WNDCLASS()
        wc.SetDialogProc()  # Make it a dialog class.
        wc.hInstance = hInstance
        wc.lpszClassName = cls.class_name
        wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
        wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        wc.hbrBackground = win32con.COLOR_WINDOW + 1
        wc.lpfnWndProc = message_map  # could also specify a wndproc.
        wc.cbWndExtra = win32con.DLGWINDOWEXTRA + struct.calcsize("Pi")

        # load icon from python executable
        this_app = win32api.GetModuleHandle(None)
        wc.hIcon = win32gui.LoadIcon(this_app, 1)

        try:
            cls.class_atom = win32gui.RegisterClass(wc)
        except win32gui.error as err_info:
            if err_info.winerror != winerror.ERROR_CLASS_ALREADY_EXISTS:
                raise

    @classmethod
    def _get_dialog_template(cls):
        style = (win32con.WS_THICKFRAME | win32con.WS_POPUP |
                 win32con.WS_VISIBLE | win32con.WS_CAPTION |
                 win32con.WS_SYSMENU | win32con.DS_SETFONT |
                 win32con.WS_MINIMIZEBOX)

        # These are "dialog coordinates," not pixels. Sigh.
        bounds = 0, 0, 250, 210

        # Window frame and title, a PyDLGTEMPLATE object
        dialog = [(cls.title, bounds, style, None, None, None, cls.class_name), ]

        return dialog

    def on_notify(self, hwnd, msg, wparam, lparam):
        self.log("on_notify", hwnd, msg, wparam, lparam)

        info = win32gui_struct.UnpackNMITEMACTIVATE(lparam)

        if wparam != IDC_EDIT:
            print('origin not edit, skipping')
            return 1

        if info.code == win32con.EN_MSGFILTER:
            logger.info('EN_MSGFILTER')
        elif info.code == win32con.EN_SELCHANGE:
            logger.info('EN_SELCHANGE')

        print(info.code)
        return 1

    def on_command(self, hwnd, msg, wparam, lparam):
        self.log("on_command", hwnd, msg, wparam, lparam)

        msg_id = win32api.HIWORD(wparam)
        print(msg_id)

        if lparam != self.edit_hwnd:
            logger.info('origin not edit, skipping')
            return 1

        if msg_id == win32con.EN_CHANGE:
            logger.info('EN_CHANGE')
        elif msg_id == win32con.EN_UPDATE:
            logger.info('EN_UPDATE')

        return 1

    def on_init_dialog(self, hwnd, msg, wparam, lparam):
        self.log("on_init_dialog", hwnd, msg, wparam, lparam)
        self.hwnd = hwnd
        self._setup_edit()
        l, t, r, b = win32gui.GetWindowRect(self.hwnd)
        self._do_size(r, b, 1)

    def _setup_edit(self):
        class_name = 'RICHEDIT50W'
        initial_text = ''
        child_style = (win32con.WS_CHILD | win32con.WS_VISIBLE |
                       win32con.WS_HSCROLL | win32con.WS_VSCROLL |
                       win32con.WS_TABSTOP | win32con.WS_VSCROLL |
                       win32con.ES_MULTILINE | win32con.ES_WANTRETURN)
        parent = self.hwnd

        self.edit_hwnd = win32gui.CreateWindow(
            class_name, initial_text, child_style,
            0, 0, 0, 0,
            parent, IDC_EDIT, hInstance, None)

    def _do_size(self, cx, cy, repaint=1):
        # expand the textbox to fill the window
        win32gui.MoveWindow(self.edit_hwnd, 0, 0, cx, cy, repaint)

    def on_close(self, hwnd, msg, wparam, lparam):
        win32gui.DestroyWindow(hwnd)

    def on_destroy(self, hwnd, msg, wparam, lparam):
        self.log("on_destroy", hwnd, msg, wparam, lparam)
        # Terminate the app
        win32gui.PostQuitMessage(0)

    def on_size(self, hwnd, msg, wparam, lparam):
        x = win32api.LOWORD(lparam)
        y = win32api.HIWORD(lparam)
        self._do_size(x, y)
        return 1

    def on_keydown(self, hwnd, msg, wparam, lparam):
        self.log("on_keydown", hwnd, msg, wparam, lparam)

    def on_keyup(self, hwnd, msg, wparam, lparam):
        self.log("on_keyup", hwnd, msg, wparam, lparam)

    def on_gettext(self, hwnd, msg, wparam, lparam):
        self.log("on_gettext", hwnd, msg, wparam, lparam)

    def log(self, name, param1="", param2="", param3="", param4=""):
        print('{} {} {} {} {}'.format(name, param1, param2, param3, param4))
        logger.info('{} {} {} {} {}'.format(name, param1, param2, param3, param4))

    def minimize(self):
        return win32gui.CloseWindow(self.hwnd)


def run_function_thread():
    import win32gui
    import threading

    def run_window_dialog():
        global window_dialog
        window_dialog = DemoWindow()
        win32gui.PumpMessages()

    thread = threading.Thread(target=run_window_dialog)
    thread.start()
    thread.join()


def run_function():
    DemoWindow()
    win32gui.PumpMessages()


def run_function_subprocess():
    import threading
    import subprocess

    def run_window_dialog():
        command = "py -3 mywindow.py"
        subprocess.call(command.split(), shell=False)

    thread = threading.Thread(target=run_window_dialog)
    thread.start()
    thread.join()


def main(arg):
    if arg == "-t":
        run_function_thread()
    elif arg == "-p":
        run_function_subprocess()
    else:
        run_function()


if __name__ == '__main__':
    arg = ""
    if len(argv) > 1:
        arg = argv[1]
    main(arg)
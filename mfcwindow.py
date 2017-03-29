import win32ui, win32con, win32gui, time, win32api
import pywintypes
import threading
from sys import argv
from os import startfile
from mywindow import DemoWindow

window_class = 'Pywin32DialogDemo'
window_obj = None


def run_window_demo(name=""):
    global window_obj
    window_obj = DemoWindow(log_file_name=name)
    win32gui.PumpMessages()


def quit_message():
    hwnd = wait_for_window()
    hwnd.PostMessage(win32con.WM_QUIT, win32con.VK_F5, 0)
    exit_code = win32api.GetLastError()
    last_error = win32api.FormatMessage(exit_code)
    print("quit_message - code: {} msg: {}".format(exit_code, last_error))
    pass


def print_last_error():
    time.sleep(2)
    code = win32api.GetLastError()
    msg = win32api.FormatMessage(code)
    print("code:{} msg:{}".format(code, msg))
    return code, msg


def wait_for_window(retry=True, seconds=20):
    global window_class
    start = time.clock()
    end = 0
    last_error = ""
    while retry and end < seconds:
        end = time.clock() - start
        try:
            hwnd = win32ui.FindWindow(window_class, None)
            if hwnd and hwnd.__class__.__name__ in win32ui.types:
                return hwnd
        except win32ui.error as e:
            last_error = str(e)
            # expects this error when a window is not ready

    print("wait for window, erro: {}".format(last_error))
    return None


def find_window(retry=True, seconds=10):
    global window_class
    start = time.clock()
    end = 0
    last_error = ""
    count = 0
    while retry and end < seconds:
        end = time.clock() - start
        print("count: {}".format(count))
        count += 1
        try:
            hwnd = wait_for_window()
            if hwnd and hwnd.__class__.__name__ in win32ui.types:
                return True
        except win32ui.error as e:
            last_error = str(e)
            # expects this error when a window is not ready

    print("A janela '{}' nao foi encontrada. Erro: {}".format(window_class, last_error))
    return False


def post_message(retry=True, seconds=20):
    global window_class

    start = time.clock()
    end = 0
    last_error = ""
    count = 0
    while retry and end < seconds:
        end = time.clock() - start
        print("count: {}".format(count))
        count += 1
        try:
            hwnd = wait_for_window()
            if hwnd and hwnd.__class__.__name__ in win32ui.types:
                time.sleep(0.5)
                hwnd.PostMessage(win32con.WM_KEYDOWN, win32con.VK_F5, 0)
                first_result, first_msg = print_last_error()
                hwnd.PostMessage(win32con.WM_KEYUP, win32con.VK_F5, 0)
                second_result, second_msg = print_last_error()

                if first_result == 0 and second_result == 0:
                    return True
        except win32ui.error as e:
            last_error = str(e)
            # expects this error when a window is not ready

    print("Nao foi possivel enviar 'post_message' apos {}s. Erro: {}".format(end, last_error))
    return False


def send_message(retry=True, seconds=10):
    global window_class
    start = time.clock()
    end = 0
    last_error = ""
    count = 0
    while retry and end < seconds:
        end = time.clock() - start
        print("count: {}".format(count))
        count += 1
        try:
            hwnd = win32ui.FindWindow(window_class, None)
            if hwnd and hwnd.__class__.__name__ in win32ui.types:
                buffer = win32gui.PyMakeBuffer(255)
                length = hwnd.SendMessage(win32con.WM_GETTEXT, 255, buffer)
                first_result, first_msg = print_last_error()

                if length > 0 and first_result == 0:
                    return True
        except win32ui.error as e:
            last_error = str(e)
            # expects this error when a window is not ready
        except pywintypes.error as e:
            last_error = str(e)
            #logger.info("SendMessageTimeout - Excedeu tempo limite da operacao")
            # expects this error when a timeout has raised before operations end

    print("Nao foi possivel enviar 'SendMessage' apos {}s. Erro: {}".format(end, last_error))
    return False


def send_message_timeout(retry=True, seconds=10):
    global window_class
    start = time.clock()
    end = 0
    last_error = ""
    count = 0
    while retry and end < seconds:
        end = time.clock() - start
        print("count: {}".format(count))
        count += 1
        try:
            hwnd = win32ui.FindWindow(window_class, None)
            if hwnd and hwnd.__class__.__name__ in win32ui.types:
                time.sleep(2)
                result, resultmsg = win32gui.SendMessageTimeout(hwnd.GetSafeHwnd(),
                                                                win32con.WM_GETTEXT, 0, "test", win32con.SMTO_NORMAL, 15)
                first_result, first_msg = print_last_error()

                if result and first_result == 0:
                    return True
        except win32ui.error as e:
            last_error = str(e)
            # expects this error when a window is not ready
        except pywintypes.error as e:
            last_error = str(e)
            #logger.info("SendMessageTimeout - Excedeu tempo limite da operacao")
            # expects this error when a timeout has raised before operations end

    print("Nao foi possivel enviar 'SendMessageTimeout' apos {}s. Erro: {}".format(end, last_error))
    return False


def window_text(retry=True, seconds=10):
    global window_class
    global window_obj
    start = time.clock()
    end = 0
    last_error = ""
    count = 0
    while retry and end < seconds:
        end = time.clock() - start
        print("count: {}".format(count))
        count += 1
        try:
            hwnd = win32ui.FindWindow(window_class, None)

            if hwnd and hwnd.__class__.__name__ in win32ui.types:

                text = hwnd.GetWindowText()
                first_result, first_msg = print_last_error()

                if len(text) > 0 and first_result == 0:
                    return True
        except win32ui.error as e:
            last_error = str(e)
            # expects this error when a window is not ready

    print(last_error)
    return False


def run_func_thread(run_func):
    func_name = run_func.__name__
    print("\n#### Testing '{}'!".format(func_name))
    response = run_func()
    print("result: {}".format(response))
    return response


def run_tests():
    results = [
        (find_window.__name__, run_func_thread(find_window)),
        (post_message.__name__, run_func_thread(post_message)),
        (send_message.__name__, run_func_thread(send_message)),
        (send_message_timeout.__name__, run_func_thread(send_message_timeout)),
        (window_text.__name__, run_func_thread(window_text))
    ]
    return results


def main(thread_enabled):

    thread = None
    if thread_enabled:
        thread = threading.Thread(target=run_window_demo)
        thread.start()
        wait_for_window()

        results = run_tests()

        thread.join()
    else:
        results = run_tests()


    print("\n\nSummary {} thread:".format("com" if thread else "sem"))
    for item in results:
        print("{} = {}".format(item[0], item[1]))

    quit_message()

if __name__ == '__main__':
    thread_enabled = False
    if len(argv) > 1 and argv[1] == "-t":
        thread_enabled = True
    main(thread_enabled)
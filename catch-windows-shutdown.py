# """ Testing Windows shutdown events """
## Win10 test of https://stackoverflow.com/questions/1411186/python-windows-shutdown-events - only works with python.exe, not pythonw.exe

import infi.systray.win32_adapter as win32
import sys
import time, uuid

WM_QUERYENDSESSION = 17
WM_ENDSESSION = 22
WM_QUIT = 18

def log_info(msg):
    """ Prints """
    print(msg)
    f = open("test.log", "a")
    f.write(msg + "\n")
    f.close()

def wndproc(hwnd, msg, wparam, lparam):
    hwnd = win32.HANDLE(hwnd)
    wparam = win32.WPARAM(wparam)
    lparam = win32.LPARAM(lparam)
    log_info(f"wndproc: {msg}, wparam: {wparam}, lparam: {lparam}")
    return win32.DefWindowProc(hwnd, msg, wparam, lparam)
    

if __name__ == "__main__":
    log_info("*** STARTING ***")
    hinst = win32.GetModuleHandle(None)
    wndclass = win32.WNDCLASS()
    wndclass.hInstance = hinst
    name = win32.encode_for_locale("testWindowClass-%s" %(str(uuid.uuid4())))
    wndclass.lpszClassName = name
    wndclass.lpfnWndProc = win32.LPFN_WNDPROC(wndproc)
    #messageMap = { WM_QUERYENDSESSION : wndproc,
    #               WM_ENDSESSION : wndproc,
    #               WM_QUIT : wndproc,
    #               win32.WM_DESTROY : wndproc,
    #               win32.WM_CLOSE : wndproc }
#
#    wndclass.lpfnWndProc = messageMap

    try:
        myWindowClass = win32.RegisterClass(wndclass)
        hwnd = win32.CreateWindowEx(0,
                                     myWindowClass, 
                                     name, 
                                     0, 
                                     0, 
                                     0, 
                                     win32.CW_USEDEFAULT, 
                                     win32.CW_USEDEFAULT, 
                                     0, 
                                     0, 
                                     hinst, 
                                     None)
    except Exception as e:
        log_info("Exception: %s" % str(e))


    if hwnd is None:
        log_info("hwnd is none!")
    else:
        log_info("hwnd: %s" % hwnd)

    while True:
        win32.PumpMessages()
        time.sleep(1)
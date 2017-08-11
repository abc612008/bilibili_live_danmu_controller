import asyncio
import sys
import time
import win32api
import win32con
from bilibiliClient import bilibiliClient

VK_CODE = {'backspace':0x08,
           'tab':0x09,
           'enter':0x0D,
           'shift':0x10,
           'ctrl':0x11,
           'alt':0x12,
           'caps_lock':0x14,
           'esc':0x1B,
           'spacebar':0x20,
           'ins':0x2D,
           'del':0x2E,
           '0':0x30,
           '1':0x31,
           '2':0x32,
           '3':0x33,
           '4':0x34,
           '5':0x35,
           '6':0x36,
           '7':0x37,
           '8':0x38,
           '9':0x39,
           'a':0x41,
           'b':0x42,
           'c':0x43,
           'd':0x44,
           'e':0x45,
           'f':0x46,
           'g':0x47,
           'h':0x48,
           'i':0x49,
           'j':0x4A,
           'k':0x4B,
           'l':0x4C,
           'm':0x4D,
           'n':0x4E,
           'o':0x4F,
           'p':0x50,
           'q':0x51,
           'r':0x52,
           's':0x53,
           't':0x54,
           'u':0x55,
           'v':0x56,
           'w':0x57,
           'x':0x58,
           'y':0x59,
           'z':0x5A,
           'left_shift':0xA0,
           'right_shift ':0xA1,
           'left_control':0xA2,
           'right_control':0xA3,
           '+':0xBB,
           ',':0xBC,
           '-':0xBD,
           '.':0xBE,
           '/':0xBF,
           '`':0xC0,
           ';':0xBA,
           '[':0xDB,
           '\\':0xDC,
           ']':0xDD,
           "'":0xDE,
           '`':0xC0}

def press(arg):
    win32api.keybd_event(VK_CODE[arg], 0,0,0)
    time.sleep(.05)
    win32api.keybd_event(VK_CODE[arg],0 ,win32con.KEYEVENTF_KEYUP ,0)
        
def callback(isAdmin,isVIP,commentUser,commentText):
    MAX_MOVE_DELTA=50
    MAX_CLICK_TIME=3000
    print(commentText)
    if commentText[:2] in ["mw","ma","ms","md"]:
        data = commentText.split(" ")
        if len(data)==2:
            move_delta = min(int(data[1]),MAX_MOVE_DELTA)
        else:
            move_delta = 10
        
        if commentText[:2]=="mw":
            win32api.mouse_event(win32con.MOUSEEVENTF_MOVE,0,-move_delta)
        elif commentText[:2]=="ma":
            win32api.mouse_event(win32con.MOUSEEVENTF_MOVE,-move_delta,0)
        elif commentText[:2]=="ms":
            win32api.mouse_event(win32con.MOUSEEVENTF_MOVE,0,move_delta)
        elif commentText[:2]=="md":
            win32api.mouse_event(win32con.MOUSEEVENTF_MOVE,move_delta,0)
    elif commentText[:2]=="mc":
        data = commentText.split(" ")
        if len(data)==2:
            click_time = min(int(data[1]),MAX_CLICK_TIME)/1000
        else:
            click_time = .5
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        time.sleep(click_time)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
    elif commentText[:3]=="mrc":
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0,0,0)
    elif commentText[:1]=="k":
        press(commentText.split(" ")[1])
    
danmuji = bilibiliClient(callback,sys.argv[1])

tasks = [
            danmuji.connectServer() ,
            danmuji.HeartbeatLoop()
        ]
loop = asyncio.get_event_loop()
try:
    loop.run_until_complete(asyncio.wait(tasks))
except KeyboardInterrupt:
    danmuji.connected = False
    for task in asyncio.Task.all_tasks():
        task.cancel()
    loop.run_forever()

loop.close()

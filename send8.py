import sys
import time
import random


import win32api
import win32gui
import win32con

def windowEnumerationHandler(hwnd, resultList):

    '''Pass to win32gui.EnumWindows() to generate list of window handle, window text tuples.'''

    resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def sendKeys(handle):
    print "Sending the messages.. cross your fingers!"
    win32api.SendMessage(handle, win32con.WM_KEYDOWN, 56, 0)
    win32api.SendMessage(handle, win32con.WM_KEYUP, 56, 0)

if len(sys.argv) < 2:
    print "Missing timeout! Try 2.5 for milling, 3.5 for Prospect, or 4.5 for DE!"
    sys.exit(1)


topWindows = []
win32gui.EnumWindows(windowEnumerationHandler, topWindows)

handle = None
for wnd,title in topWindows:
    if title == "World of Warcraft":
        handle = wnd
        print "Found the window!"

if handle is None:
    print "Did not find the window, aborting! :("
    sys.exit(1)


# main loop
while True:
    sendKeys(handle)
    time.sleep(float(sys.argv[1]))
    # sleep some more, so we're more realistic!
    time.sleep(random.random() / 2)


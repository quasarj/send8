import sys
import time
import random
import argparse


import win32api
import win32gui
import win32con

def windowEnumerationHandler(hwnd, resultList):

    '''Pass to win32gui.EnumWindows() to generate list of window handle, window text tuples.'''

    resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def sendKeys(handle):
    print "Sending 8..."
    win32api.SendMessage(handle, win32con.WM_KEYDOWN, 56, 0)
    win32api.SendMessage(handle, win32con.WM_KEYUP, 56, 0)

def main(args):
    topWindows = []
    win32gui.EnumWindows(windowEnumerationHandler, topWindows)

    handle = None
    for wnd,title in topWindows:
        if title == "World of Warcraft":
            handle = wnd
            print "Found the window. Everything looks good."

    if handle is None:
        print "WoW window not found. Is it running?"
        sys.exit(1)

    # determine sleep time
    if args.mill:
        sleep = 2.4
        print "Milling mode, sleeping 2.5 seconds between sends."
    elif args.disenchant:
        sleep = 4.5
        print "Disenchant mode, sleeping 4.5 seconds between sends."
    else:
        sleep = 3.5
        print "Prospect mode, sleeping 3.5 seconds between sends. See -h for other options."

    # main loop
    for i in range(args.count):
        if i != 0:
            time.sleep(sleep)
            # sleep some more, so we're more realistic!
            time.sleep(random.random() / 2)
        sendKeys(handle)
        print "{} 8s left.".format(args.count - (i + 1))

parser = argparse.ArgumentParser(description="Send 8.")
parser.add_argument('--mill', '-m', action='store_const', const=True,
    help="Milling mode. Waits at least 2.5 seconds between sends.")
parser.add_argument('--disenchant', '-d', action='store_const', const=True,
    help="Disenchant mode. Waits at least 4.5 seconds between sends.")
parser.add_argument('--prospect', '-p', action='store_const', const=True,
    help="Prospecting mode. Waits at least 3.5 seconds between sends. This is the default mode.")
parser.add_argument('count', type=int, help="Number of times to send 8.")

args = parser.parse_args()

print "Sending 8 {} times!".format(args.count)
main(args)


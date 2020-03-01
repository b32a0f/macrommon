# -*- coding: utf-8 -*-

# RGB convert https://stackoverflow.com/questions/2262100/rgb-int-to-rgb-python
# Mouse Click https://stackoverflow.com/questions/37622438/emulating-a-mouseclick-using-postmessage
# Virtual Key https://docs.microsoft.com/ko-kr/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN

import time
import random
import win32api as wa, win32con as wc, win32gui as wg
import pywintypes

# Find window handle.
def getHwnd(pa_class, pa_caption, ch_class=None, ch_caption=None):
    pahwnd = wg.FindWindow(pa_class, pa_caption)
    chhwnd = wg.FindWindowEx(pahwnd, None, ch_class, ch_caption)

    return (pahwnd, chhwnd)

# Divide 24bit color integer into each 8bit integer tuple. (0bBBBBBBBBGGGGGGGGRRRRRRRR)
def bgr2rgb(BGRint):
    red   = (BGRint      ) & 255
    green = (BGRint >>  8) & 255
    blue  = (BGRint >> 16) & 255

    return (red, green, blue)

# Get pixel color.
# 
# Fixed issue failing call GetPixel 10,000 times.
# If keyboard or mouse input occurred, Sometimes has error.
# For avoid it, call function recursively.
# 
# [Issue Refs]
# https://stackoverflow.com/questions/19623135/pywin32-win32gui-getpixel-fails-predictably-near-10-000th-call
# https://stackoverflow.com/questions/48735829/python-win32gui-getpixel-fails-when-righ-clicking-with-mouse
# https://stackoverflow.com/questions/6701087/ambiguous-pywintypes-error-when-calling-win32gui-getpixel
# 
def getColor(hwnd=wg.GetDesktopWindow(), x=0, y=0):
    try:
        hdc   = wg.GetWindowDC(hwnd)
        color = wg.GetPixel(hdc, x, y)

        wg.ReleaseDC(hwnd, hdc)

        return bgr2rgb(color)
    except pywintypes.error:
        return getColor(hwnd, x, y)

# Compare pixel color.
def colorMatch(hwnd=wg.GetDesktopWindow(), x=0, y=0, rgb=(0, 0, 0)):
    if getColor(hwnd, x, y) == rgb:
        return True
    else:
        return False

# Wait for pixel color.
def colorWait(hwnd=wg.GetDesktopWindow(), x=0, y=0, rgb=(0, 0, 0), match=True, inter=0):
    if match:
        while not colorMatch(hwnd, x, y, rgb):
            time.sleep(0.1)
    else:
        while colorMatch(hwnd, x, y, rgb):
            time.sleep(0.1)

    time.sleep(inter)

# Random click.
def randClick(hwnd=wg.GetDesktopWindow(), x=0, y=0, tol=10, inter=0.3):
    rx     = x + random.randint(-tol, tol)
    ry     = y + random.randint(-tol, tol)
    lParam = wa.MAKELONG(rx, ry)
    
    wg.PostMessage(hwnd, wc.WM_LBUTTONDOWN, wc.MK_LBUTTON, lParam)
    wg.PostMessage(hwnd, wc.WM_LBUTTONUP  , 0            , lParam)

    time.sleep(inter)

# Random Drag.
def randDrag(hwnd=wg.GetDesktopWindow(), fx=0, fy=0, tx=0, ty=0, tol=10, inter=0.3):
    rfx      = fx + random.randint(-tol, tol)
    rfy      = fy + random.randint(-tol, tol)
    rtx      = tx + random.randint(-tol, tol)
    rty      = ty + random.randint(-tol, tol)
    f_lParam = wa.MAKELONG(rfx, rfy)
    t_lParam = wa.MAKELONG(rtx, rty)
    
    wg.PostMessage(hwnd, wc.WM_LBUTTONDOWN, wc.MK_LBUTTON, f_lParam)

    for mid in range(20):
        mx = ((rtx - rfx) // 20) * mid
        my = ((rty - rfy) // 20) * mid
        m_lParam = wa.MAKELONG(mx, my)

        wg.PostMessage(hwnd, wc.WM_MOUSEMOVE, 0, m_lParam)

        time.sleep(0.01)

    wg.PostMessage(hwnd, wc.WM_LBUTTONUP  , 0            , t_lParam)

    time.sleep(inter)
    
# Input Keyboard.
def sendKey(hwnd=wg.GetDesktopWindow(), key=0x16, inter=0.03):
    wg.PostMessage(hwnd, wc.WM_KEYDOWN, key, 0)
    wg.PostMessage(hwnd, wc.WM_KEYUP  , key, 0)

    time.sleep(inter)

# For develop.
def getPos(hwnd):
    print(wg.ScreenToClient(hwnd, wg.GetCursorPos()))

# -*- coding: utf-8 -*-

"""
Inactive automation for Windows.

Using windows API, click/drag or get pixel color.
Mouse actions are random based.
"""

import time
import random
import win32api as wa, win32con as wc, win32gui as wg, win32ui as wu
import pywintypes
import ctypes
from PIL import Image

class Macrommon:
    def __init__(self, parent_class=None, parent_title=None, child_class=None, child_title=None):
        if parent_class or parent_title:
            self.setHwnd(parent_class, parent_title)

            if child_class or child_title:
                self.setHwnd(child_class, child_title, child_mode=True)
        else:
            self.resetHwnd()

        self._size = self.window_size if self.hwnd else None
        self._zero = (0, 0)

    # ---------------- #
    #       HWND       #
    # ---------------- #

    @property
    def hwnd(self) -> int:
        """Return current window handle."""

        return self._hwnd

    @property
    def hwnd_info(self) -> tuple or None:
        """Return current window class and title."""

        try:
            _wnd_class = wg.GetClassName(self._hwnd)
            _wnd_title = wg.GetWindowText(self._hwnd)

            return _wnd_class, _wnd_title
        except pywintypes.error:
            return None

    def hwnd_list(self):
        def callback(hwnd, hwnds):
            if wg.IsWindowVisible(hwnd):
                _wnd_class = wg.GetClassName(hwnd)
                _wnd_title = wg.GetWindowText(hwnd)

                hwnds.append((hwnd, _wnd_class, _wnd_title))

        _root_hwnds  = []
        _child_hwnds = []

        wg.EnumWindows(callback, _root_hwnds)

        for _root in _root_hwnds:
            wg.EnumChildWindows(_root[0], callback, _child_hwnds)

        return _root_hwnds + _child_hwnds

    def setHwnd(self, wnd_class=None, wnd_title=None, child_mode=False, delay=0):
        """Find window handle."""

        if not child_mode:
            self._hwnd = wg.FindWindow(wnd_class, wnd_title)
        else:
            self._hwnd = wg.FindWindowEx(self._hwnd, None, wnd_class, wnd_title)

        time.sleep(delay)

    def setHwndFromCursor(self, delay=0):
        """Find window handle from cursor position."""

        input('Press ENTER after cursor on position.')

        self._hwnd = wg.WindowFromPoint(wg.GetCursorPos())

        time.sleep(delay)

    def setHwndFromPosition(self, x, y, delay=0):
        """Find window handle from position."""

        self._hwnd = wg.WindowFromPoint((x, y))

        time.sleep(delay)

    def resetHwnd(self, delay=0):
        """Reset window handle to Desktop."""

        self._hwnd = wg.GetDesktopWindow()

        time.sleep(delay)

    # ---------------- #
    #     POSITION     #
    # ---------------- #

    def _offset(self, x, y, abs=True) -> tuple:
        """
        Adjust coordinate point.
        (x, y) is absolute point.

        If abs is True, return absolute point from (0, 0).
        Otherwise, return relative point from current zero point.

        Axial direction is same as axis of display.

                -y
                 ↑
            -x ← 0 → +x
                 ↓
                +y

        Args:
            x, y: Position x, y
            abs : Switch normal coodination to calculate.

        Returns:
            (x, y)
        """

        if abs:
            _x = (self._zero[0] + x)
            _y = (self._zero[1] + y)
        else:
            _x = (self._zero[0] - x) * -1
            _y = (self._zero[1] - y) * -1

        return _x, _y

    @property
    def window_rect(self) -> tuple:
        """Return window rect."""

        return wg.GetWindowRect(self._hwnd) if self._hwnd else None

    @property
    def window_size(self) -> tuple:
        """Return window size."""

        if self._hwnd:
            _rect = self.window_rect
            _x    = _rect[2] - _rect[0]
            _y    = _rect[3] - _rect[1]

            return _x, _y
        else:
            return None

    @property
    def zero(self) -> tuple:
        """Return zero point."""

        return self._zero

    def setZero(self, zero: str or tuple='lu'):
        """
        Set zero point.

        Args:
            zero: Position (x, y)
                  string   'lu' (Left Up), 'ld' (Left Down), 'ru' (Right Up), 'rd' (Right Down), 'ct' (Centre).
        """

        _size = self.window_size

        if type(zero) == tuple:
            _zero = zero
        elif type(zero) == str:
            if zero == 'lu':
                _zero = (0, 0)
            elif zero == 'ld':
                _zero = (0, _size[1])
            elif zero == 'ru':
                _zero = (_size[0], 0)
            elif zero == 'rd':
                _zero = _size
            elif zero == 'ct':
                _zero = (_size[0] // 2, _size[1] // 2)

        self._zero = _zero

    def moveWindow(self, x, y, w, h, delay=0):
        """
        Set window position and size.

        Args:
            x, y: Position x, y
            w, h: Size     width, height
        """

        wg.MoveWindow(self._hwnd, x, y, w, h, True)

        time.sleep(delay)

    def getPos(self, with_color=False) -> tuple:
        """
        Return cursor position of client.
        
        Args:
            with_color: Show with pixel color info.

        Returns:
            (x, y)
            (x, y, (R, G, B))
        """

        _x, _y = wg.ScreenToClient(self._hwnd, wg.GetCursorPos())
        _x, _y = self._offset(_x, _y, False)

        if not with_color:
            return _x, _y
        else:
            return _x, _y, self.getColor(_x, _y)

    # ---------------- #
    #      COLOR       #
    # ---------------- #

    def _bgr2rgb(self, BGRint) -> tuple:
        """
        Divide 24bit color integer into each 8bit integer tuple.

        Args:
            BGRint: Windows API Color system (0bBBBBBBBBGGGGGGGGRRRRRRRR)

        Returns:
            (R, G, B)
        """

        # RGB convert https://stackoverflow.com/questions/2262100/rgb-int-to-rgb-python

        _red   = (BGRint      ) & 255
        _green = (BGRint >>  8) & 255
        _blue  = (BGRint >> 16) & 255

        return (_red, _green, _blue)

    def getColor(self, x=0, y=0) -> tuple:
        """
        Return pixel color.

        Fixed issue failing call GetPixel 10,000 times.
        If keyboard or mouse input occurred, Sometimes has error.
        To avoid it, call recursively.

        Args:
            x, y : Position x, y

        Returns:
            (R, G, B)
        """

        # [Issue Refs]
        # https://stackoverflow.com/questions/19623135/pywin32-win32gui-getpixel-fails-predictably-near-10-000th-call
        # https://stackoverflow.com/questions/48735829/python-win32gui-getpixel-fails-when-righ-clicking-with-mouse
        # https://stackoverflow.com/questions/6701087/ambiguous-pywintypes-error-when-calling-win32gui-getpixel

        try:
            _x, _y = self._offset(x, y)
            _hdc   = wg.GetWindowDC(self._hwnd)
            _color = wg.GetPixel(_hdc, _x, _y)

            wg.ReleaseDC(self._hwnd, _hdc)

            return self._bgr2rgb(_color)
        except pywintypes.error:
            return self.getColor(x, y)

    def colorMatch(self, x=0, y=0, rgb=(0, 0, 0), tol=10) -> bool:
        """
        Compare pixel color.

        Args:
            x, y : Position x, y
            rgb  : Color    (R, G, B)
            tol  : Range    [-tol ~ +tol]

        Returns:
            True or False
        """

        _fr, _fg, _fb = rgb
        _tr, _tg, _tb = self.getColor(x, y)
        _cmp_tol      = [abs(_fr - _tr), abs(_fg - _tg), abs(_fb - _tb)]

        if max(_cmp_tol) <= tol:
            return True
        else:
            return False

    def colorWait(self, x=0, y=0, rgb=(0, 0, 0), match=True, tol=10, inter=0.1, delay=0):
        """
        Wait for pixel color.
        
        Args:
            x, y : Position x, y
            rgb  : Color    (Red, Green, Blue)
            match: Switch match or not match.
            tol  : Range    [-tol ~ +tol]
            inter: Sleep to interval.
        """

        if match:
            while not self.colorMatch(x, y, rgb, tol):
                time.sleep(inter)
        else:
            while self.colorMatch(x, y, rgb, tol):
                time.sleep(inter)

        time.sleep(delay)

    # ---------------- #
    #      MOUSE       #
    # ---------------- #

    def randClick(self, x=0, y=0, right=False, tol=10, delay=0.3):
        """
        Random click.

        Args:
            x, y : Position x, y
            right: Right button.
            tol  : Range    [-tol ~ +tol]
        """

        # Mouse Click https://stackoverflow.com/questions/37622438/emulating-a-mouseclick-using-postmessage

        _btn_dn = wc.WM_RBUTTONDOWN if right else wc.WM_LBUTTONDOWN
        _btn_up = wc.WM_RBUTTONUP   if right else wc.WM_LBUTTONUP
        _btn_mk = wc.MK_RBUTTON     if right else wc.MK_LBUTTON
        _x, _y  = self._offset(x, y)
        _rx     = _x + random.randint(-tol, tol)
        _ry     = _y + random.randint(-tol, tol)
        _lParam = wa.MAKELONG(_rx, _ry)

        wg.PostMessage(self._hwnd, _btn_dn, _btn_mk, _lParam)
        wg.PostMessage(self._hwnd, _btn_up,       0, _lParam)

        time.sleep(delay)

    def randDrag(self, fx=0, fy=0, tx=0, ty=0, right=False, step=20, tol=10, inter=0.01, delay=0.3):
        """
        Random Drag.

        Args:
            fx, fy: Position From  x, y
            tx, ty: Position To    x, y
            right : Right button.
            step  : Divide move points.
            tol   : Range          [-tol ~ +tol]
            inter : Sleep to interval each step.
        """

        _btn_dn   = wc.WM_RBUTTONDOWN if right else wc.WM_LBUTTONDOWN
        _btn_up   = wc.WM_RBUTTONUP   if right else wc.WM_LBUTTONUP
        _btn_mk   = wc.MK_RBUTTON     if right else wc.MK_LBUTTON
        _fx, _fy  = self._offset(fx, fy)
        _tx, _ty  = self._offset(tx, ty)
        _rfx      = _fx + random.randint(-tol, tol)
        _rfy      = _fy + random.randint(-tol, tol)
        _rtx      = _tx + random.randint(-tol, tol)
        _rty      = _ty + random.randint(-tol, tol)
        _f_lParam = wa.MAKELONG(_rfx, _rfy)
        _t_lParam = wa.MAKELONG(_rtx, _rty)

        wg.PostMessage(self._hwnd, _btn_dn, _btn_mk, _f_lParam)

        for mid in range(step):
            _mx = ((_rtx - _rfx) // step) * mid
            _my = ((_rty - _rfy) // step) * mid
            _m_lParam = wa.MAKELONG(_mx, _my)

            wg.PostMessage(self._hwnd, wc.WM_MOUSEMOVE, 0, _m_lParam)

            time.sleep(inter)

        wg.PostMessage(self._hwnd, _btn_up,       0, _t_lParam)

        time.sleep(delay)

    # ---------------- #
    #     KEYBOARD     #
    # ---------------- #

    def sendKey(self, key=0x16, inter=0.01, delay=0.03):
        """
        Input Keyboard.

        Refer to microsoft docs: virtual-key-codes.
        """

        # Virtual Key https://docs.microsoft.com/ko-kr/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN

        wg.PostMessage(self._hwnd, wc.WM_KEYDOWN, key, 0)
        time.sleep(inter)
        wg.PostMessage(self._hwnd, wc.WM_KEYUP  , key, 0)

        time.sleep(delay)

    def sendKeyEx(self, key, mods, special, inter=0.01, delay=0.03):
        """
        Input Keyboard with modifier keys. (e.g. CTRL, SHIFT, ALT)

        Exsample, 
            sendKeyEx(key=ord('S'), mods='ctrl', special=False)
            sendKeyEx(key=ord('M'), mods='ctrl shift alt', special=False)
        """

        # https://dev-qa.com/1321809/how-to-transfer-keyboard-shortcuts-ctrl-etc-inactive-window

        modkeys = []
        modkeys.append(wc.VK_CONTROL) if mods.count('ctrl')  else None
        modkeys.append(wc.VK_SHIFT)   if mods.count('shift') else None
        modkeys.append(wc.VK_MENU)    if mods.count('alt')   else None

        _user32                  = ctypes.WinDLL("user32")
        PBYTE256                 = ctypes.c_ubyte * 256
        GetKeyboardState         = _user32.GetKeyboardState
        SetKeyboardState         = _user32.SetKeyboardState
        GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
        AttachThreadInput        = _user32.AttachThreadInput
        MapVirtualKeyA           = _user32.MapVirtualKeyA
        #MapVirtualKeyW           = _user32.MapVirtualKeyW

        if wg.IsWindow(self._hwnd):
            ThreadId = GetWindowThreadProcessId(self._hwnd, None)
            lparam   = wa.MAKELONG(0, MapVirtualKeyA(key, 0))
            msg_down = wc.WM_KEYDOWN
            msg_up   = wc.WM_KEYUP

            if special:
                lparam |= 0x1000000

            if len(modkeys) > 0:
                pKeyBuffers     = PBYTE256()
                pKeyBuffers_old = PBYTE256()

                wg.SendMessage(self._hwnd, wc.WM_ACTIVATE, wc.WA_ACTIVE, 0)
                AttachThreadInput(wa.GetCurrentThreadId(), ThreadId, True)
                GetKeyboardState(ctypes.byref(pKeyBuffers_old))

                for modkey in modkeys:
                    if modkey == wc.VK_MENU:
                        lparam   |= 0x20000000
                        msg_down  = wc.WM_SYSKEYDOWN
                        msg_up    = wc.WM_SYSKEYUP

                    pKeyBuffers[modkey] |= 128

                SetKeyboardState(ctypes.byref(pKeyBuffers))
                time.sleep(inter)
                wa.PostMessage(self._hwnd, msg_down, key, lparam)
                time.sleep(inter)
                wa.PostMessage(self._hwnd, msg_up  , key, lparam | 0xC0000000)
                time.sleep(inter)
                SetKeyboardState(ctypes.byref(pKeyBuffers_old))
                time.sleep(inter)
                AttachThreadInput(wa.GetCurrentThreadId(), ThreadId, False)
            else:
                wg.SendMessage(self._hwnd, msg_down, key, lparam)
                wg.SendMessage(self._hwnd, msg_up  , key, lparam | 0xC0000000)

        time.sleep(delay)

    # ---------------- #
    #      IMAGE       #
    # ---------------- #

    def screenshot(self, path, x=0, y=0, w=0, h=0, delay=0):
        """
        Save Screenshot to path.

        If (x, y) is (0, 0), it will be current zero position.
        If (w, h) is (0, 0), it will be full size of window.

        Args:
            path: Image path to save.
            x, y: Position x, y
            w, h: Size     width, height
        """

        # https://stackoverflow.com/questions/19695214/screenshot-of-inactive-window-printwindow-win32gui

        _x, _y  = (abs(x), abs(y)) if x and y else self._zero
        _w, _h  = (_x + w, _y + h) if w and h else self._size
        _mfc_DC = wu.CreateDCFromHandle(wg.GetWindowDC(self._hwnd))
        _mem_DC = _mfc_DC.CreateCompatibleDC()
        _bitmap = wu.CreateBitmap()

        _bitmap.CreateCompatibleBitmap(_mfc_DC, self._size[0], self._size[1])
        _mem_DC.SelectObject(_bitmap)

        _chk_succ = ctypes.windll.user32.PrintWindow(self._hwnd, _mem_DC.GetSafeHdc(), 2)
        _bmp_info = _bitmap.GetInfo()
        _bmp_bits = _bitmap.GetBitmapBits(True)
        _image    = Image.frombuffer('RGB', (_bmp_info['bmWidth'], _bmp_info['bmHeight']), _bmp_bits, 'raw', 'BGRX', 0, 1)
        _image    = _image.crop((_x, _y, _w, _h))

        wg.DeleteObject(_bitmap.GetHandle())
        _mem_DC.DeleteDC()
        _mfc_DC.DeleteDC()

        try:
            if _chk_succ == 1: _image.save(path)
        except ValueError:
            print('ERROR: Wrong path.')

        time.sleep(delay)

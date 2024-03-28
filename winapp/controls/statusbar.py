# https://learn.microsoft.com/en-us/windows/win32/controls/status-bar-reference

from ctypes import byref, pointer, create_unicode_buffer
from ctypes.wintypes import INT, RECT
from winapp.const import WM_USER, WS_CHILD, WS_VISIBLE, BF_MONO, BF_FLAT, BDR_OUTER
from winapp.window import *
from winapp.dlls import user32
from winapp.themes import *

########################################
# Class Name
########################################
STATUSCLASSNAME = WC_STATUSBAR = 'msctls_statusbar32'

SB_SETPARTS = WM_USER + 4
SB_SETTEXTA = WM_USER + 1
SB_SETTEXTW = WM_USER + 11
SB_GETRECT = WM_USER + 10
SB_GETBORDERS = WM_USER + 7

SBARS_SIZEGRIP = 0x100
SBARS_TOOLTIPS = 0x0800

# this is a status bar flag, preference to SBARS_TOOLTIPS
SBT_TOOLTIPS = 0x0800

SB_SETTEXTW             =(WM_USER+11)
SB_GETTEXTLENGTHW       =(WM_USER+12)
SB_GETTEXTW            = (WM_USER+13)
SB_SETPARTS             =(WM_USER+4)
SB_GETPARTS             =(WM_USER+6)
SB_GETBORDERS           =(WM_USER+7)
SB_SETMINHEIGHT         =(WM_USER+8)
SB_SIMPLE               =(WM_USER+9)
SB_GETRECT              =(WM_USER+10)
SB_ISSIMPLE             =(WM_USER+14)
SB_SETICON              =(WM_USER+15)
SB_SETTIPTEXTW          =(WM_USER+17)
SB_GETTIPTEXTW          =(WM_USER+19)
SB_GETICON              =(WM_USER+20)

SBT_OWNERDRAW            =0x1000
SBT_NOBORDERS            =0x0100
SBT_POPOUT               =0x0200
SBT_RTLREADING           =0x0400
SBT_NOTABPARSING         =0x0800

# custom (for SB_GETBORDERS)
class BORDERINFO(Structure):
    _fields_ = [
        ("horizontal", INT),
        ("vertical", INT),
        ("between", INT),
    ]

SEPARATOR_BRUSH_DARK = gdi32.CreateSolidBrush(0x646464)


class StatusBar(Window):

    def __init__(self,
            parent_window=None,
            style=WS_CHILD | WS_VISIBLE,
            ex_style=0, window_title=0,
            parts=[],
            parts_right_aligned=False):

        self.parts = parts

        num_parts = len(parts) if parts else 1
        if num_parts < 2:
            parts_right_aligned = False

        super().__init__(
            STATUSCLASSNAME,
            parent_window=parent_window,
            style=style,
            ex_style=ex_style,
            window_title=window_title
            )

        if num_parts > 1:
            rc = RECT()
            user32.GetWindowRect(parent_window.hwnd, byref(rc))
            w = rc.right - rc.left
            sb_parts = (INT * num_parts)()

            if parts_right_aligned:
                for i in range(num_parts - 1):
                    sb_parts[i] = w - sum(parts[i+1:])
            else:
                # fixed position (left aligned)
                for i in range(num_parts - 1):
                    sb_parts[i] = parts[i]

            sb_parts[num_parts - 1] = -1
            user32.SendMessageW(self.hwnd, SB_SETPARTS, num_parts, sb_parts)

        # get height of statusbar
        rc = RECT()
        user32.SendMessageW(self.hwnd, SB_GETRECT, 0, byref(rc))
        self.height = rc.bottom

    def set_text(self, msg='', part=0):
        user32.SendMessageW(self.hwnd, SB_SETTEXTW, part, msg)

    def set_icon(self, hicon, part=0):
        user32.SendMessageW(self.hwnd, SB_SETICON, part, hicon)

    def right_align_parts(self, width):
        status_parts_count = len(self.parts)
        sb_parts = (INT * status_parts_count)()
        for i in range(status_parts_count - 1):
            sb_parts[i] = width - sum(self.parts[i + 1:])
        sb_parts[status_parts_count - 1] = -1
        user32.SendMessageW(self.hwnd, SB_SETPARTS, status_parts_count, sb_parts)

    def update_size(self, *args):
        user32.SendMessageW(self.hwnd, WM_SIZE, 0, 0)

    def apply_theme(self, is_dark):
        super().apply_theme(is_dark)

        if is_dark:
            hfont = user32.SendMessageW(self.hwnd, WM_GETFONT, 0, 0)
            rc_part = RECT()

            def _on_WM_PAINT(hwnd, wparam, lparam):
                ps = PAINTSTRUCT()
                hdc = user32.BeginPaint(self.hwnd, byref(ps))

                ps.rcPaint.right -= 1
                user32.FillRect(hdc, byref(ps.rcPaint), BG_BRUSH_DARK)

                gdi32.SelectObject(hdc, hfont)
                gdi32.SetBkMode(hdc, TRANSPARENT)
                gdi32.SetTextColor(hdc, TEXT_COLOR_DARK)

                for i in range(len(self.parts)):
                    user32.SendMessageW(self.hwnd, SB_GETRECT, i, byref(rc_part))
                    if ps.rcPaint.left >= rc_part.right or ps.rcPaint.right < rc_part.left:
                        continue

                    # Draw text
                    text_len = user32.SendMessageW(self.hwnd, SB_GETTEXTLENGTHW, i, 0) + 1
                    buf = create_unicode_buffer(text_len)
                    user32.SendMessageW(self.hwnd, SB_GETTEXTW, i, buf)
                    user32.DrawTextW(hdc, buf.value, text_len,
                            byref(RECT(rc_part.left + 2, rc_part.top + 1, rc_part.right, rc_part.bottom - 1)),
                            DT_SINGLELINE | DT_VCENTER | DT_LEFT)

                    # Draw separator
                    if i < len(self.parts) - 1:
                        user32.FillRect(hdc, byref(RECT(rc_part.right - 1, rc_part.top + 1, rc_part.right, rc_part.bottom - 3)),
                                SEPARATOR_BRUSH_DARK)

                user32.EndPaint(self.hwnd, byref(ps))
                return 0

            def _on_WM_ERASEBKGND(hwnd, wparam, lparam):
                return 0

            self.register_message_callback(WM_PAINT, _on_WM_PAINT)
            self.register_message_callback(WM_ERASEBKGND, _on_WM_ERASEBKGND)
        else:
            self.unregister_message_callback(WM_PAINT)
            self.unregister_message_callback(WM_ERASEBKGND)

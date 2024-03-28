# https://learn.microsoft.com/en-us/windows/win32/controls/static-controls

from winapp.const import WS_CHILD, WS_VISIBLE
from winapp.window import *
from winapp.dlls import user32
from winapp.themes import *

########################################
# Class Name
########################################
STATIC_CLASS = WC_STATIC = 'Static'

########################################
# Static Control Constants
########################################

SS_LEFT             =0x00000000
SS_CENTER           =0x00000001
SS_RIGHT            =0x00000002
SS_ICON             =0x00000003
SS_BLACKRECT        =0x00000004
SS_GRAYRECT         =0x00000005
SS_WHITERECT        =0x00000006
SS_BLACKFRAME       =0x00000007
SS_GRAYFRAME        =0x00000008
SS_WHITEFRAME       =0x00000009
SS_USERITEM         =0x0000000A
SS_SIMPLE           =0x0000000B
SS_LEFTNOWORDWRAP   =0x0000000C
SS_OWNERDRAW        =0x0000000D
SS_BITMAP           =0x0000000E
SS_ENHMETAFILE      =0x0000000F
SS_ETCHEDHORZ       =0x00000010
SS_ETCHEDVERT       =0x00000011
SS_ETCHEDFRAME      =0x00000012
SS_TYPEMASK         =0x0000001F
SS_REALSIZECONTROL  =0x00000040
SS_NOPREFIX         =0x00000080  # Don't do "&" character translation
SS_NOTIFY           =0x00000100
SS_CENTERIMAGE      =0x00000200
SS_RIGHTJUST        =0x00000400
SS_REALSIZEIMAGE    =0x00000800
SS_SUNKEN           =0x00001000
SS_EDITCONTROL      =0x00002000
SS_ENDELLIPSIS      =0x00004000
SS_PATHELLIPSIS     =0x00008000
SS_WORDELLIPSIS     =0x0000C000
SS_ELLIPSISMASK     =0x0000C000

########################################
# Static Control Mesages
########################################
STM_SETICON         =0x0170
STM_GETICON         =0x0171
STM_SETIMAGE        =0x0172
STM_GETIMAGE        =0x0173
STN_CLICKED         =0
STN_DBLCLK          =1
STN_ENABLE          =2
STN_DISABLE         =3
STM_MSGMAX          =0x0174


########################################
# Wrapper Class
########################################
class Static(Window):

    ########################################
    #
    ########################################
    def __init__(self, parent_window=None, style=WS_CHILD | WS_VISIBLE, ex_style=0,
            left=0, top=0, width=0, height=0, window_title=0,
            bg_color=COLOR_WINDOW + 1,
            bg_color_dark=BG_COLOR_DARK
            ):

        self.bg_color = bg_color
        self.bg_color_dark = bg_color_dark

        super().__init__(
            WC_STATIC,
            parent_window=parent_window,
            style=style,
            ex_style=ex_style,
            left=left,
            top=top,
            width=width,
            height=height,
            window_title=window_title,
            )

    ########################################
    #
    ########################################
    def destroy_window(self):
        if self.is_dark:
            self.parent_window.unregister_message_callback(WM_CTLCOLORSTATIC, self._on_WM_CTLCOLORSTATIC)
        super().destroy_window()

    ########################################
    #
    ########################################
    def _on_WM_CTLCOLORSTATIC(self, hwnd, wparam, lparam):
        if lparam == self.hwnd:
            gdi32.SetTextColor(wparam, TEXT_COLOR_DARK)
            gdi32.SetBkColor(wparam, BG_COLOR_DARK)
            gdi32.SetDCBrushColor(wparam,  BG_COLOR_DARK)
            return gdi32.GetStockObject(DC_BRUSH)

    ########################################
    #
    ########################################
    def set_image(self, hbitmap):
        user32.SendMessageW(self.hwnd, STM_SETIMAGE, IMAGE_BITMAP, hbitmap)

    ########################################
    #
    ########################################
    def set_icon(self, hicon):
        user32.SendMessageW(self.hwnd, STM_SETICON, hicon, 0)

    ########################################
    #
    ########################################
    def apply_theme(self, is_dark):
        self.is_dark = is_dark
        if is_dark:
            self.parent_window.register_message_callback(WM_CTLCOLORSTATIC, self._on_WM_CTLCOLORSTATIC)
        else:
            self.parent_window.unregister_message_callback(WM_CTLCOLORSTATIC, self._on_WM_CTLCOLORSTATIC)
        self.force_redraw_window()  # triggers WM_CTLCOLORSTATIC

# https://learn.microsoft.com/en-us/windows/win32/controls/buttons

from ctypes import *
from ctypes.wintypes import *
from winapp.wintypes_extended import MAKELPARAM
from winapp.window import *
from winapp.dlls import user32
from winapp.controls.static import *

########################################
# Class Name
########################################
BUTTON_CLASS = WC_BUTTON = 'Button'

########################################
# Mesages
########################################
BM_GETCHECK        =0x00F0
BM_SETCHECK        =0x00F1
BM_GETSTATE        =0x00F2
BM_SETSTATE        =0x00F3
BM_SETSTYLE        =0x00F4
BM_CLICK           =0x00F5
BM_GETIMAGE        =0x00F6
BM_SETIMAGE        =0x00F7
BM_SETDONTCLICK    =0x00F8

BCM_FIRST                =0x1600
BCM_GETIDEALSIZE         =(BCM_FIRST + 0x0001)
BCM_SETIMAGELIST         =(BCM_FIRST + 0x0002)
BCM_GETIMAGELIST         =(BCM_FIRST + 0x0003)
BCM_SETTEXTMARGIN        =(BCM_FIRST + 0x0004)
BCM_GETTEXTMARGIN        =(BCM_FIRST + 0x0005)
BCM_SETDROPDOWNSTATE     =(BCM_FIRST + 0x0006)
BCM_SETSPLITINFO         =(BCM_FIRST + 0x0007)
BCM_GETSPLITINFO         =(BCM_FIRST + 0x0008)
BCM_SETNOTE              =(BCM_FIRST + 0x0009)
BCM_GETNOTE              =(BCM_FIRST + 0x000A)
BCM_GETNOTELENGTH        =(BCM_FIRST + 0x000B)
BCM_SETSHIELD            =(BCM_FIRST + 0x000C)

########################################
# Notifications (BN_... moved to common.py)
########################################
#BN_CLICKED          =0
#BN_PAINT            =1
#BN_HILITE           =2
#BN_UNHILITE         =3
#BN_DISABLE          =4
#BN_DOUBLECLICKED    =5
#BN_PUSHED           =BN_HILITE
#BN_UNPUSHED         =BN_UNHILITE
#BN_DBLCLK           =BN_DOUBLECLICKED
#BN_SETFOCUS         =6
#BN_KILLFOCUS        =7

#define BCN_FIRST               (0U-1250U)
#define BCN_LAST                (0U-1350U)
BCN_FIRST = (-1250)
BCN_HOTITEMCHANGE       =(BCN_FIRST + 0x0001)
BCN_DROPDOWN            =(BCN_FIRST + 0x0002)

########################################
# Constants / Styles
########################################
BS_PUSHBUTTON       =0x00000000
BS_DEFPUSHBUTTON    =0x00000001
BS_CHECKBOX         =0x00000002
BS_AUTOCHECKBOX     =0x00000003
BS_RADIOBUTTON      =0x00000004
BS_3STATE           =0x00000005
BS_AUTO3STATE       =0x00000006
BS_GROUPBOX         =0x00000007
BS_USERBUTTON       =0x00000008
BS_AUTORADIOBUTTON  =0x00000009
BS_PUSHBOX          =0x0000000A
BS_OWNERDRAW        =0x0000000B
BS_TYPEMASK         =0x0000000F
BS_LEFTTEXT         =0x00000020
BS_TEXT             =0x00000000
BS_ICON             =0x00000040
BS_BITMAP           =0x00000080
BS_LEFT             =0x00000100
BS_RIGHT            =0x00000200
BS_CENTER           =0x00000300
BS_TOP              =0x00000400
BS_BOTTOM           =0x00000800
BS_VCENTER          =0x00000C00
BS_PUSHLIKE         =0x00001000
BS_MULTILINE        =0x00002000
BS_NOTIFY           =0x00004000
BS_FLAT             =0x00008000
BS_RIGHTBUTTON      =BS_LEFTTEXT

BUTTON_IMAGELIST_ALIGN_LEFT     =0
BUTTON_IMAGELIST_ALIGN_RIGHT    =1
BUTTON_IMAGELIST_ALIGN_TOP      =2
BUTTON_IMAGELIST_ALIGN_BOTTOM   =3
BUTTON_IMAGELIST_ALIGN_CENTER   =4

BCCL_NOGLYPH = -1

########################################
# Macros
########################################
Button_GetIdealSize = lambda hwnd, psize: user32.SendMessageW(hwnd, BCM_GETIDEALSIZE, 0, psize)
Button_SetImageList = lambda hwnd, pbuttonImagelist: user32.SendMessageW(hwnd, BCM_SETIMAGELIST, 0, pbuttonImagelist)
Button_GetImageList = lambda hwnd, pbuttonImagelist: user32.SendMessageW(hwnd, BCM_GETIMAGELIST, 0, pbuttonImagelist)
Button_SetTextMargin = lambda hwnd, pmargin: user32.SendMessageW(hwnd, BCM_SETTEXTMARGIN, 0, pmargin)
Button_GetTextMargin = lambda hwnd, pmargin: user32.SendMessageW(hwnd, BCM_GETTEXTMARGIN, 0, pmargin)

Button_SetDropDownState = lambda hwnd, fDropDown: user32.SendMessageW(hwnd, BCM_SETDROPDOWNSTATE, fDropDown, 0)
Button_SetSplitInfo = lambda hwnd, pInfo: user32.SendMessageW(hwnd, BCM_SETSPLITINFO, 0, pInfo)
Button_GetSplitInfo = lambda hwnd, pInfo: user32.SendMessageW(hwnd, BCM_GETSPLITINFO, 0, pInfo)
Button_SetNote = lambda hwnd, psz: user32.SendMessageW(hwnd, BCM_SETNOTE, 0, psz)
Button_GetNote = lambda hwnd, psz, pcc: user32.SendMessageW(hwnd, BCM_GETNOTE, pcc, psz)
Button_GetNoteLength = lambda hwnd: user32.SendMessageW(hwnd, BCM_GETNOTELENGTH, 0, 0)
Button_SetElevationRequiredStat = lambda hwnd, fRequired: user32.SendMessageW(hwnd, BCM_SETSHIELD, 0, fRequired)

Button_Enable = lambda hwndCtl, fEnable:         user32.EnableWindow(hwndCtl, fEnable)
Button_GetText = lambda hwndCtl, lpch, cchMax:   user32.GetWindowText(hwndCtl, lpch, cchMax)
Button_GetTextLength = lambda hwndCtl:           user32.GetWindowTextLength(hwndCtl)
Button_SetText = lambda hwndCtl, lpsz:           user32.SetWindowText(hwndCtl, lpsz)

BST_UNCHECKED      =0x0000
BST_CHECKED        =0x0001
BST_INDETERMINATE  =0x0002
BST_PUSHED         =0x0004
BST_FOCUS          =0x0008

Button_GetCheck = lambda hwndCtl:            user32.SendMessageW(hwndCtl, BM_GETCHECK, 0, 0)
Button_SetCheck = lambda hwndCtl, check:     user32.SendMessageW(hwndCtl, BM_SETCHECK, check, 0)
Button_GetState = lambda hwndCtl:            user32.SendMessageW(hwndCtl, BM_GETSTATE, 0, 0)
Button_SetState = lambda hwndCtl, state:     user32.SendMessageW(hwndCtl, BM_SETSTATE, state, 0)
Button_SetStyle = lambda hwndCtl, style, fRedraw: user32.SendMessageW(hwndCtl, BM_SETSTYLE, style, MAKELPARAM(fRedraw, 0))


########################################
# Button Control Structures
########################################
class BUTTON_IMAGELIST(Structure):
    _fields_ = [
        ("himl", HANDLE),
        ("margin", RECT),
        ("uAlign", UINT),
    ]


########################################
# Wrapper Class
########################################
class Button(Window):

    ########################################
    #
    ########################################
    def __init__(self, parent_window, style=WS_CHILD | WS_VISIBLE, ex_style=0,
            left=0, top=0, width=94, height=23, window_title='OK', wrap_hwnd=None):

        super().__init__(
            WC_BUTTON,
            parent_window=parent_window,
            style=style,
            ex_style=ex_style,
            left=left,
            top=top,
            width=width,
            height=height,
            window_title=window_title,
            wrap_hwnd=wrap_hwnd
            )

    ########################################
    #
    ########################################
    def destroy_window(self):
        if self.is_dark:
            self.parent_window.unregister_message_callback(WM_CTLCOLORBTN, self._on_WM_CTLCOLORBTN)
        super().destroy_window()

    ########################################
    #
    ########################################
    def apply_theme(self, is_dark):
        super().apply_theme(is_dark)
        uxtheme.SetWindowTheme(self.hwnd, 'DarkMode_Explorer' if is_dark else 'Explorer', None)
        if is_dark:
            self.parent_window.register_message_callback(WM_CTLCOLORBTN, self._on_WM_CTLCOLORBTN)
        else:
            self.parent_window.unregister_message_callback(WM_CTLCOLORBTN, self._on_WM_CTLCOLORBTN)

    ########################################
    #
    ########################################
    def _on_WM_CTLCOLORBTN(self, hwnd, wparam, lparam):
        if lparam == self.hwnd:
            gdi32.SetDCBrushColor(wparam, BG_COLOR_DARK)
            return gdi32.GetStockObject(DC_BRUSH)


########################################
# Wrapper Class
########################################
class CheckBox(Window):

    ########################################
    #
    ########################################
    def __init__(self, parent_window, style=WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, ex_style=0,
            left=0, top=0, width=0, height=0, window_title=0):

        super().__init__(
            WC_BUTTON,
            parent_window=parent_window,
            style=style,
            ex_style=ex_style,
            left=left,
            top=top,
            width=width,
            height=height,
            )

        self.checkbox_static = Static(
            self,
            style=WS_CHILD | SS_SIMPLE | WS_VISIBLE,
            ex_style=WS_EX_TRANSPARENT,
            left=16,
            top=1,
            width=width - 16,
            height=height,
            window_title=window_title.replace('&', '')
        )

    ########################################
    #
    ########################################
    def set_font(self, **kwargs):
        self.checkbox_static.set_font(**kwargs)

    ########################################
    #
    ########################################
    def apply_theme(self, is_dark):
        super().apply_theme(is_dark)
        uxtheme.SetWindowTheme(self.hwnd, 'DarkMode_Explorer' if is_dark else 'Explorer', None)

# https://learn.microsoft.com/en-us/windows/win32/controls/combo-boxes

#from ctypes import *
from ctypes.wintypes import *
from winapp.const import WS_CHILD, WS_VISIBLE
#from winapp.wintypes_extended import MAKELONG
from winapp.window import *
from winapp.dlls import user32


########################################
# Class Name
########################################
COMBOBOX_CLASS = WC_COMBOBOX = 'ComboBox'

########################################
# Combo Box return Values
########################################
CB_OKAY             =0
CB_ERR              =(-1)
CB_ERRSPACE         =(-2)

########################################
# Combo Box Notification Codes
########################################
CBN_ERRSPACE        =(-1)
CBN_SELCHANGE       =1
CBN_DBLCLK          =2
CBN_SETFOCUS        =3
CBN_KILLFOCUS       =4
CBN_EDITCHANGE      =5
CBN_EDITUPDATE      =6
CBN_DROPDOWN        =7
CBN_CLOSEUP         =8
CBN_SELENDOK        =9
CBN_SELENDCANCEL    =10

########################################
# Combo Box styles
########################################
CBS_SIMPLE            =0x0001
CBS_DROPDOWN          =0x0002
CBS_DROPDOWNLIST      =0x0003
CBS_OWNERDRAWFIXED    =0x0010
CBS_OWNERDRAWVARIABLE =0x0020
CBS_AUTOHSCROLL       =0x0040
CBS_OEMCONVERT        =0x0080
CBS_SORT              =0x0100
CBS_HASSTRINGS        =0x0200
CBS_NOINTEGRALHEIGHT  =0x0400
CBS_DISABLENOSCROLL   =0x0800
#if(WINVER >= 0x0400)
#define CBS_UPPERCASE         0x2000L
#define CBS_LOWERCASE         0x4000L
#endif /* WINVER >= 0x0400 */

########################################
# Mesages
########################################
CB_GETEDITSEL               =0x0140
CB_LIMITTEXT                =0x0141
CB_SETEDITSEL               =0x0142
CB_ADDSTRING                =0x0143
CB_DELETESTRING             =0x0144
CB_DIR                      =0x0145
CB_GETCOUNT                 =0x0146
CB_GETCURSEL                =0x0147
CB_GETLBTEXT                =0x0148
CB_GETLBTEXTLEN             =0x0149
CB_INSERTSTRING             =0x014A
CB_RESETCONTENT             =0x014B
CB_FINDSTRING               =0x014C
CB_SELECTSTRING             =0x014D
CB_SETCURSEL                =0x014E
CB_SHOWDROPDOWN             =0x014F
CB_GETITEMDATA              =0x0150
CB_SETITEMDATA              =0x0151
CB_GETDROPPEDCONTROLRECT    =0x0152
CB_SETITEMHEIGHT            =0x0153
CB_GETITEMHEIGHT            =0x0154
CB_SETEXTENDEDUI            =0x0155
CB_GETEXTENDEDUI            =0x0156
CB_GETDROPPEDSTATE          =0x0157
CB_FINDSTRINGEXACT          =0x0158
CB_SETLOCALE                =0x0159
CB_GETLOCALE                =0x015A
CB_GETTOPINDEX              =0x015b
CB_SETTOPINDEX              =0x015c
CB_GETHORIZONTALEXTENT      =0x015d
CB_SETHORIZONTALEXTENT      =0x015e
CB_GETDROPPEDWIDTH          =0x015f
CB_SETDROPPEDWIDTH          =0x0160
CB_INITSTORAGE              =0x0161
CB_GETCOMBOBOXINFO   =0x0164

#if(_WIN32_WINNT >= 0x0501)
#define CB_MSGMAX                   0x0165
#elif defined(_WIN32_WCE) && (_WIN32_WCE >= 0x0400)
#define CB_MSGMAX                   0x0163
#elif(WINVER >= 0x0400)
#define CB_MSGMAX                   0x0162
#else
#define CB_MSGMAX                   0x015B
#endif


class COMBOBOXINFO(Structure):
    def __init__(self, *args, **kwargs):
        super(COMBOBOXINFO, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)
    _fields_ = [
        ("cbSize", DWORD),
        ("rcItem",  RECT),
        ("rcButton", RECT),
        ("stateButton",  DWORD),
        ("hwndCombo",  HWND),
        ("hwndItem",  HWND),
        ("hwndList",  HWND),
    ]


########################################
# Wrapper Class
########################################
class ComboBox(Window):

    ########################################
    #
    ########################################
    def __init__(self, parent_window, style=WS_CHILD | WS_VISIBLE, ex_style=0,
            left=0, top=0, width=0, height=0, window_title=0, wrap_hwnd=None):

        super().__init__(
            WC_COMBOBOX,
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

        self.__has_edit = style & CBS_DROPDOWN

    ########################################
    #
    ########################################
    def destroy_window(self):
        if self.is_dark:
            self.unregister_message_callback(WM_CTLCOLORLISTBOX,self.on_WM_CTLCOLORLISTBOX)
            if self.__has_edit:
                self.unregister_message_callback(WM_CTLCOLOREDIT, self._on_WM_CTLCOLOREDIT)
        super().destroy_window()

    def add_string(self, s):
        user32.SendMessageW(self.hwnd, CB_ADDSTRING, 0, s)

    def set_current_selection(self, idx):
        user32.SendMessageW(self.hwnd, CB_SETCURSEL, idx, 0)

    def get_current_selection(self):
        return user32.SendMessageW(self.hwnd, CB_GETCURSEL, 0, 0)

    ########################################
    #
    ########################################
    def apply_theme(self, is_dark):
        #print('ComboBox apply_theme')
        self.is_dark = is_dark

        uxtheme.SetWindowTheme(self.hwnd, 'DarkMode_CFD' if is_dark else 'CFD', None)

        # scrollbar colors
        ci = COMBOBOXINFO()
        user32.SendMessageW(self.hwnd, CB_GETCOMBOBOXINFO, 0, byref(ci))
        uxtheme.SetWindowTheme(ci.hwndList, 'DarkMode_Explorer' if is_dark else 'Explorer', None)

        if is_dark:
            self.register_message_callback(WM_CTLCOLORLISTBOX, self.on_WM_CTLCOLORLISTBOX)
            if self.__has_edit:
                self.register_message_callback(WM_CTLCOLOREDIT, self._on_WM_CTLCOLOREDIT)
        else:
            self.unregister_message_callback(WM_CTLCOLORLISTBOX,self.on_WM_CTLCOLORLISTBOX)
            if self.__has_edit:
                self.unregister_message_callback(WM_CTLCOLOREDIT, self._on_WM_CTLCOLOREDIT)

    ########################################
    #
    ########################################
    def on_WM_CTLCOLORLISTBOX(self, hwnd, wparam, lparam):
        gdi32.SetTextColor(wparam, TEXT_COLOR_DARK)
        gdi32.SetBkColor(wparam, CONTROL_COLOR_DARK)
        gdi32.SetDCBrushColor(wparam, CONTROL_COLOR_DARK)
        return gdi32.GetStockObject(DC_BRUSH)

    ########################################
    #
    ########################################
    def _on_WM_CTLCOLOREDIT(self, hwnd, wparam, lparam):
        gdi32.SetTextColor(wparam, TEXT_COLOR_DARK)
        gdi32.SetBkColor(wparam, CONTROL_COLOR_DARK)
        gdi32.SetDCBrushColor(wparam, CONTROL_COLOR_DARK)
        return gdi32.GetStockObject(DC_BRUSH)

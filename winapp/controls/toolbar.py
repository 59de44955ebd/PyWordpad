# https://learn.microsoft.com/en-us/windows/win32/controls/toolbar-control-reference

from ctypes import Structure, c_wchar_p, sizeof, byref, cast
from ctypes.wintypes import *

from winapp.wintypes_extended import DWORD_PTR, UINT_PTR, MAKELONG
from winapp.const import WM_USER, WS_EX_COMPOSITED, WS_CHILD, WS_VISIBLE, WS_CLIPCHILDREN, IMAGE_BITMAP, LR_LOADFROMFILE, LR_LOADTRANSPARENT, LR_LOADMAP3DCOLORS
from winapp.controls.common import *
from winapp.window import *
from winapp.dlls import user32
from winapp.themes import *

########################################
# Class Name
########################################
TOOLBARCLASSNAME = WC_TOOLBAR = 'ToolbarWindow32'

TBBUTTON_RESERVED_SIZE = 6

class TBBUTTON(Structure):
    _fields_ = [
        ("iBitmap", INT),
        ("idCommand", INT),
        ("fsState", BYTE),
        ("fsStyle", BYTE),
        ("bReserved", BYTE * TBBUTTON_RESERVED_SIZE),
        ("dwData", DWORD_PTR),
        ("iString", c_wchar_p)
    ]

class TBADDBITMAP(Structure):
    _fields_ = [
        ("hInst", HINSTANCE),
        ("nID", UINT_PTR),
    ]

class TBBUTTONINFOW(Structure):
    def __init__(self, *args, **kwargs):
        super(TBBUTTONINFOW, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)
    _fields_ = [
        ("cbSize", UINT),
        ("dwMask", DWORD),
        ("idCommand", INT),
        ("iImage", INT),
        ("fsState", BYTE),
        ("fsStyle", BYTE),
        ("cx", WORD),
        ("lParam", DWORD_PTR),
        ("pszText", LPWSTR),
        ("cchText", INT)
    ]

class NMTBCUSTOMDRAW(Structure):
    _fields_ = [
        ("nmcd", NMCUSTOMDRAW),
        ("hbrMonoDither", HBRUSH),
        ("hbrLines", HBRUSH),
        ("hpenLines", HPEN),
        ("clrText", COLORREF),
        ("clrMark", COLORREF),
        ("clrTextHighlight", COLORREF),
        ("clrBtnFace", COLORREF),
        ("clrBtnHighlight", COLORREF),
        ("clrHighlightHotTrack", COLORREF),
        ("rcText", RECT),
        ("nStringBkMode", INT),
        ("nHLStringBkMode", INT),
        ("iListGap", INT),
    ]
LPNMTBCUSTOMDRAW = POINTER(NMTBCUSTOMDRAW)

class TBREPLACEBITMAP(Structure):
    _fields_ = [
        ("hInstOld", HINSTANCE),
        ("nIDOld", UINT_PTR),
        ("hInstNew", HINSTANCE),
        ("nIDNew", UINT_PTR),
        ("nButtons", INT),
    ]

class NMTOOLBARW(Structure):
    _fields_ = [
        ("hdr", NMHDR),
        ("iItem", INT),
        ("tbButton", TBBUTTON),
        ("cchText", INT),
        ("pszText", LPWSTR),
        ("rcButton", RECT),
    ]
LPNMTOOLBARW = POINTER(NMTOOLBARW)


CCM_FIRST =0x2000 # Common control shared messages
CCM_LAST= CCM_FIRST + 0x200
CCM_SETCOLORSCHEME = (CCM_FIRST + 2) # lParam is color scheme

TB_ENABLEBUTTON          = (WM_USER + 1)
TB_CHECKBUTTON           = (WM_USER + 2)
TB_PRESSBUTTON           = (WM_USER + 3)
TB_HIDEBUTTON            = (WM_USER + 4)
TB_INDETERMINATE         = (WM_USER + 5)
TB_MARKBUTTON            = (WM_USER + 6)
TB_ISBUTTONENABLED       = (WM_USER + 9)
TB_ISBUTTONCHECKED       = (WM_USER + 10)
TB_ISBUTTONPRESSED       = (WM_USER + 11)
TB_ISBUTTONHIDDEN        = (WM_USER + 12)
TB_ISBUTTONINDETERMINATE = (WM_USER + 13)
TB_ISBUTTONHIGHLIGHTED   = (WM_USER + 14)
TB_SETSTATE              = (WM_USER + 17)
TB_GETSTATE              = (WM_USER + 18)
TB_ADDBITMAP             = (WM_USER + 19)
TB_GETBUTTON             = (WM_USER + 23)
TB_COMMANDTOINDEX        = (WM_USER + 25)
TB_SAVERESTOREW          = (WM_USER + 76)
TB_CUSTOMIZE             = (WM_USER + 27)
TB_ADDSTRINGW            = (WM_USER + 77)
TB_GETITEMRECT           = (WM_USER + 29)
TB_SETBUTTONSIZE         = (WM_USER + 31)
TB_SETBITMAPSIZE         = (WM_USER + 32)
TB_GETTOOLTIPS           = (WM_USER + 35)
TB_SETTOOLTIPS           = (WM_USER + 36)
TB_SETPARENT             = (WM_USER + 37)
TB_SETROWS               = (WM_USER + 39)
TB_GETROWS               = (WM_USER + 40)
TB_SETCMDID              = (WM_USER + 42)
TB_CHANGEBITMAP          = (WM_USER + 43)
TB_GETBITMAP             = (WM_USER + 44)
TB_REPLACEBITMAP         = (WM_USER + 46)
TB_SETINDENT             = (WM_USER + 47)
TB_GETRECT               = (WM_USER + 51) # wParam is the Cmd instead of index
TB_SETHOTIMAGELIST       = (WM_USER + 52)
TB_GETHOTIMAGELIST       = (WM_USER + 53)
TB_SETDISABLEDIMAGELIST  = (WM_USER + 54)
TB_GETDISABLEDIMAGELIST  = (WM_USER + 55)
TB_SETSTYLE              = (WM_USER + 56)
TB_GETSTYLE              = (WM_USER + 57)
TB_GETBUTTONSIZE         = (WM_USER + 58)
TB_SETBUTTONWIDTH        = (WM_USER + 59)
TB_SETMAXTEXTROWS        = (WM_USER + 60)
TB_GETTEXTROWS           = (WM_USER + 61)
TB_GETOBJECT             = (WM_USER + 62)  # wParam == IID, lParam void **ppv
TB_GETHOTITEM            = (WM_USER + 71)
TB_SETHOTITEM            = (WM_USER + 72)  # wParam == iHotItem
TB_SETANCHORHIGHLIGHT    = (WM_USER + 73)  # wParam == TRUE/FALSE
TB_GETANCHORHIGHLIGHT    = (WM_USER + 74)
TB_GETBUTTONTEXTW        = (WM_USER + 75)
TB_GETINSERTMARK         = (WM_USER + 79)  # lParam == LPTBINSERTMARK
TB_SETINSERTMARK         = (WM_USER + 80)  # lParam == LPTBINSERTMARK
TB_INSERTMARKHITTEST     = (WM_USER + 81)  # wParam == LPPOINT lParam == LPTBINSERTMARK
TB_MOVEBUTTON            = (WM_USER + 82)
TB_GETMAXSIZE            = (WM_USER + 83)  # lParam == LPSIZE
TB_SETEXTENDEDSTYLE      = (WM_USER + 84)  # For TBSTYLE_EX_*
TB_GETEXTENDEDSTYLE      = (WM_USER + 85)  # For TBSTYLE_EX_*
TB_GETPADDING            = (WM_USER + 86)
TB_SETPADDING            = (WM_USER + 87)
TB_SETINSERTMARKCOLOR    = (WM_USER + 88)
TB_GETINSERTMARKCOLOR    = (WM_USER + 89)
TB_SETCOLORSCHEME        = CCM_SETCOLORSCHEME  # lParam is color scheme
#TB_GETCOLORSCHEME        = CCM_GETCOLORSCHEME      # fills in COLORSCHEME pointed to by lParam
#TB_SETUNICODEFORMAT      = CCM_SETUNICODEFORMAT
#TB_GETUNICODEFORMAT      = CCM_GETUNICODEFORMAT
TB_MAPACCELERATORW       = (WM_USER + 90)  # wParam == ch, lParam int * pidBtn
TB_SETIMAGELIST          = (WM_USER + 48)
TB_GETIMAGELIST          = (WM_USER + 49)
TB_LOADIMAGES            = (WM_USER + 50)
TB_BUTTONSTRUCTSIZE      = (WM_USER + 30)
TB_ADDBUTTONS            = (WM_USER + 68)
TB_AUTOSIZE              = (WM_USER + 33)
TB_INSERTBUTTONW         = (WM_USER + 67)
TB_DELETEBUTTON          = (WM_USER + 22)
TB_BUTTONCOUNT           = (WM_USER + 24)
TB_GETBUTTONINFO         = (WM_USER + 63)
TB_SETBUTTONINFO         = (WM_USER + 64)

TB_GETIDEALSIZE          = (WM_USER + 99)

TB_SETWINDOWTHEME       =CCM_SETWINDOWTHEME

TBSTYLE_ALTDRAG          = 0x0400
TBSTYLE_AUTOSIZE = 0x0010
TBSTYLE_BUTTON = 0x0000
TBSTYLE_CUSTOMERASE      = 0x2000
TBSTYLE_DROPDOWN = 8
TBSTYLE_FLAT             = 0x0800
TBSTYLE_LIST             = 0x1000
TBSTYLE_REGISTERDROP     = 0x4000
TBSTYLE_TOOLTIPS         = 0x0100
TBSTYLE_TRANSPARENT      = 0x8000
TBSTYLE_WRAPABLE         = 0x0200

BTNS_BUTTON      = 0x0000
BTNS_SEP         = 0x0001
BTNS_CHECK       = 0x0002
BTNS_GROUP       = 0x0004
BTNS_DROPDOWN    = 0x0008
BTNS_AUTOSIZE    = 0x0010 #automatically calculate the cx of the button
BTNS_NOPREFIX    = 0x0020 #this button should not have accel prefix
BTNS_SHOWTEXT    = 0x0040 # ignored unless TBSTYLE_EX_MIXEDBUTTONS is set
BTNS_WHOLEDROPDOWN = 0x0080 # draw drop-down arrow, but without split arrow section
BTNS_CHECKGROUP  = BTNS_CHECK | BTNS_GROUP

TBSTYLE_EX_MIXEDBUTTONS              = 0x00000008
TBSTYLE_EX_HIDECLIPPEDBUTTONS        = 0x00000010 # don't show partially obscured buttons
TBSTYLE_EX_DOUBLEBUFFER              = 0x00000080 # Double Buffer the toolbar
TBSTYLE_EX_DRAWDDARROWS              = 0x00000001

TBSTATE_CHECKED         =0x01
TBSTATE_PRESSED         =0x02
TBSTATE_ENABLED         =0x04
TBSTATE_HIDDEN          =0x08
TBSTATE_INDETERMINATE   =0x10
TBSTATE_WRAP            =0x20
TBSTATE_ELLIPSES        =0x40
TBSTATE_MARKED          =0x80

#define TBSTYLE_BUTTON          0x0000  // obsolete; use BTNS_BUTTON instead
#define TBSTYLE_SEP             0x0001  // obsolete; use BTNS_SEP instead
#TBSTYLE_CHECK           =0x0002  # obsolete; use BTNS_CHECK instead
#TBSTYLE_GROUP           =0x0004  # obsolete; use BTNS_GROUP instead
#TBSTYLE_CHECKGROUP      =(TBSTYLE_GROUP | TBSTYLE_CHECK)  # obsolete; use BTNS_CHECKGROUP instead
#define TBSTYLE_DROPDOWN        0x0008  // obsolete; use BTNS_DROPDOWN instead
#define TBSTYLE_AUTOSIZE        0x0010  // obsolete; use BTNS_AUTOSIZE instead
#define TBSTYLE_NOPREFIX        0x0020  // obsolete; use BTNS_NOPREFIX instead

#define TBSTYLE_TOOLTIPS        0x0100
#define TBSTYLE_WRAPABLE        0x0200
#define TBSTYLE_ALTDRAG         0x0400
#define TBSTYLE_FLAT            0x0800
#define TBSTYLE_LIST            0x1000
#define TBSTYLE_CUSTOMERASE     0x2000
#define TBSTYLE_REGISTERDROP    0x4000
#define TBSTYLE_TRANSPARENT     0x8000

#define TBSTYLE_EX_DRAWDDARROWS 0x00000001

#define BTNS_BUTTON     TBSTYLE_BUTTON      // 0x0000
#define BTNS_SEP        TBSTYLE_SEP         // 0x0001
#define BTNS_CHECK      TBSTYLE_CHECK       // 0x0002
#define BTNS_GROUP      TBSTYLE_GROUP       // 0x0004
#BTNS_CHECKGROUP = TBSTYLE_CHECKGROUP
#define BTNS_DROPDOWN   TBSTYLE_DROPDOWN    // 0x0008
#define BTNS_AUTOSIZE   TBSTYLE_AUTOSIZE    // 0x0010; automatically calculate the cx of the button
#define BTNS_NOPREFIX   TBSTYLE_NOPREFIX    // 0x0020; this button should not have accel prefix
#define BTNS_SHOWTEXT   0x0040              // ignored unless TBSTYLE_EX_MIXEDBUTTONS is set
#define BTNS_WHOLEDROPDOWN  0x0080          // draw drop-down arrow, but without split arrow section

#define TBSTYLE_EX_MIXEDBUTTONS             0x00000008
#define TBSTYLE_EX_HIDECLIPPEDBUTTONS       0x00000010  // don't show partially obscured buttons

#define TBSTYLE_EX_MULTICOLUMN              0x00000002 // conflicts w/ TBSTYLE_WRAPABLE
#define TBSTYLE_EX_VERTICAL                 0x00000004

TBIF_IMAGE =0x1
TBIF_TEXT =0x2
TBIF_STATE =0x4
TBIF_STYLE =0x8
TBIF_LPARAM =0x10
TBIF_COMMAND =0x20
TBIF_SIZE =0x40
TBIF_BYINDEX =0x80000000

#IDB_STD_SMALL_COLOR      = 0
#IDB_STD_LARGE_COLOR      = 1
#IDB_VIEW_SMALL_COLOR     = 4
#IDB_VIEW_LARGE_COLOR     = 5
#IDB_HIST_SMALL_COLOR     = 8
#IDB_HIST_LARGE_COLOR     = 9
#IDB_HIST_NORMAL          = 12
#IDB_HIST_HOT             = 13
#IDB_HIST_DISABLED        = 14
#IDB_HIST_PRESSED         = 15
#
#ILC_COLOR = 0
#ILC_COLOR4 = 4
#ILC_COLOR8 = 8
#ILC_COLOR16 = 16
#ILC_COLOR24 = 24
#ILC_COLOR32 = 32
#ILC_COLORDDB = 254
#ILC_MASK = 1
#ILC_PALETTE = 2048

# index values for IDB_STD_LARGE_COLOR and IDB_STD_SMALL_COLOR
#STD_CUT                  = 0
#STD_COPY                 = 1
#STD_PASTE                = 2
#STD_UNDO                 = 3
#STD_REDOW                = 4
#STD_DELETE               = 5
#STD_FILENEW              = 6
#STD_FILEOPEN             = 7
#STD_FILESAVE             = 8
#STD_PRINTPRE             = 9
#STD_PROPERTIES           = 10
#STD_HELP                 = 11
#STD_FIND                 = 12
#STD_REPLACE              = 13
#STD_PRINT                = 14

TBN_FIRST               =-700
#define TBN_LAST                (0U-720U)

TBN_GETBUTTONINFOA      =(TBN_FIRST-0)
TBN_BEGINDRAG           =(TBN_FIRST-1)
TBN_ENDDRAG             =(TBN_FIRST-2)
TBN_BEGINADJUST         =(TBN_FIRST-3)
TBN_ENDADJUST           =(TBN_FIRST-4)
TBN_RESET               =(TBN_FIRST-5)
TBN_QUERYINSERT         =(TBN_FIRST-6)
TBN_QUERYDELETE         =(TBN_FIRST-7)
TBN_TOOLBARCHANGE       =(TBN_FIRST-8)
TBN_CUSTHELP            =(TBN_FIRST-9)
TBN_DROPDOWN            =(TBN_FIRST-10)
TBN_GETOBJECT           =(TBN_FIRST-12)

DARK_TOOLBAR_BUTTON_BORDER_BRUSH = gdi32.CreateSolidBrush(0x636363)
DARK_TOOLBAR_BUTTON_BG_BRUSH = gdi32.CreateSolidBrush(0x424242)
DARK_TOOLBAR_BUTTON_ROLLOVER_BG_COLOR = 0x2b2b2b
DARK_TOOLBAR_BUTTON_ROLLOVER_BG_BRUSH = gdi32.CreateSolidBrush(DARK_TOOLBAR_BUTTON_ROLLOVER_BG_COLOR)

def _draw_arrow(hdc, x, y):
    user32.FillRect(hdc, byref(RECT(x, y,  x + 7, y + 1)), WHITE_BRUSH)
    user32.FillRect(hdc, byref(RECT(x + 1, y + 1,  x + 6, y + 2)), WHITE_BRUSH)
    user32.FillRect(hdc, byref(RECT(x + 2, y + 2,  x + 5, y + 3)), WHITE_BRUSH)
    user32.FillRect(hdc, byref(RECT(x + 3, y + 3,  x + 4, y + 4)), WHITE_BRUSH)


########################################
# Wrapper Class
########################################
class ToolBar(Window):

    def __init__(self,
            parent_window=None,
            toolbar_bmp=None,
            toolbar_bmp_dark=None,
            toolbar_buttons=None,
            icon_size=16,
            left=0, top=0,
            style=WS_CHILD | WS_VISIBLE | CCS_NODIVIDER,
            ex_style=0, #WS_EX_COMPOSITED,
            window_title=''
            ):

        super().__init__(
            TOOLBARCLASSNAME,
            parent_window=parent_window,
            left=left, top=top,
            style=style,
            ex_style=ex_style,
            window_title=window_title
            )

        if window_title:
            user32.SetWindowTextW(self.hwnd, window_title)

        # The size can be set only before adding any bitmaps to the toolbar.
        # If an application does not explicitly set the bitmap size, the size defaults to 16 by 15 pixel
        user32.SendMessageW(self.hwnd, TB_SETBITMAPSIZE, 0, MAKELONG(icon_size, icon_size))

        user32.SendMessageW(self.hwnd, TB_SETPADDING, 0, MAKELONG(6, 6))  # 6 => 28x28

        # Do not forget to send TB_BUTTONSTRUCTSIZE if the toolbar was created by using CreateWindowEx.
        user32.SendMessageW(self.hwnd, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0)

        self.__h_bitmap_dark = None
        self.__num_buttons = len(toolbar_buttons) if toolbar_buttons else 0

        self.__wholedropdown_button_ids = []
        self.__dropdown_button_ids = []

        if toolbar_buttons:
            if toolbar_bmp:
                self.__h_bitmap = user32.LoadImageW(0, toolbar_bmp, IMAGE_BITMAP, 0, 0,
                        LR_LOADFROMFILE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS
                        )
                tb = TBADDBITMAP()
                tb.hInst = 0
                tb.nID = self.__h_bitmap
                image_list_id = user32.SendMessageW(self.hwnd, TB_ADDBITMAP, self.__num_buttons, byref(tb))

                if toolbar_bmp_dark:
                    self.__h_bitmap_dark = user32.LoadImageW(0, toolbar_bmp_dark, IMAGE_BITMAP, 0, 0,
                            LR_LOADFROMFILE | LR_LOADTRANSPARENT | LR_LOADMAP3DCOLORS
                            )
            else:
                image_list_id = 0

            tb_buttons = (TBBUTTON * self.__num_buttons)()

            j = 0
            for (i, btn) in enumerate(toolbar_buttons):
                if btn[0] == '-':
                    tb_buttons[i] = TBBUTTON(
                        5,
                        0,
                        TBSTATE_ENABLED,
                        BTNS_SEP,
                    )
                else:
                    tb_buttons[i] = TBBUTTON(
                        MAKELONG(j, image_list_id),
                        btn[1], # command_id
                        btn[3] if len(btn) > 3 else TBSTATE_ENABLED,
                        btn[2] if len(btn) > 2 else BTNS_BUTTON,
                        (BYTE * 6)(),
                        btn[4] if len(btn) > 4 else 0,  # dwData
                        btn[0]
                    )

                    if len(btn) > 2:
                        if btn[2] & BTNS_DROPDOWN or btn[2] & BTNS_WHOLEDROPDOWN:
                            self.__dropdown_button_ids.append(btn[1])
                        if btn[2] & BTNS_WHOLEDROPDOWN:
                            self.__wholedropdown_button_ids.append(btn[1])

                    j += 1

            # add buttons
            ok = user32.SendMessageW(self.hwnd, TB_ADDBUTTONS, self.__num_buttons, tb_buttons)

        # remove text from buttons
        #user32.SendMessageW(self.hwnd, TB_SETMAXTEXTROWS, 0, 0)

        user32.SendMessageW(self.hwnd, TB_AUTOSIZE, 0, 0)

        rc = RECT()
        user32.GetWindowRect(self.hwnd, byref(rc))
        self.height = rc.bottom - rc.top # - 2

        self.is_dark_mode = False
        self.parent_window.register_message_callback(WM_NOTIFY, self.on_WM_NOTIFY)

    def check_button(self, button_id, flag):
        #user32.SendMessageW(self.hwnd, TB_CHECKBUTTON, button_id, flag)
        user32.PostMessageW(self.hwnd, TB_CHECKBUTTON, button_id, flag)

    def update_size(self, *args):
        user32.SendMessageW(self.hwnd, WM_SIZE, 0, 0)

    def set_indent(self, indent):
        user32.SendMessageW(self.hwnd, TB_SETINDENT, indent, 0)

    def set_imagelist(self, h_imagelist):
        user32.SendMessageW(self.hwnd, TB_SETIMAGELIST, 0, h_imagelist)

    def apply_theme(self, is_dark):
        super().apply_theme(is_dark)
        if is_dark:
            if self.__h_bitmap_dark:
                rb = TBREPLACEBITMAP()
                rb.hInstOld = 0
                rb.hInstNew = 0
                rb.nIDOld = self.__h_bitmap
                rb.nIDNew = self.__h_bitmap_dark
                rb.nButtons = self.__num_buttons
                image_list_id = user32.SendMessageW(self.hwnd, TB_REPLACEBITMAP, 0, byref(rb))
        else:
            if self.__h_bitmap_dark:
                rb = TBREPLACEBITMAP()
                rb.hInstOld = 0
                rb.hInstNew = 0
                rb.nIDOld = self.__h_bitmap_dark
                rb.nIDNew = self.__h_bitmap
                rb.nButtons = self.__num_buttons
                image_list_id = user32.SendMessageW(self.hwnd, TB_REPLACEBITMAP, 0, byref(rb))

        # update tooltip colors
        hwnd_tooltip = user32.SendMessageW(self.hwnd, TB_GETTOOLTIPS, 0, 0)
        if hwnd_tooltip:
            uxtheme.SetWindowTheme(hwnd_tooltip, 'DarkMode_Explorer' if is_dark else 'Explorer', None)

    def on_WM_NOTIFY(self, hwnd, wparam, lparam):
        nmhdr = cast(lparam, LPNMHDR).contents
        msg = nmhdr.code
        if msg == NM_CUSTOMDRAW and nmhdr.hwndFrom == self.hwnd:

            nmtb = cast(lparam, LPNMTBCUSTOMDRAW).contents
            nmcd = nmtb.nmcd

            if nmcd.dwDrawStage == CDDS_PREPAINT:
                # toolbar background
                user32.FillRect(nmcd.hdc, byref(nmcd.rc), BG_BRUSH_DARK if self.is_dark else COLOR_3DFACE + 1)  # COLOR_WINDOW
                return CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTERASE | TBCDRF_USECDCOLORS

            if not self.is_dark:
                return

            elif nmcd.dwDrawStage == CDDS_ITEMPREPAINT:
                # make button rect 1px smaller
                if nmcd.lItemlParam not in self.__dropdown_button_ids:
                    nmcd.rc.left += 1
                    nmcd.rc.right -= 1
                nmcd.rc.top += 2
                nmcd.rc.bottom -= 2

                if nmcd.uItemState & CDIS_CHECKED:
                    ########################################
                    # checked button state
                    ########################################

                    # border
                    user32.FillRect(nmcd.hdc, byref(nmcd.rc), DARK_TOOLBAR_BUTTON_BORDER_BRUSH)

                    # make 1px smaller
                    nmcd.rc.left += 1
                    nmcd.rc.right -= 1
                    nmcd.rc.top += 1
                    nmcd.rc.bottom -= 1

                    user32.FillRect(nmcd.hdc, byref(nmcd.rc), DARK_TOOLBAR_BUTTON_BG_BRUSH)

                    if nmcd.lItemlParam in self.__dropdown_button_ids and not nmcd.lItemlParam in self.__wholedropdown_button_ids:
                        return CDRF_NOTIFYPOSTPAINT | TBCDRF_NOBACKGROUND | TBCDRF_NOOFFSET | TBCDRF_NOETCHEDEFFECT | TBCDRF_NOEDGES

                    return TBCDRF_NOBACKGROUND | TBCDRF_NOOFFSET | TBCDRF_NOETCHEDEFFECT | TBCDRF_NOEDGES

                elif nmcd.uItemState & CDIS_HOT:
                    ########################################
                    # hot (rollover) button state
                    ########################################

                    # border
                    user32.FillRect(nmcd.hdc, byref(nmcd.rc), DARK_TOOLBAR_BUTTON_BORDER_BRUSH)

                    nmcd.rc.left += 1
                    nmcd.rc.right -= 1
                    nmcd.rc.top += 1
                    nmcd.rc.bottom -= 1

                    nmtb.clrHighlightHotTrack = DARK_TOOLBAR_BUTTON_ROLLOVER_BG_COLOR

                    if nmcd.lItemlParam in self.__dropdown_button_ids:
                        return CDRF_NOTIFYPOSTPAINT | TBCDRF_NOOFFSET | TBCDRF_NOETCHEDEFFECT | TBCDRF_NOEDGES | TBCDRF_HILITEHOTTRACK #| TBCDRF_USECDCOLORS

                    return TBCDRF_HILITEHOTTRACK | TBCDRF_NOOFFSET | TBCDRF_NOETCHEDEFFECT | TBCDRF_NOEDGES

                elif nmcd.uItemState & CDIS_DISABLED:
                    ########################################
                    # disabled button state
                    ########################################
                    return TBCDRF_BLENDICON

                else:
                    ########################################
                    # default button state
                    ########################################
                    if nmcd.lItemlParam in self.__dropdown_button_ids:
                        return CDRF_NOTIFYPOSTPAINT
                    return CDRF_DODEFAULT

            elif nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT:
                if nmcd.uItemState & CDIS_HOT or nmcd.uItemState & CDIS_CHECKED:
                    if nmcd.lItemlParam in self.__wholedropdown_button_ids:
                        _draw_arrow(nmcd.hdc, nmcd.rc.left + 22, nmcd.rc.top + 8)
                    else:
                        user32.FillRect(nmcd.hdc, byref(RECT(nmcd.rc.left + 21, nmcd.rc.top - 4, nmcd.rc.left + 22, nmcd.rc.bottom + 4)),
                                DARK_TOOLBAR_BUTTON_BORDER_BRUSH)
                        _draw_arrow(nmcd.hdc, nmcd.rc.left + 25, nmcd.rc.top + 4)
                else:
                    if nmcd.lItemlParam in self.__wholedropdown_button_ids:
                        _draw_arrow(nmcd.hdc, nmcd.rc.left + 23, nmcd.rc.top + 9)
                    else:
                        _draw_arrow(nmcd.hdc, nmcd.rc.left + 26, nmcd.rc.top + 5)
                return CDRF_SKIPDEFAULT

# https://learn.microsoft.com/en-us/windows/win32/controls/rebar-control-reference

from ctypes import *
from ctypes.wintypes import *
#from winapp.wintypes_extended import MAKELPARAM
from winapp.const import WM_USER
from winapp.controls.common import *
from winapp.controls.toolbar import *

from winapp.window import *


########################################
# Class Name
########################################
REBARCLASSNAME = WC_REBAR = 'ReBarWindow32'

RBS_TOOLTIPS                  =0x00000100
RBS_VARHEIGHT                 =0x00000200
RBS_BANDBORDERS               =0x00000400
RBS_FIXEDORDER                =0x00000800
RBS_REGISTERDROP              =0x00001000
RBS_AUTOSIZE                  =0x00002000
RBS_VERTICALGRIPPER           =0x00004000
RBS_DBLCLKTOGGLE              =0x00008000

#https://learn.microsoft.com/en-us/windows/win32/api/commctrl/ns-commctrl-rebarbandinfow
class REBARBANDINFOW(Structure):
    def __init__(self, *args, **kwargs):
        super(REBARBANDINFOW, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)  # REBARBANDINFO_V6_SIZE
    _fields_ = [
        ("cbSize",                UINT),
        ("fMask",                 UINT),
        ("fStyle",                UINT),
        ("clrFore",               COLORREF),
        ("clrBack",               COLORREF),
        ("lpText",                LPWSTR),
        ("cch",                   UINT),
        ("iImage",                INT),
        ("hwndChild",             HWND),
        ("cxMinChild",            UINT),
        ("cyMinChild",            UINT),
        ("cx",                    UINT),
        ("hbmBack",               HBITMAP),
        ("wID",                   UINT),
        ("cyChild",               UINT),
        ("cyMaxChild",            UINT),
        ("cyIntegral",            UINT),
        ("cxIdeal",               UINT),
        ("lParam",                LPARAM),
        ("cxHeader",              UINT),

        ("rcChevronLocation",     RECT),
        ("uChevronState",         UINT),
    ]

#RB_INSERTBANDA  =(WM_USER +  1)
RB_DELETEBAND   =(WM_USER +  2)
RB_GETBARINFO   =(WM_USER +  3)
RB_SETBARINFO   =(WM_USER +  4)
#RB_SETBANDINFOA =(WM_USER +  6)
RB_SETPARENT    =(WM_USER +  7)
RB_HITTEST      =(WM_USER +  8)
RB_GETRECT      =(WM_USER +  9)
RB_INSERTBANDW  =(WM_USER +  10)
RB_SETBANDINFOW =(WM_USER +  11)
RB_GETBANDCOUNT =(WM_USER +  12)
RB_GETROWCOUNT  =(WM_USER +  13)
RB_GETROWHEIGHT =(WM_USER +  14)
RB_IDTOINDEX    =(WM_USER +  16)
RB_GETTOOLTIPS  =(WM_USER +  17)
RB_SETTOOLTIPS  =(WM_USER +  18)
RB_SETBKCOLOR   =(WM_USER +  19)
RB_GETBKCOLOR   =(WM_USER +  20)
RB_SETTEXTCOLOR =(WM_USER +  21)
RB_GETTEXTCOLOR =(WM_USER +  22)

#define RBSTR_CHANGERECT            0x0001   // flags for RB_SIZETORECT

RB_SIZETORECT   =(WM_USER +  23) # resize the rebar/break bands and such to this rect (lparam)
RB_SETCOLORSCHEME   =CCM_SETCOLORSCHEME  # lParam is color scheme
RB_GETCOLORSCHEME   =CCM_GETCOLORSCHEME  # fills in COLORSCHEME pointed to by lParam


#// for manual drag control
#// lparam == cursor pos
#        // -1 means do it yourself.
#        // -2 means use what you had saved before
#define RB_BEGINDRAG    (WM_USER + 24)
#define RB_ENDDRAG      (WM_USER + 25)
#define RB_DRAGMOVE     (WM_USER + 26)
#define RB_GETBARHEIGHT (WM_USER + 27)
#define RB_GETBANDINFOW (WM_USER + 28)
#define RB_GETBANDINFOA (WM_USER + 29)

#ifdef UNICODE
#define RB_GETBANDINFO   RB_GETBANDINFOW
#else
#define RB_GETBANDINFO   RB_GETBANDINFOA
#endif
RB_MINIMIZEBAND =(WM_USER + 30)
RB_MAXIMIZEBAND =(WM_USER + 31)
#define RB_GETDROPTARGET (CCM_GETDROPTARGET)
#define RB_GETBANDBORDERS (WM_USER + 34)  // returns in lparam = lprc the amount of edges added to band wparam
RB_SHOWBAND     =(WM_USER + 35)
RB_SETPALETTE   =(WM_USER + 37)
RB_GETPALETTE   =(WM_USER + 38)
#define RB_MOVEBAND     (WM_USER + 39)
#define RB_SETUNICODEFORMAT     CCM_SETUNICODEFORMAT
#define RB_GETUNICODEFORMAT     CCM_GETUNICODEFORMAT
#define RB_GETBANDMARGINS   (WM_USER + 40)
RB_SETWINDOWTHEME       =CCM_SETWINDOWTHEME
#define RB_SETEXTENDEDSTYLE (WM_USER + 41)
#define RB_GETEXTENDEDSTYLE (WM_USER + 42)
#define RB_PUSHCHEVRON      (WM_USER + 43)
RB_SETBANDWIDTH =(WM_USER + 44)

RBN_FIRST = -831
RBN_HEIGHTCHANGE    =(RBN_FIRST - 0)
#define RBN_GETOBJECT       (RBN_FIRST - 1)
#define RBN_LAYOUTCHANGED   (RBN_FIRST - 2)
#define RBN_AUTOSIZE        (RBN_FIRST - 3)
#define RBN_BEGINDRAG       (RBN_FIRST - 4)
#define RBN_ENDDRAG         (RBN_FIRST - 5)
#define RBN_DELETINGBAND    (RBN_FIRST - 6)     // Uses NMREBAR
#define RBN_DELETEDBAND     (RBN_FIRST - 7)     // Uses NMREBAR
#define RBN_CHILDSIZE       (RBN_FIRST - 8)
RBN_CHEVRONPUSHED   =(RBN_FIRST - 10)
#define RBN_SPLITTERDRAG    (RBN_FIRST - 11)
#define RBN_MINMAX          (RBN_FIRST - 21)
#define RBN_AUTOBREAK       (RBN_FIRST - 22)

RBBS_BREAK          =0x00000001  # break to new line
RBBS_FIXEDSIZE      =0x00000002  # band can't be sized
RBBS_CHILDEDGE      =0x00000004  # edge around top & bottom of child window
RBBS_HIDDEN         =0x00000008  # don't show
RBBS_NOVERT         =0x00000010  # don't show when vertical
RBBS_FIXEDBMP       =0x00000020  # bitmap doesn't move during band resize
RBBS_VARIABLEHEIGHT =0x00000040  # allow autosizing of this child vertically
RBBS_GRIPPERALWAYS  =0x00000080  # always show the gripper
RBBS_NOGRIPPER      =0x00000100  # never show the gripper
RBBS_USECHEVRON     =0x00000200  # display drop-down button for this band if it's sized smaller than ideal width
RBBS_HIDETITLE      =0x00000400  # keep band title hidden
RBBS_TOPALIGN       =0x00000800  # keep band in top row

RBBIM_STYLE           =0x00000001
RBBIM_COLORS          =0x00000002
RBBIM_TEXT            =0x00000004
RBBIM_IMAGE           =0x00000008
RBBIM_CHILD           =0x00000010
RBBIM_CHILDSIZE       =0x00000020
RBBIM_SIZE            =0x00000040
RBBIM_BACKGROUND      =0x00000080
RBBIM_ID              =0x00000100
RBBIM_IDEALSIZE       =0x00000200
RBBIM_LPARAM          =0x00000400
RBBIM_HEADERSIZE      =0x00000800
RBBIM_CHEVRONLOCATION =0x00001000
RBBIM_CHEVRONSTATE    =0x00002000


########################################
# Wrapper Class
########################################
class ReBar(Window):

    def __init__(self, parent_window=None,
            style=WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | RBBS_USECHEVRON, # | RBS_VARHEIGHT | CCS_NODIVIDER, # | RBS_BANDBORDERS,
            ex_style=WS_EX_TOOLWINDOW,
            left=0, top=0, width=0, height=0
            ):

        super().__init__(
            REBARCLASSNAME,
            left=left, top=top, width=width, height=height,
            style=style,
            ex_style=ex_style,
            parent_window=parent_window
            )

        uxtheme.SetWindowTheme(self.hwnd, '', '')

    ########################################
    #
    ########################################
    def apply_theme(self, is_dark):
        self.is_dark = is_dark
        if is_dark:

            self._light_scheme = COLORSCHEME()
            user32.SendMessageW(self.hwnd, RB_GETCOLORSCHEME, 0, byref(self._light_scheme))

            # border colors, 3d. also used for separators
            cs = COLORSCHEME()
            cs.clrBtnHighlight = COLORREF(0x202020)
            cs.clrBtnShadow = COLORREF(0x202020)
            user32.SendMessageW(self.hwnd, RB_SETCOLORSCHEME, 0, byref(cs))

            rbBand = REBARBANDINFOW()
            rbBand.fMask  = RBBIM_COLORS
            rbBand.clrBack = COLORREF(0x202020)

            cnt = user32.SendMessageW(self.hwnd, RB_GETBANDCOUNT, 0, 0)
            for i in range(cnt):
                user32.SendMessageW(self.hwnd, RB_SETBANDINFOW, i, byref(rbBand))

        else:
            user32.SendMessageW(self.hwnd, RB_SETCOLORSCHEME, 0, byref(self._light_scheme))

            rbBand = REBARBANDINFOW()
            rbBand.fMask  = RBBIM_COLORS
            rbBand.clrBack = COLORREF(0xf0f0f0)

            cnt = user32.SendMessageW(self.hwnd, RB_GETBANDCOUNT, 0, 0)
            for i in range(cnt):
                user32.SendMessageW(self.hwnd, RB_SETBANDINFOW, i, byref(rbBand))

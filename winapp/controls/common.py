from ctypes import Structure, POINTER, sizeof
from ctypes.wintypes import HWND, INT, UINT, POINT, LPARAM, DWORD, COLORREF, HDC, RECT, HBRUSH, HPEN, BOOL, BYTE, LONG, WCHAR
#from winapp.const import WM_USER
from winapp.wintypes_extended import UINT_PTR, ULONG_PTR, DWORD_PTR


class NMHDR(Structure):
    _pack_ = 8
    _fields_ = [
        ("hwndFrom", HWND),
        ("idFrom", UINT_PTR),
        ("code", INT),
    ]
LPNMHDR = POINTER(NMHDR)

#typedef struct tagNMCUSTOMDRAWINFO {
#  NMHDR     hdr;
#  DWORD     dwDrawStage;
#  HDC       hdc;
#  RECT      rc;
#  DWORD_PTR dwItemSpec;
#  UINT      uItemState;
#  LPARAM    lItemlParam;
#} NMCUSTOMDRAW, *LPNMCUSTOMDRAW;

class NMCUSTOMDRAW(Structure):
    _fields_ = [
        ("hdr", NMHDR),
        ("dwDrawStage", DWORD),
        ("hdc", HDC),
        ("rc", RECT),
        ("dwItemSpec", POINTER(DWORD)),
        ("uItemState", UINT),
        ("lItemlParam", LPARAM),
    ]
LPNMCUSTOMDRAW = POINTER(NMCUSTOMDRAW)

#typedef struct tagNMMOUSE {
#  NMHDR     hdr;
#  DWORD_PTR dwItemSpec;
#  DWORD_PTR dwItemData;
#  POINT     pt;
#  LPARAM    dwHitInfo;
#} NMMOUSE, *LPNMMOUSE;

class NMMOUSE(Structure):
    _fields_ = [
        ("hdr", NMHDR),
        ("dwItemSpec", DWORD_PTR),
        ("dwItemData", DWORD_PTR),
        ("pt", POINT),
        ("dwHitInfo", LPARAM),
    ]
LPNMMOUSE = POINTER(NMMOUSE)


CDRF_DODEFAULT          =0x00000000
CDRF_NEWFONT            =0x00000002
CDRF_SKIPDEFAULT        =0x00000004
CDRF_DOERASE            =0x00000008 # draw the background
CDRF_SKIPPOSTPAINT      =0x00000100 # don't draw the focus rect

CDRF_NOTIFYPOSTPAINT    =0x00000010
CDRF_NOTIFYITEMDRAW     =0x00000020
CDRF_NOTIFYSUBITEMDRAW  =0x00000020  # flags are the same, we can distinguish by context
CDRF_NOTIFYPOSTERASE    =0x00000040

# drawstage flags
# values under 0x00010000 are reserved for global custom draw values.
# above that are for specific controls
CDDS_PREPAINT           =0x00000001
CDDS_POSTPAINT          =0x00000002
CDDS_PREERASE           =0x00000003
CDDS_POSTERASE          =0x00000004
# the 0x000010000 bit means it's individual item specific
CDDS_ITEM               =0x00010000
CDDS_ITEMPREPAINT       =(CDDS_ITEM | CDDS_PREPAINT)
CDDS_ITEMPOSTPAINT      =(CDDS_ITEM | CDDS_POSTPAINT)
CDDS_ITEMPREERASE       =(CDDS_ITEM | CDDS_PREERASE)
CDDS_ITEMPOSTERASE      =(CDDS_ITEM | CDDS_POSTERASE)
CDDS_SUBITEM            =0x00020000

# custom draw itemState flags
CDIS_SELECTED          =0x0001
CDIS_GRAYED            =0x0002
CDIS_DISABLED          =0x0004
CDIS_CHECKED           =0x0008
CDIS_FOCUS             =0x0010
CDIS_DEFAULT           =0x0020
CDIS_HOT               =0x0040
CDIS_MARKED            =0x0080
CDIS_INDETERMINATE     =0x0100
CDIS_SHOWKEYBOARDCUES   =0x0200
CDIS_NEARHOT            =0x0400
CDIS_OTHERSIDEHOT       =0x0800
CDIS_DROPHILITED        =0x1000

# Toolbar custom draw return flags
TBCDRF_NOEDGES             =0x00010000  # Don't draw button edges
TBCDRF_HILITEHOTTRACK      =0x00020000  # Use color of the button bk when hottracked
TBCDRF_NOOFFSET            =0x00040000  # Don't offset button if pressed
TBCDRF_NOMARK              =0x00080000  # Don't draw default highlight of image/text for TBSTATE_MARKED
TBCDRF_NOETCHEDEFFECT      =0x00100000  # Don't draw etched effect for disabled items
TBCDRF_BLENDICON           =0x00200000  # Use ILD_BLEND50 on the icon image
TBCDRF_NOBACKGROUND        =0x00400000  # Use ILD_BLEND50 on the icon image
TBCDRF_USECDCOLORS         =0x00800000  # Use CustomDrawColors to RenderText regardless of VisualStyle

# Common Control Styles
# https://learn.microsoft.com/en-us/windows/win32/controls/common-control-styles
CCS_TOP = 1
CCS_NOMOVEY = 2
CCS_BOTTOM = 3
CCS_NORESIZE = 4
CCS_NOPARENTALIGN = 8
CCS_ADJUSTABLE = 32
CCS_NODIVIDER = 64
CCS_VERT = 128
CCS_LEFT = 129
CCS_NOMOVEX = 130
CCS_RIGHT = 131

# Common control shared messages
CCM_FIRST = 0x2000
CCM_SETBKCOLOR          =(CCM_FIRST + 1) # lParam is bkColor
CCM_SETWINDOWTHEME      =(CCM_FIRST + 0xb)

#typedef struct tagCOLORSCHEME {
#   DWORD            dwSize;
#   COLORREF         clrBtnHighlight;       // highlight color
#   COLORREF         clrBtnShadow;          // shadow color
#} COLORSCHEME, *LPCOLORSCHEME;

class COLORSCHEME(Structure):
    def __init__(self, *args, **kwargs):
        super(COLORSCHEME, self).__init__(*args, **kwargs)
        self.dwSize = sizeof(self)
    _fields_ = [
        ("dwSize",                DWORD),
        ("clrBtnHighlight",       COLORREF),
        ("clrBtnShadow",          COLORREF),
    ]

CCM_SETCOLORSCHEME      =(CCM_FIRST + 2) # lParam is color scheme
CCM_GETCOLORSCHEME      =(CCM_FIRST + 3) # fills in COLORSCHEME pointed to by lParam
CCM_GETDROPTARGET       =(CCM_FIRST + 4)
CCM_SETUNICODEFORMAT    =(CCM_FIRST + 5)
CCM_GETUNICODEFORMAT    =(CCM_FIRST + 6)

#if (NTDDI_VERSION >= NTDDI_WINXP)
#define COMCTL32_VERSION  6
#else
#define COMCTL32_VERSION  5
#endif

#define CCM_SETVERSION          (CCM_FIRST + 0x7)
#define CCM_GETVERSION          (CCM_FIRST + 0x8)
#define CCM_SETNOTIFYWINDOW     (CCM_FIRST + 0x9) // wParam == hwndParent.
#if (NTDDI_VERSION >= NTDDI_WINXP)
#define CCM_SETWINDOWTHEME      (CCM_FIRST + 0xb)
#define CCM_DPISCALE            (CCM_FIRST + 0xc) // wParam == Awareness


#typedef struct tagPAINTSTRUCT {
#  HDC  hdc;
#  BOOL fErase;
#  RECT rcPaint;
#  BOOL fRestore;
#  BOOL fIncUpdate;
#  BYTE rgbReserved[32];
#} PAINTSTRUCT

class PAINTSTRUCT(Structure):
    _fields_ = [
        ("hdc",            HDC),
        ("fErase",         BOOL),
        ("rcPaint",        RECT),
        ("fRestore",       BOOL),
        ("fIncUpdate",     BOOL),
        ("rgbReserved",    BYTE * 32),
    ]


#====== Generic WM_NOTIFY notification codes =================================
NM_FIRST = 0
NM_OUTOFMEMORY          =(NM_FIRST-1)
NM_CLICK                =(NM_FIRST-2)    # uses NMCLICK struct
NM_DBLCLK               =(NM_FIRST-3)
NM_RETURN               =(NM_FIRST-4)
NM_RCLICK               =(NM_FIRST-5)    # uses NMCLICK struct
NM_RDBLCLK              =(NM_FIRST-6)
NM_SETFOCUS             =(NM_FIRST-7)
NM_KILLFOCUS            =(NM_FIRST-8)
NM_CUSTOMDRAW           =(NM_FIRST-12)
NM_HOVER                =(NM_FIRST-13)
NM_NCHITTEST            =(NM_FIRST-14)   # uses NMMOUSE struct
NM_KEYDOWN              =(NM_FIRST-15)   # uses NMKEY struct
NM_RELEASEDCAPTURE      =(NM_FIRST-16)
NM_SETCURSOR            =(NM_FIRST-17)   # uses NMMOUSE struct
NM_CHAR                 =(NM_FIRST-18)   # uses NMCHAR struct
NM_TOOLTIPSCREATED      =(NM_FIRST-19)   # notify of when the tooltips window is create
NM_LDOWN                =(NM_FIRST-20)
NM_RDOWN                =(NM_FIRST-21)
NM_THEMECHANGED         =(NM_FIRST-22)
NM_FONTCHANGED          =(NM_FIRST-23)
NM_CUSTOMTEXT           =(NM_FIRST-24)   # uses NMCUSTOMTEXT struct
NM_TVSTATEIMAGECHANGING =(NM_FIRST-24)   # uses NMTVSTATEIMAGECHANGING struct, defined after HTREEITEM

########################################
# Button Notifications (also used e.g. by Toolbars)
########################################
BN_CLICKED          =0
BN_PAINT            =1
BN_HILITE           =2
BN_UNHILITE         =3
BN_DISABLE          =4
BN_DOUBLECLICKED    =5
BN_PUSHED           =BN_HILITE
BN_UNPUSHED         =BN_UNHILITE
BN_DBLCLK           =BN_DOUBLECLICKED
BN_SETFOCUS         =6
BN_KILLFOCUS        =7

#ICC_LISTVIEW_CLASSES   =0x00000001 # listview, header
#ICC_TREEVIEW_CLASSES   =0x00000002 # treeview, tooltips
#ICC_BAR_CLASSES        =0x00000004 # toolbar, statusbar, trackbar, tooltips
#ICC_TAB_CLASSES        =0x00000008 # tab, tooltips
#ICC_UPDOWN_CLASS       =0x00000010 # updown
#ICC_PROGRESS_CLASS     =0x00000020 # progress
#ICC_HOTKEY_CLASS       =0x00000040 # hotkey
#ICC_ANIMATE_CLASS      =0x00000080 # animate
#ICC_WIN95_CLASSES      =0x000000FF
#ICC_DATE_CLASSES       =0x00000100 # month picker, date picker, time picker, updown
#ICC_USEREX_CLASSES     =0x00000200 # comboex
#ICC_COOL_CLASSES       =0x00000400 # rebar (coolbar) control
#ICC_INTERNET_CLASSES   =0x00000800
#ICC_PAGESCROLLER_CLASS =0x00001000 # page scroller
#ICC_NATIVEFNTCTL_CLASS =0x00002000 # native font control
#ICC_STANDARD_CLASSES   =0x00004000
#ICC_LINK_CLASS         =0x00008000
#
#class INITCOMMONCONTROLSEX(Structure):
#    _fields_ = (
#        ('dwSize', DWORD),
#        ('dwICC', DWORD),
#    )

#def init_common_controls(icc=ICC_WIN95_CLASSES):
#    ice = INITCOMMONCONTROLSEX()
#    ice.dwSize = sizeof(INITCOMMONCONTROLSEX)
#    ice.dwICC = icc
#    return Comctl32.InitCommonControlsEx(byref(ice))

#typedef struct tagMEASUREITEMSTRUCT {
#  UINT      CtlType;
#  UINT      CtlID;
#  UINT      itemID;
#  UINT      itemWidth;
#  UINT      itemHeight;
#  ULONG_PTR itemData;
#} MEASUREITEMSTRUCT, *PMEASUREITEMSTRUCT, *LPMEASUREITEMSTRUCT;

class MEASUREITEMSTRUCT(Structure):
    _fields_ = [
        ("CtlType", UINT),
        ("CtlID", UINT),
        ("itemID", UINT),
        ("itemWidth", UINT),
        ("itemHeight", UINT),
        ("lItemlParam", ULONG_PTR),
    ]
LPMEASUREITEMSTRUCT = POINTER(MEASUREITEMSTRUCT)

class DRAWITEMSTRUCT(Structure):
    _fields_ = [
        ("CtlType", UINT),
        ("CtlID", UINT),
        ("itemID", UINT),
        ("itemAction", UINT),
        ("itemState", UINT),

        ("hwndItem", HWND),
        ("hDC", HDC),
        ("rcItem", RECT),
        ("itemData", ULONG_PTR),
    ]

#typedef struct tagTEXTMETRICW {
#  LONG  tmHeight;
#  LONG  tmAscent;
#  LONG  tmDescent;
#  LONG  tmInternalLeading;
#  LONG  tmExternalLeading;
#  LONG  tmAveCharWidth;
#  LONG  tmMaxCharWidth;
#  LONG  tmWeight;
#  LONG  tmOverhang;
#  LONG  tmDigitizedAspectX;
#  LONG  tmDigitizedAspectY;
#  WCHAR tmFirstChar;
#  WCHAR tmLastChar;
#  WCHAR tmDefaultChar;
#  WCHAR tmBreakChar;
#  BYTE  tmItalic;
#  BYTE  tmUnderlined;
#  BYTE  tmStruckOut;
#  BYTE  tmPitchAndFamily;
#  BYTE  tmCharSet;
#} TEXTMETRICW, *PTEXTMETRICW, *NPTEXTMETRICW, *LPTEXTMETRICW;

class TEXTMETRICW(Structure):
    _fields_ = [
        ("tmHeight", 			    LONG),
        ("tmAscent",                LONG),
        ("tmDescent",               LONG),
        ("tmInternalLeading",       LONG),
        ("tmExternalLeading",       LONG),
        ("tmAveCharWidth",          LONG),
        ("tmMaxCharWidth",          LONG),
        ("tmWeight",                LONG),
        ("tmOverhang",              LONG),
        ("tmDigitizedAspectX",      LONG),
        ("tmDigitizedAspectY",      LONG),
        ("tmFirstChar",             WCHAR),
        ("tmLastChar",              WCHAR),
        ("tmDefaultChar",           WCHAR),
        ("tmBreakChar",             WCHAR),
        ("tmItalic",                BYTE),
        ("tmUnderlined",            BYTE),
        ("tmStruckOut",             BYTE),
        ("tmPitchAndFamily",        BYTE),
        ("tmCharSet",               BYTE),
    ]

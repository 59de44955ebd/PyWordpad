# https://learn.microsoft.com/en-us/windows/win32/controls/rich-edit-controls

from ctypes import Structure, sizeof, CFUNCTYPE, WINFUNCTYPE, c_void_p, c_uint64
from ctypes.wintypes import *

#from winapp.wintypes_extended import MAKELPARAM
from winapp.const import WM_USER
from winapp.window import *
from winapp.dlls import kernel32

kernel32.LoadLibraryW('Msftedit.dll')

#For Microsoft Rich Edit 4.1 (Msftedit.dll), specify MSFTEDIT_CLASS as the window class.
#For all previous versions, specify RICHEDIT_CLASS. For more information, see Versions of Rich Edit.

MSFTEDIT_CLASS = "RICHEDIT50W"

#// NOTE:  MSFTEDIT.DLL only registers MSFTEDIT_CLASS.  If an application wants
#// to use the following RichEdit classes, it needs to load riched20.dll.
#// Otherwise, CreateWindow with RICHEDIT_CLASS will fail.
#// This also applies to any dialog that uses RICHEDIT_CLASS
#// RichEdit 2.0 Window Class

RICHEDIT_CLASSW = "RichEdit20W"

# RichEdit messages

#define WM_CONTEXTMENU			0x007B
#define WM_UNICHAR				0x0109
#define WM_PRINTCLIENT			0x0318
#define EM_GETLIMITTEXT 		(WM_USER + 37)
#ifndef EM_POSFROMCHAR
#define EM_POSFROMCHAR			(WM_USER + 38)
#define EM_CHARFROMPOS			(WM_USER + 39)
#define EM_SCROLLCARET			(WM_USER + 49)
EM_CANPASTE 			=(WM_USER + 50)
#define EM_DISPLAYBAND			(WM_USER + 51)
EM_EXGETSEL 			=(WM_USER + 52)
#define EM_EXLIMITTEXT			(WM_USER + 53)
#define EM_EXLINEFROMCHAR		(WM_USER + 54)
#define EM_EXSETSEL 			(WM_USER + 55)
#define EM_FINDTEXT 			(WM_USER + 56)
EM_FORMATRANGE			=(WM_USER + 57)
EM_GETCHARFORMAT		=(WM_USER + 58)
#define EM_GETEVENTMASK 		(WM_USER + 59)
#define EM_GETOLEINTERFACE		(WM_USER + 60)
EM_GETPARAFORMAT		=(WM_USER + 61)
EM_GETSELTEXT			=(WM_USER + 62)
#define EM_HIDESELECTION		(WM_USER + 63)
#define EM_PASTESPECIAL 		(WM_USER + 64)
#define EM_REQUESTRESIZE		(WM_USER + 65)
#define EM_SELECTIONTYPE		(WM_USER + 66)
EM_SETBKGNDCOLOR		=(WM_USER + 67)
EM_SETCHARFORMAT		=(WM_USER + 68)
EM_SETEVENTMASK 		=(WM_USER + 69)
#define EM_SETOLECALLBACK		(WM_USER + 70)
EM_SETPARAFORMAT		=(WM_USER + 71)
EM_SETTARGETDEVICE		=(WM_USER + 72)
EM_STREAMIN 			=(WM_USER + 73)
EM_STREAMOUT			=(WM_USER + 74)
#define EM_GETTEXTRANGE 		(WM_USER + 75)
#define EM_FINDWORDBREAK		(WM_USER + 76)
#define EM_SETOPTIONS			(WM_USER + 77)
#define EM_GETOPTIONS			(WM_USER + 78)
#define EM_FINDTEXTEX			(WM_USER + 79)
#define EM_GETWORDBREAKPROCEX	(WM_USER + 80)
#define EM_SETWORDBREAKPROCEX	(WM_USER + 81)
EM_SETUNDOLIMIT =(WM_USER + 82)
EM_REDO =(WM_USER + 84)
#define EM_CANREDO (WM_USER + 85)
#define EM_GETUNDONAME (WM_USER + 86)
#define EM_GETREDONAME (WM_USER + 87)
#define EM_STOPGROUPTYPING (WM_USER + 88)
EM_SETTEXTMODE =(WM_USER + 89)
#define EM_GETTEXTMODE (WM_USER + 90)

#define EM_GETAUTOURLDETECT 	(WM_USER + 92)
#define EM_SETPALETTE			(WM_USER + 93)
EM_GETTEXTEX			=(WM_USER + 94)
EM_GETTEXTLENGTHEX		=(WM_USER + 95)
#define EM_SHOWSCROLLBAR		(WM_USER + 96)
EM_SETTEXTEX			=(WM_USER + 97)

#define EM_SETPUNCTUATION (WM_USER + 100)
#define EM_GETPUNCTUATION (WM_USER + 101)
EM_SETWORDWRAPMODE =(WM_USER + 102)
#define EM_GETWORDWRAPMODE (WM_USER + 103)
#define EM_SETIMECOLOR (WM_USER + 104)
#define EM_GETIMECOLOR (WM_USER + 105)
#define EM_SETIMEOPTIONS (WM_USER + 106)
#define EM_GETIMEOPTIONS (WM_USER + 107)
#define EM_CONVPOSITION (WM_USER + 108)
#define EM_SETLANGOPTIONS (WM_USER + 120)
#define EM_GETLANGOPTIONS (WM_USER + 121)
#define EM_GETIMECOMPMODE (WM_USER + 122)
EM_FINDTEXTW =(WM_USER + 123)
#define EM_FINDTEXTEXW (WM_USER + 124)
#define EM_RECONVERSION (WM_USER + 125)
#define EM_SETIMEMODEBIAS (WM_USER + 126)
#define EM_GETIMEMODEBIAS (WM_USER + 127)
#define EM_SETBIDIOPTIONS (WM_USER + 200)
#define EM_GETBIDIOPTIONS (WM_USER + 201)
#define EM_SETTYPOGRAPHYOPTIONS (WM_USER + 202)
#define EM_GETTYPOGRAPHYOPTIONS (WM_USER + 203)
EM_SETEDITSTYLE =(WM_USER + 204)
#define EM_GETEDITSTYLE (WM_USER + 205)

#define EM_OUTLINE (WM_USER + 220)
#define EM_GETSCROLLPOS (WM_USER + 221)
#define EM_SETSCROLLPOS (WM_USER + 222)
EM_SETFONTSIZE =(WM_USER + 223)
#define EM_GETZOOM (WM_USER + 224)
EM_SETZOOM =(WM_USER + 225)
#define EM_GETVIEWKIND (WM_USER + 226)
#define EM_SETVIEWKIND (WM_USER + 227)
#define EM_GETPAGE (WM_USER + 228)
#define EM_SETPAGE (WM_USER + 229)
#define EM_GETHYPHENATEINFO (WM_USER + 230)
#define EM_SETHYPHENATEINFO (WM_USER + 231)
#define EM_GETPAGEROTATE (WM_USER + 235)
#define EM_SETPAGEROTATE (WM_USER + 236)
#define EM_GETCTFMODEBIAS (WM_USER + 237)
#define EM_SETCTFMODEBIAS (WM_USER + 238)
#define EM_GETCTFOPENSTATUS (WM_USER + 240)
#define EM_SETCTFOPENSTATUS (WM_USER + 241)
#define EM_GETIMECOMPTEXT (WM_USER + 242)
#define EM_ISIME (WM_USER + 243)
#define EM_GETIMEPROPERTY (WM_USER + 244)
#define EM_GETQUERYRTFOBJ (WM_USER + 269)
#define EM_SETQUERYRTFOBJ (WM_USER + 270)

#define EM_GETTABLEPARMS (WM_USER + 265)
#define EM_SETEDITSTYLEEX (WM_USER + 275)
#define EM_GETEDITSTYLEEX (WM_USER + 276)

#define EM_GETSTORYTYPE (WM_USER + 290)
#define EM_SETSTORYTYPE (WM_USER + 291)
#define EM_GETELLIPSISMODE (WM_USER + 305)
#define EM_SETELLIPSISMODE (WM_USER + 306)
#define ELLIPSIS_MASK 0x00000003
#define ELLIPSIS_NONE 0x00000000
#define ELLIPSIS_END 0x00000001
#define ELLIPSIS_WORD 0x00000003
#define EM_SETTABLEPARMS (WM_USER + 307)
#define EM_GETTOUCHOPTIONS (WM_USER + 310)
#define EM_SETTOUCHOPTIONS (WM_USER + 311)
EM_INSERTIMAGE =(WM_USER + 314)
#define EM_SETUIANAME (WM_USER + 320)
#define EM_GETELLIPSISSTATE (WM_USER + 322)

# from edit
EM_SETTABSTOPS          =0x00CB
EM_SETMARGINS           =0x00D3
EM_GETMARGINS           =0x00D4

# Edit control EM_SETMARGIN parameters
EC_LEFTMARGIN       =0x0001
EC_RIGHTMARGIN      =0x0002
EC_USEFONTINFO      =0xffff

#define ES_SAVESEL 0x00008000
#define ES_SUNKEN 0x00004000
#define ES_DISABLENOSCROLL 0x00002000
#define ES_SELECTIONBAR 0x01000000
ES_NOOLEDRAGDROP =0x00000008
#define ES_EX_NOCALLOLEINIT 0x00000000
#define ES_VERTICAL 0x00400000
#define ES_NOIME 0x00080000
#define ES_SELFIME 0x00040000

#define ES_EX_ALLOWEOL_CR             0x0001
#define ES_EX_ALLOWEOL_LF             0x0002
#define ES_EX_ALLOWEOL_ALL            (ES_EX_ALLOWEOL_CR | ES_EX_ALLOWEOL_LF)
#define ES_EX_CONVERT_EOL_ON_PASTE    0x0004
ES_EX_ZOOMABLE                =0x0010

TM_PLAINTEXT = 1
TM_RICHTEXT = 2
TM_SINGLELEVELUNDO = 4
TM_MULTILEVELUNDO = 8
TM_SINGLECODEPAGE = 16
TM_MULTICODEPAGE = 32

# CHARFORMAT masks
#CFM_BOLD		=0x00000001
#CFM_ITALIC		=0x00000002
#CFM_UNDERLINE	=0x00000004
#CFM_STRIKEOUT	=0x00000008
#CFM_PROTECTED	=0x00000010
#CFM_LINK		=0x00000020
#CFM_SIZE		=0x80000000
CFM_COLOR		=0x40000000
#CFM_FACE		=0x20000000
#CFM_OFFSET		=0x10000000
#CFM_CHARSET 	=0x08000000

#define SES_USECRLF 32
#define SES_USEAIMM 64
#define SES_NOIME 128
#define SES_ALLOWBEEPS 256
#define SES_UPPERCASE 512
#define SES_LOWERCASE 1024
#define SES_NOINPUTSEQUENCECHK 2048
#define SES_BIDI 4096
#define SES_SCROLLONKILLFOCUS 8192
#define SES_XLTCRCRLFTOCR 16384
#define SES_DRAFTMODE 32768
#define SES_USECTF 0x00010000
#define SES_HIDEGRIDLINES 0x00020000
#define SES_USEATFONT 0x00040000
#define SES_CUSTOMLOOK 0x00080000
#define SES_LBSCROLLNOTIFY 0x00100000
SES_CTFALLOWEMBED =0x00200000
#define SES_CTFALLOWSMARTTAG 0x00400000
#define SES_CTFALLOWPROOFING 0x00800000
#if _RICHEDIT_VER >= 0x0500
#define SES_LOGICALCARET 0x01000000
#define SES_WORDDRAGDROP 0x02000000
#define SES_SMARTDRAGDROP 0x04000000
#define SES_MULTISELECT 0x08000000
#define SES_CTFNOLOCK 0x10000000
#define SES_NOEALINEHEIGHTADJUST 0x20000000
#define SES_MAX 0x20000000

SF_TEXT =0x0001
SF_RTF =0x0002
SF_RTFNOOBJS =0x0003
SF_TEXTIZED =0x0004
SF_UNICODE =0x0010
SF_USECODEPAGE =0x0020
#SF_NCRFORNONASCII 0x40

#SFF_WRITEXTRAPAR 0x80
SFF_SELECTION =0x8000
SFF_PLAINRTF =0x4000
#SFF_PERSISTVIEWSCALE 0x2000
#SFF_KEEPDOCINFO 0x1000
#SFF_PWD 0x0800
#SF_RTFVAL 0x0700

ST_DEFAULT = 0x00
ST_KEEPUNDO = 0x01
ST_SELECTION = 0x02
ST_NEWCHARS = 0x04
ST_UNICODE = 0x08
ST_PLACEHOLDERTEXT = 0x10
ST_PLAINTEXTONLY = 0x20

GT_DEFAULT = 0
GT_USECRLF = 1
GT_SELECTION = 2
GT_RAWTEXT = 4
GT_NOHIDDENTEXT = 8

GTL_DEFAULT = 0
GTL_USECRLF = 1
GTL_PRECISE = 2
GTL_CLOSE = 4
GTL_NUMCHARS = 8
GTL_NUMBYTES = 16


class CHARFORMATW(Structure):
    def __init__(self, *args, **kwargs):
        super(CHARFORMATW, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)
    _fields_ = [
        ("cbSize", UINT),
        ("dwMask", DWORD),
        ("dwEffects", DWORD),
        ("yHeight", LONG),
        ("yOffset", LONG),
        ("crTextColor", COLORREF),
        ("bCharSet", BYTE),
        ("bPitchAndFamily", BYTE),
        ("szFaceName", WCHAR * LF_FACESIZE),
    ]

class PARAFORMAT(Structure):
    def __init__(self, *args, **kwargs):
        super(PARAFORMAT, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)
    _fields_ = [
        ("cbSize", UINT),
        ("dwMask", DWORD),
        ("wNumbering", WORD),

        ("wReserved", WORD),

        ("dxStartIndent", LONG),
        ("dxRightIndent", LONG),
        ("dxOffset", LONG),

        ("wAlignment", WORD),
        ("cTabCount", SHORT),
        ("rgxTabs", LONG * MAX_TAB_STOPS),
    ]

class PARAFORMAT2(Structure):
    def __init__(self, *args, **kwargs):
        super(PARAFORMAT2, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)
    _fields_ = [
        ("cbSize", UINT),
        ("dwMask", DWORD),
        ("wNumbering", WORD),
        ("wReserved", WORD),
        ("dxStartIndent", LONG),
        ("dxRightIndent", LONG),
        ("dxOffset", LONG),
        ("wAlignment", WORD),
        ("cTabCount", SHORT),
        ("rgxTabs", LONG * MAX_TAB_STOPS),

        ("dySpaceBefore", LONG),
        ("dySpaceAfter", LONG),
        ("dyLineSpacing", LONG),
        ("sStyle", SHORT),
        ("bLineSpacingRule", BYTE),
        ("bOutlineLevel", BYTE),
        ("wShadingWeight", WORD),
        ("wShadingStyle", WORD),
        ("wNumberingStart", WORD),
        ("wNumberingStyle", WORD),
        ("wNumberingTab", WORD),
        ("wBorderSpace", WORD),
        ("wBorderWidth", WORD),
        ("wBorders", WORD),
    ]

class SETTEXTEX(Structure):
    _fields_ = [
        ("flags", DWORD),
        ("codepage", UINT),
    ]

class GETTEXTEX(Structure):
    _fields_ = [
        ("cb", DWORD),
        ("flags", DWORD),
        ("codepage", UINT),
        ("lpDefaultChar", LPCSTR),
        ("lpUsedDefChar", LPBOOL)
    ]

class GETTEXTLENGTHEX(Structure):
    _fields_ = [
        ("flags", DWORD),
        ("codepage", UINT),
    ]

class ENDROPFILES(Structure):
    _pack_ = 4
    _fields_ = [
        ("nmhdr", NMHDR),
        ("hDrop", HANDLE),
        ("cp", LONG),
        ("fProtected", BOOL),
    ]

class CHARRANGE(Structure):
    _fields_ = [
        ("cpMin", LONG),
        ("cpMax", LONG),
    ]

class FINDTEXTW(Structure):
    _fields_ = [
        ("chrg", CHARRANGE),
        ("lpstrText", LPCWSTR),
    ]

class FORMATRANGE(Structure):
    _fields_ = [
        ("hdc", HDC),
        ("hdcTarget", HDC),
        ("rc", RECT),
        ("rcPage", RECT),
        ("chrg", CHARRANGE),
    ]

EDITSTREAMCALLBACK = WINFUNCTYPE(DWORD, DWORD_PTR, LPBYTE, LONG, POINTER(LONG))

class EDITSTREAM(Structure):
    _pack_ = 2
    _fields_ = [
        ("dwCookie", DWORD_PTR),
        ("dwError", DWORD),
        ("pfnCallback", EDITSTREAMCALLBACK),
    ]


########################################
# Wrapper Class
########################################
class RichEdit(Window):

    ########################################
    #
    ########################################
    def __init__(self, parent_window=None, style=WS_CHILD | WS_VISIBLE, ex_style=0,
            left=0, top=0, width=0, height=0, window_title=0):

        super().__init__(
            MSFTEDIT_CLASS,
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
            self.parent_window.unregister_message_callback(WM_CTLCOLOREDIT, self._on_WM_CTLCOLOREDIT)
        super().destroy_window()

    ########################################
    #
    ########################################
    def apply_theme(self, is_dark):
        super().apply_theme(is_dark)
        uxtheme.SetWindowTheme(self.hwnd, 'DarkMode_Explorer' if is_dark else 'Explorer', None)
        user32.SendMessageW(self.hwnd, EM_SETBKGNDCOLOR, 0, CONTROL_COLOR_DARK if is_dark else 0xffffff)
        user32.SendMessageW(self.hwnd, EM_SETUNDOLIMIT, 0, 0)
        cf = CHARFORMATW()
        cf.dwMask = CFM_COLOR
        cf.crTextColor = COLORREF(0xffffff if is_dark else 0x000000)
        user32.SendMessageW(self.hwnd, EM_SETCHARFORMAT, SCF_ALL, byref(cf))
        user32.SendMessageW(self.hwnd, EM_SETUNDOLIMIT, 100, 0)

# https://learn.microsoft.com/en-us/windows/win32/controls/image-list-reference

from ctypes import Structure
from ctypes.wintypes import HBITMAP, INT, RECT
#from winapp.wintypes_extended import MAKELPARAM

#from winapp.dlls import comctl32, gdi32, user32, kernel32, ole32
#from winapp.image import *

#typedef struct tagIMAGEINFO {
#  HBITMAP hbmImage;
#  HBITMAP hbmMask;
#  int     Unused1;
#  int     Unused2;
#  RECT    rcImage;
#} IMAGEINFO;

class IMAGEINFO(Structure):
    _fields_ = [
        ("hbmImage", HBITMAP),
        ("hbmMask",  HBITMAP),
        ("Unused1",  INT),
        ("Unused2", INT),
        ("rcImage", RECT),
    ]

#BOOL ImageList_GetImageInfo(
#  HIMAGELIST himl,
#  int        i,
#  IMAGEINFO  *pImageInfo
#)
#comctl32.ImageList_GetImageInfo.argtypes = [HANDLE, INT, POINTER(IMAGEINFO)]
#
#comctl32.ImageList_LoadImageW.restype = HANDLE
#
#comctl32.ImageList_Draw.argtypes = [HANDLE, INT, HDC, INT, INT, UINT]
#
#comctl32.ImageList_GetIcon.argtypes = [HANDLE, INT, UINT]
#comctl32.ImageList_GetIcon.restype = HICON
#
#comctl32.ImageList_ReplaceIcon.argtypes = [HANDLE, INT, HICON]  # (himl, -1, hicon)
#
#comctl32.ImageList_Create.restype = HANDLE
#
##BOOL ImageList_GetIconSize(
##  HIMAGELIST himl,
##  int        *cx,
##  int        *cy
##);
#comctl32.ImageList_GetIconSize.argtypes = [HANDLE, POINTER(INT), POINTER(INT)]




ILC_MASK                =0x00000001
ILC_COLOR               =0x00000000
ILC_COLORDDB            =0x000000FE
ILC_COLOR4              =0x00000004
ILC_COLOR8              =0x00000008
ILC_COLOR16             =0x00000010
ILC_COLOR24             =0x00000018
ILC_COLOR32             =0x00000020
ILC_PALETTE             =0x00000800
ILC_MIRROR              =0x00002000
ILC_PERITEMMIRROR       =0x00008000
ILC_ORIGINALSIZE        =0x00010000
ILC_HIGHQUALITYSCALE    =0x00020000

#WINCOMMCTRLAPI HIMAGELIST  WINAPI ImageList_Create(int cx, int cy, UINT flags, int cInitial, int cGrow);
#WINCOMMCTRLAPI BOOL        WINAPI ImageList_Destroy(_In_opt_ HIMAGELIST himl);
#
#WINCOMMCTRLAPI int         WINAPI ImageList_GetImageCount(_In_ HIMAGELIST himl);
#WINCOMMCTRLAPI BOOL        WINAPI ImageList_SetImageCount(_In_ HIMAGELIST himl, _In_ UINT uNewCount);
#
#WINCOMMCTRLAPI int         WINAPI ImageList_Add(_In_ HIMAGELIST himl, _In_ HBITMAP hbmImage, _In_opt_ HBITMAP hbmMask);
#
#WINCOMMCTRLAPI int         WINAPI ImageList_ReplaceIcon(_In_ HIMAGELIST himl, _In_ int i, _In_ HICON hicon);
#WINCOMMCTRLAPI COLORREF    WINAPI ImageList_SetBkColor(_In_ HIMAGELIST himl, _In_ COLORREF clrBk);
#WINCOMMCTRLAPI COLORREF    WINAPI ImageList_GetBkColor(_In_ HIMAGELIST himl);
#WINCOMMCTRLAPI BOOL        WINAPI ImageList_SetOverlayImage(_In_ HIMAGELIST himl, _In_ int iImage, _In_ int iOverlay);

#ImageList_AddIcon = lambda himl, hicon: comctl32.ImageList_ReplaceIcon(himl, -1, hicon)

ILD_NORMAL              =0x00000000
ILD_TRANSPARENT         =0x00000001
ILD_MASK                =0x00000010
ILD_IMAGE               =0x00000020
ILD_ROP                 =0x00000040
ILD_BLEND25             =0x00000002
ILD_BLEND50             =0x00000004
ILD_OVERLAYMASK         =0x00000F00
#INDEXTOOVERLAYMASK(i)   =((i) << 8)
ILD_PRESERVEALPHA       =0x00001000  # This preserves the alpha channel in dest
ILD_SCALE               =0x00002000  # Causes the image to be scaled to cx, cy instead of clipped
ILD_DPISCALE            =0x00004000
ILD_ASYNC               =0x00008000
ILD_SELECTED            =ILD_BLEND50
ILD_FOCUS               =ILD_BLEND25
ILD_BLEND               =ILD_BLEND50


#define ILS_NORMAL              0x00000000
#define ILS_GLOW                0x00000001
#define ILS_SHADOW              0x00000002
#define ILS_SATURATE            0x00000004
#define ILS_ALPHA               0x00000008

#if (NTDDI_VERSION >= NTDDI_VISTA)
#define ILGT_NORMAL             0x00000000
#define ILGT_ASYNC              0x00000001

from ctypes import create_unicode_buffer, create_string_buffer, memmove
import json
import os
import re
import shutil
import subprocess
import sys
import time

NATIVE_DOC_EXTENSIONS = ('.rtf', '.txt')
if shutil.which('pandoc.exe') is None:
    HAS_PANDOC = False
    DOC_EXTENSIONS = NATIVE_DOC_EXTENSIONS
else:
    HAS_PANDOC = True
    DOC_EXTENSIONS = ('.rtf', '.txt', '.docx', '.odt', '.md', '.htm', '.html', '.epub')

if shutil.which('magick.exe') is None:
    HAS_MAGICK = False
    IMAGE_EXTENSIONS = ('.bmp',)
else:
    HAS_MAGICK = True
    IMAGE_EXTENSIONS = ('.bmp', '.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff')

from winapp.mainwin import *
from winapp.dialog import *
from winapp.menu import *
from winapp.dlls import *

from winapp.controls.combobox import *
from winapp.controls.rebar import *
from winapp.controls.richedit import *
from winapp.controls.statusbar import *
from winapp.controls.toolbar import *

import locale
locale.setlocale(locale.LC_TIME, '')  # use system locale for formatting date/time

from resources.const import *

APP_NAME = 'PyWordpad'
APP_VERSION = 1
APP_COPYRIGHT = '2024 github.com/59de44955ebd'
APP_DIR = os.path.dirname(os.path.abspath(__file__))

LANG = locale.windows_locale[kernel32.GetUserDefaultUILanguage()]
if not os.path.isdir(os.path.join(APP_DIR, 'resources', LANG)):
    LANG = 'en_US'

with open(os.path.join(APP_DIR, 'resources', LANG, 'StringTable1.json'), 'rb') as f:
    __ = json.loads(f.read())

def _(s):
    return __[s] if s in __ else s

TOOLBAR_ICON_SIZE = 16
TOOLBAR_HEIGHT = 24
REBAR_HEIGHT = 28

STATUSBAR_PART_ZOOM = 1

ANSI = 1
UTF_16_LE = 2
UTF_16_BE = 3
UTF_8 = 4
UTF_8_BOM = 5

ENCODINGS = {
        ANSI:       'ansi',
        UTF_16_LE:  'utf-16-le',
        UTF_16_BE:  'utf-16-be',
        UTF_8:      'utf-8',
        UTF_8_BOM:  'utf-8-sig',
        }

MAX_RECENT_FILES = 10

def twips_to_cm(twips):
    return twips * 0.0017605633802816902

def cm_to_twips(cm):
    return round(cm / 0.0017605633802816902)


class App(MainWin):

    def __init__(self, args=[]):

#        self._current_font = ['Arial', 10, 400, FALSE]  # Wordpad Win XP/Vista default
        self._current_font = ['Calibri', 11, 400, FALSE]  # Wordpad Win 7/8/10 default

        self._is_dark_mode = False
        self._show_toolbar = True
        self._show_formatbar = True
        self._show_statusbar = True
        self._filename = None
        self._is_dirty = False

        self._word_wrap = True
        self._zoom = 100

        self._search_term = ''
        self._saved_search_term = ''
        self._replace_term = ''
        self._match_case = FALSE
        self._wrap_arround = FALSE
        self._search_up = FALSE
        self._whole_word = FALSE

        self._recent_files = []
        self._current_tabs = []

        self._current_bullet_style = 0
        self._last_bullet_style = PFN_BULLET

        # used for alignment, indents, spacing, bullet style
        self._current_paraformat2 = PARAFORMAT2()

        self._print_paper_size = POINT(21000, 29700)
        self._print_margins = RECT(2000, 2500, 2000, 2500)

        left, top, width, height = self._load_state()

        # load menu resource
        with open(os.path.join(APP_DIR, 'resources', LANG, 'Menu128.json'), 'rb') as f:
            menu_data = json.loads(f.read())

        self.COMMAND_MESSAGE_MAP = {
            # File
            IDM_NEW:                    self.action_new,
            IDM_OPEN:                   self.action_open,
            IDM_SAVE:                   self.action_save,
            IDM_SAVE_AS:                self.action_save_as,
            IDM_PRINT:                  self.action_print,
            IDM_PAGE_SETUP:             self.action_page_setup,
            IDM_EXIT:                   self.action_exit,

            # Edit
            IDM_UNDO:                   self.action_undo,
            IDM_REDO:                   self.action_redo,
            IDM_CUT:                    self.action_cut,
            IDM_COPY:                   self.action_copy,
            IDM_PASTE:                  self.action_paste,
            IDM_CLEAR:                  self.action_clear,
            IDM_SELECT_ALL:             self.action_select_all,
            IDM_FIND:                   self.action_find,
            IDM_FIND_NEXT:              self.action_find_next,
            IDM_REPLACE:                self.action_replace,

            # View
            IDM_ZOOM_IN:                self.action_zoom_in,
            IDM_ZOOM_OUT:               self.action_zoom_out,
            IDM_RESTORE_DEFAULT_ZOOM:   self.action_restore_default_zoom,
            IDM_TOOLBAR:                self.action_toggle_toolbar,
            IDM_FORMAT_BAR:             self.action_toggle_formatbar,
            IDM_STATUS_BAR:             self.action_toggle_statusbar,
            IDM_DARK_MODE:              self.action_toggle_dark,

            # Insert
            IDM_DATE_AND_TIME:          self.action_insert_date_time,
            IDM_IMAGE:                  self.action_insert_image,

            # Format
            IDM_WORD_WRAP:              self.action_toggle_word_wrap,
            IDM_FONT:                   self.action_font,

            IDM_LINE_SPACING_1:         lambda: self.action_set_line_spacing(IDM_LINE_SPACING_1),
            IDM_LINE_SPACING_115:       lambda: self.action_set_line_spacing(IDM_LINE_SPACING_115),
            IDM_LINE_SPACING_15:        lambda: self.action_set_line_spacing(IDM_LINE_SPACING_15),
            IDM_LINE_SPACING_2:         lambda: self.action_set_line_spacing(IDM_LINE_SPACING_2),

            IDM_BULLET_STYLE:           self.action_toggle_bullets,
            IDM_BULLET_STYLE_NONE:      lambda: self.action_set_bullet_style(0),
            IDM_BULLET_STYLE_BULLET:    lambda: self.action_set_bullet_style(1),
            IDM_BULLET_STYLE_ARABIC:    lambda: self.action_set_bullet_style(2),
            IDM_BULLET_STYLE_LCLETTER:  lambda: self.action_set_bullet_style(3),
            IDM_BULLET_STYLE_UCLETTER:  lambda: self.action_set_bullet_style(4),
            IDM_BULLET_STYLE_LCROMAN:   lambda: self.action_set_bullet_style(5),
            IDM_BULLET_STYLE_UCROMAN:   lambda: self.action_set_bullet_style(6),

            IDM_PARAGRAPH:              self.action_paragraph,
            IDM_TABS:                   self.action_tabs,

            # Help
            IDM_ABOUT_WORDPAD:          self.action_about_wordpad,
            IDM_ABOUT_WINDOWS:          self.action_about_windows,

            IDM_BOLD:                   self.action_toggle_bold,
            IDM_ITALIC:                 self.action_toggle_italic,
            IDM_UNDERLINE:              self.action_toggle_underline,

            IDM_ALIGN_LEFT:             lambda: self.action_set_alignment(PFA_LEFT),
            IDM_ALIGN_CENTER:           lambda: self.action_set_alignment(PFA_CENTER),
            IDM_ALIGN_RIGHT:            lambda: self.action_set_alignment(PFA_RIGHT),
        }

        ########################################
        # create main window
        ########################################
        super().__init__(
                self._get_caption(),
                left=left, top=top, width=width, height=height,
                color=COLOR_3DFACE,
                menu_data=menu_data,
                menu_mod_translation_table=__,
                hicon=user32.LoadImageW(0, os.path.join(APP_DIR, 'resources', 'wordpad.ico'), IMAGE_ICON, 16, 16, LR_LOADFROMFILE),
                )

        self.hmenu_context = user32.CreatePopupMenu()
        user32.AppendMenuW(self.hmenu_context, MF_POPUP, user32.GetSubMenu(self.hmenu, 1), _('Edit'))
        user32.AppendMenuW(self.hmenu_context, MF_POPUP, user32.GetSubMenu(self.hmenu, 3), _('Insert'))
        user32.AppendMenuW(self.hmenu_context, MF_POPUP, user32.GetSubMenu(self.hmenu, 4), _('Format'))

        if self._show_toolbar:
            user32.CheckMenuItem(self.hmenu, IDM_TOOLBAR, MF_BYCOMMAND | MF_CHECKED)
        if self._show_formatbar:
            user32.CheckMenuItem(self.hmenu, IDM_FORMAT_BAR, MF_BYCOMMAND | MF_CHECKED)
        if self._show_statusbar:
            user32.CheckMenuItem(self.hmenu, IDM_STATUS_BAR, MF_BYCOMMAND | MF_CHECKED)
        if self._word_wrap:
            user32.CheckMenuItem(self.hmenu, IDM_WORD_WRAP, MF_BYCOMMAND | MF_CHECKED)

        if len(self._recent_files):
            self._update_recent(False)

        self._create_statusbar()
        self._create_rebar()
        self._create_richedit()
        self._create_dialogs()

        ########################################
        #
        ########################################
        def _on_WM_SIZE(hwnd, wparam, lparam):
            width, height = lparam & 0xFFFF, (lparam >> 16) & 0xFFFF
            self.statusbar.update_size()
            self.statusbar.right_align_parts(width)  # keep statusbar parts right aligned
            #user32.SetWindowPos(hwnd_richedit, 0, 0, 0, width, height - status_bar.height, 0)
            self.rich_edit.resize_window(width, height - self.rebar_height -
                    (self.statusbar.height if self._show_statusbar else 0) - 2)
        self.register_message_callback(WM_SIZE, _on_WM_SIZE)

        ########################################
        #
        ########################################
        def _on_WM_NOTIFY(hwnd, wparam, lparam):
            nmhdr = cast(lparam, LPNMHDR).contents
            msg = nmhdr.code

            if msg == EN_SELCHANGE:
                self._update_formatting()
                self._check_if_text_selected()

            elif msg == EN_DROPFILES:
                edf = cast(lparam, POINTER(ENDROPFILES))
                for filename in self.get_dropped_items(edf.contents.hDrop):
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in IMAGE_EXTENSIONS:
                        self.create_timer(lambda: self.action_insert_image(filename), 0, True)
                    elif ext in DOC_EXTENSIONS:
                        self.create_timer(lambda: self.action_open(filename), 0, True)
                        break
                    else:
                        self.statusbar.set_text(_('Unsupported file format'))
                return 1

            elif msg == RBN_HEIGHTCHANGE:
                rc = self.get_client_rect()
                width, height = rc.right, rc.bottom
                rc = self.rebar.get_window_rect()
                self.rebar_height = rc.bottom - rc.top
                self.rich_edit.set_window_pos(0, self.rebar_height + 2, width, height - self.rebar_height -
                        (self.statusbar.height if self._show_statusbar else 0) - 2, SWP_NOSIZE)

            elif msg == TBN_DROPDOWN:
                nmtb = cast(lparam, LPNMTOOLBARW).contents
                if nmtb.iItem == IDM_BULLET_STYLE:
                    hmenu_format = user32.GetSubMenu(self.hmenu, 4)
                    hmenu_numbering = user32.GetSubMenu(hmenu_format, 2)
                    user32.MapWindowPoints(self.toolbar_formatting.hwnd, None, byref(nmtb.rcButton), 2)
                    self._show_popupmenu(hmenu_numbering, nmtb.rcButton.left, nmtb.rcButton.bottom)
                elif nmtb.iItem == IDM_LINE_SPACING:
                    hmenu_format = user32.GetSubMenu(self.hmenu, 4)
                    hmenu_line_spacing = user32.GetSubMenu(hmenu_format, 1)
                    user32.MapWindowPoints(self.toolbar_formatting.hwnd, None, byref(nmtb.rcButton), 2)
                    self._show_popupmenu(hmenu_line_spacing, nmtb.rcButton.left, nmtb.rcButton.bottom)

        self.register_message_callback(WM_NOTIFY, _on_WM_NOTIFY)

        ########################################
        #
        ########################################
        def _on_WM_COMMAND(hwnd, wparam, lparam):

            if lparam == 0 or lparam == self.toolbar_standard.hwnd or lparam == self.toolbar_formatting.hwnd:
                command_id = LOWORD(wparam)
                if command_id in self.COMMAND_MESSAGE_MAP:
                    self.COMMAND_MESSAGE_MAP[command_id]()
                    return FALSE

            notification_code = HIWORD(wparam)

            if lparam == self.rich_edit.hwnd:
                if notification_code == EN_CHANGE:
                    if self._filename:
                        self._is_dirty = user32.SendMessageW(self.rich_edit.hwnd, EM_CANUNDO, 0, 0)
                    else:
                        self._is_dirty = user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXTLENGTH, 0, 0) > 0
                    self.set_window_text(self._get_caption())
                    self._check_if_has_text()

            elif lparam == self.combobox_font.hwnd and notification_code == CBN_SELCHANGE:
                self.action_set_font()

            elif lparam == self.combobox_fontsize.hwnd and notification_code == CBN_SELCHANGE:
                self.action_set_fontsize()

            return FALSE
        self.register_message_callback(WM_COMMAND, _on_WM_COMMAND)

        self.register_message_callback(WM_CLOSE, self.action_exit, True)

        if self._is_dark_mode:
            self.apply_theme(True)
            user32.CheckMenuItem(self.hmenu, IDM_DARK_MODE, MF_BYCOMMAND | MF_CHECKED)

        self.show()

        # only works after show
        user32.SendMessageW(self.rebar.hwnd, RB_MAXIMIZEBAND, 1, 1)

        # only works after show
        rc = self.rich_edit.get_client_rect()
        rc.left += 15
        rc.top += 5
        rc.right -= 15
        rc.bottom -= 5
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETRECT, 0, byref(rc))

        user32.SetFocus(self.rich_edit.hwnd)

#        self._load_file(os.path.join(APP_DIR, 'test_files', 'test.rtf'))
#        self._load_file(os.path.join(APP_DIR, 'test_files', 'README.md'))

    ########################################
    #
    ########################################
    def _update_recent(self, clear_menu=True):
        hmenu_file = user32.GetSubMenu(self.hmenu, 0)
        hmenu_recent = user32.GetSubMenu(hmenu_file, 8)  # TODO
        if clear_menu:
            while user32.DeleteMenu(hmenu_recent, 0, MF_BYPOSITION):
                pass
            for i in range(MAX_RECENT_FILES):
                if IDM_RECENT_FILE + i not in self.COMMAND_MESSAGE_MAP:
                    break
                del self.COMMAND_MESSAGE_MAP[IDM_RECENT_FILE + i]
        if len(self._recent_files):
            for i, filename in enumerate(self._recent_files):
                user32.AppendMenuW(hmenu_recent, MF_STRING, IDM_RECENT_FILE + i, filename)
                self.COMMAND_MESSAGE_MAP[IDM_RECENT_FILE + i] = lambda fn=filename: self.action_open(fn)

    ########################################
    #
    ########################################
    def _update_formatting(self):

        cf = CHARFORMATW()
        user32.SendMessageW(self.rich_edit.hwnd, EM_GETCHARFORMAT, SCF_SELECTION, byref(cf))
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, IDM_BOLD,
                TBSTATE_ENABLED | TBSTATE_CHECKED if cf.dwEffects & CFE_BOLD else TBSTATE_ENABLED)
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, IDM_ITALIC,
                TBSTATE_ENABLED | TBSTATE_CHECKED if cf.dwEffects & CFE_ITALIC else TBSTATE_ENABLED)
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, IDM_UNDERLINE,
                TBSTATE_ENABLED | TBSTATE_CHECKED if cf.dwEffects & CFE_UNDERLINE else TBSTATE_ENABLED)

        fn = cf.szFaceName
        self._current_font[0] = fn
        if fn in self._fonts:
            user32.SendMessageW(self.combobox_font.hwnd, CB_SETCURSEL, self._fonts.index(fn), 0)

        fs = cf.yHeight // 20
        self._current_font[1] = fs
        if fs in self._font_sizes:
            user32.SendMessageW(self.combobox_fontsize.hwnd, CB_SETCURSEL, self._font_sizes.index(fs), 0)

        self._current_font[2] = 700 if cf.dwEffects & CFE_BOLD else 400
        self._current_font[3] = TRUE if cf.dwEffects & CFE_ITALIC else FALSE

        pf = self._current_paraformat2
        user32.SendMessageW(self.rich_edit.hwnd, EM_GETPARAFORMAT, 0, byref(pf))

        for id in (IDM_ALIGN_LEFT, IDM_ALIGN_CENTER, IDM_ALIGN_RIGHT):
            user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, id,
                    TBSTATE_ENABLED | TBSTATE_CHECKED if id - IDM_ALIGN_LEFT + 1 == pf.wAlignment else TBSTATE_ENABLED)

        if pf.bLineSpacingRule == 5 and pf.dyLineSpacing in [20, 23, 30, 40]:
            id_line_spacing = [20, 23, 30, 40].index(pf.dyLineSpacing) + IDM_LINE_SPACING_1
        elif pf.bLineSpacingRule == 0:
            id_line_spacing = IDM_LINE_SPACING_1
        elif pf.bLineSpacingRule == 1:
            id_line_spacing = IDM_LINE_SPACING_15
        elif pf.bLineSpacingRule == 2:
            id_line_spacing = IDM_LINE_SPACING_2
        else:
            id_line_spacing = IDM_LINE_SPACING_115

        for i in range(IDM_LINE_SPACING_1, IDM_LINE_SPACING_2 + 1):
            user32.CheckMenuItem(self.hmenu, i, MF_BYCOMMAND | (MF_CHECKED if id_line_spacing == i else MF_UNCHECKED))

        self._current_bullet_style = pf.wNumbering
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, IDM_BULLET_STYLE,
                TBSTATE_ENABLED | TBSTATE_CHECKED if pf.wNumbering != 0 else TBSTATE_ENABLED)
        for i in range(PFN_UCROMAN + 1):
            user32.CheckMenuItem(self.hmenu, IDM_BULLET_STYLE_NONE + i, MF_BYCOMMAND | (MF_CHECKED if pf.wNumbering == i else MF_UNCHECKED))

        self._current_tabs = pf.rgxTabs[:pf.cTabCount]

    ########################################
    #
    ########################################
    def _create_statusbar(self):
        self.statusbar = StatusBar(self, parts=(0, 50), parts_right_aligned=True)
        self.statusbar.set_text('100%', STATUSBAR_PART_ZOOM)

    ########################################
    # https://learn.microsoft.com/en-us/windows/win32/controls/rebar-control-styles
    ########################################
    def _create_rebar(self):
        self.rebar = ReBar(
                self,
                height=REBAR_HEIGHT,
                style=
                    WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
                    | TBSTYLE_TOOLTIPS
                )

        toolbar_buttons = (
            (_('New'), IDM_NEW), #self.action_new),
            (_('Open'), IDM_OPEN),
            (_('Save'), IDM_SAVE),
            (_('Print'), IDM_PRINT),
            (_('Find'), IDM_FIND),
            (_('Cut'), IDM_CUT),
            (_('Copy'), IDM_COPY),
            (_('Paste'), IDM_PASTE),
            (_('Undo'), IDM_UNDO),
        )

        self.toolbar_standard = self._add_toolbar('Standard',
                toolbar_buttons,
                os.path.join(APP_DIR, 'resources', 'Bitmap140.bmp'),
                os.path.join(APP_DIR, 'resources', 'Bitmap140_dark.bmp')
                )

        toolbar_buttons = (
            (_('Bold'), IDM_BOLD),
            (_('Italic'), IDM_ITALIC),
            (_('Underline'), IDM_UNDERLINE),
            (_('Align left'), IDM_ALIGN_LEFT, BTNS_BUTTON, TBSTATE_ENABLED | TBSTATE_CHECKED),
            (_('Center'), IDM_ALIGN_CENTER),
            (_('Align Right'), IDM_ALIGN_RIGHT),
            (_('Bullet Style'), IDM_BULLET_STYLE, BTNS_DROPDOWN, TBSTATE_ENABLED, IDM_BULLET_STYLE),
            (_('Line Spacing'), IDM_LINE_SPACING, BTNS_WHOLEDROPDOWN, TBSTATE_ENABLED, IDM_LINE_SPACING),
        )

        self.toolbar_formatting = self._add_toolbar('Formatting',
                toolbar_buttons,
                os.path.join(APP_DIR, 'resources', LANG, 'Bitmap139.bmp'),
                os.path.join(APP_DIR, 'resources', LANG, 'Bitmap139_dark.bmp'),
                4 + 160 + 8 + 50 + 8) #  + 8 + 120

        fonts = []
        def _enum_fonts(lpelfe, lpntme, FontType, lParam):
            if FontType > 1:
                fonts.append(lpelfe.contents.lfFaceName)
            return 1
        proc = FONTENUMPROCW(_enum_fonts)
        hdc = user32.GetDC(0)
        lf = LOGFONTW()
        lf.lfCharSet = ANSI_CHARSET  # DEFAULT_CHARSET
        lf.lfFaceName = ''
        lf.lfPitchAndFamily = 0
        res = gdi32.EnumFontFamiliesExW(hdc, byref(lf), proc, 0, 0)
        self._fonts = sorted(fonts)

        self.combobox_font = ComboBox(
            self.toolbar_formatting,
            style=CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL,
            left=4,
            top=2,
            width=160,
            height=450  # height of dropdown!
            )
        self.combobox_font.set_font()

        for f in self._fonts:
            user32.SendMessageW(self.combobox_font.hwnd, CB_ADDSTRING, 0, f)
        if self._current_font[0] in self._fonts:
            user32.SendMessageW(self.combobox_font.hwnd, CB_SETCURSEL, self._fonts.index(self._current_font[0]), 0)

        self.combobox_fontsize = ComboBox(
            self.toolbar_formatting,
            style=CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE,
            left=4 + 160 + 8,
            top=2,
            width=50,
            height=200 # height of dropdown!
            )
        self.combobox_fontsize.set_font()
        self._font_sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
        for s in self._font_sizes:
            user32.SendMessageW(self.combobox_fontsize.hwnd, CB_ADDSTRING, 0, str(s))
        if self._current_font[1] in self._font_sizes:
            user32.SendMessageW(self.combobox_fontsize.hwnd, CB_SETCURSEL, self._font_sizes.index(self._current_font[1]), 0)

        rc = self.rebar.get_window_rect()
        self.rebar_height = rc.bottom - rc.top

        # TEST
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS)

    ########################################
    #
    ########################################
    def _add_toolbar(self, caption, toolbar_buttons, bitmap, bitmap_dark, indent=0):

        toolbar = ToolBar(
                self,
                window_title=caption,
                icon_size=TOOLBAR_ICON_SIZE,

                style=WS_CHILD
                    | WS_VISIBLE
                    | WS_CLIPSIBLINGS
                    | BTNS_AUTOSIZE
                    | CCS_NODIVIDER
                    | CCS_NOMOVEY
                    | CCS_NOPARENTALIGN
                    | CCS_NORESIZE
                    | CCS_TOP
                    | TBSTYLE_AUTOSIZE
                    | TBSTYLE_FLAT
                    | TBSTYLE_TOOLTIPS
#                    | TBSTYLE_TRANSPARENT
                    | TBSTYLE_WRAPABLE
                ,

                ex_style=WS_EX_COMPOSITED | TBSTYLE_EX_HIDECLIPPEDBUTTONS,

                toolbar_bmp=bitmap,
                toolbar_bmp_dark=bitmap_dark,
                toolbar_buttons=toolbar_buttons,
                )

        # remove text from buttons
        user32.SendMessageW(toolbar.hwnd, TB_SETMAXTEXTROWS, 0, 0)

        # Sets the indentation for the first button in a toolbar control.
        user32.SendMessageW(toolbar.hwnd, TB_SETINDENT, indent, 0)

        # LOWORD specifies the horizontal padding, in pixels. The HIWORD specifies the vertical padding
        user32.SendMessageW(toolbar.hwnd, TB_SETPADDING, 0, MAKELONG(10, 0))

        #LOWORD specifies the width, HIWORD specifies the height
        user32.SendMessageW(toolbar.hwnd, TB_SETBUTTONSIZE, 0, MAKELONG(0, 25))

        # get ideal width (for chevron)
        sz = SIZE()
        # TRUE** to retrieve the ideal height, **FALSE** to retrieve the ideal width
        user32.SendMessageW(toolbar.hwnd, TB_GETIDEALSIZE, FALSE, byref(sz))
        ideal_size = sz.cx + indent

        # Initialize band info used by both bands.
        rbBand = REBARBANDINFOW()
        rbBand.fMask  = (
                RBBIM_CHILD       # hwndChild is valid.
                | RBBIM_TEXT        # lpText is valid.
                | RBBIM_STYLE       # fStyle is valid.
                | RBBIM_CHILDSIZE   # child size members are valid.
                | RBBIM_SIZE
                | RBBIM_IDEALSIZE   # otherwise no chevron
#                | RBBIM_COLORS
                )

        rbBand.hwndChild = toolbar.hwnd
        rbBand.lpText = caption

        rbBand.fStyle = RBBS_HIDETITLE | RBBS_TOPALIGN # | RBBS_VARIABLEHEIGHT  # RBBS_NOGRIPPER  | RBBS_CHILDEDGE #  | RBBS_USECHEVRON

        # Length of the band, in pixels.
        # The default width should be set to some value wider than the text.
        rbBand.cx = ideal_size

        # Ideal width of the band, in pixels. If the band is maximized to the ideal width (see RB_MAXIMIZEBAND), the rebar control will attempt to make the band this width.
        rbBand.cxIdeal = ideal_size

        # Minimum width of the child window, in pixels. The band can't be sized smaller than this value.
        rbBand.cxMinChild = ideal_size

        # Minimum height of the child window, in pixels. The band can't be sized smaller than this value.
        rbBand.cyMinChild = REBAR_HEIGHT

        #  Initial height of the band, in pixels. This member is ignored unless the RBBS_VARIABLEHEIGHT style is specified.
        rbBand.cyChild = TOOLBAR_HEIGHT

        # Maximum height of the band, in pixels. This member is ignored unless the RBBS_VARIABLEHEIGHT style is specified.
        rbBand.cyMaxChild = TOOLBAR_HEIGHT

        # Add the band to the rebar
        user32.SendMessageW(self.rebar.hwnd, RB_INSERTBANDW, -1, byref(rbBand))  # -1

        return toolbar

    ########################################
    #
    ########################################
    def _create_richedit(self):
        self.rich_edit = RichEdit(
            self,
            left=0,
            top=self.rebar_height + 2,
            style=WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | WS_BORDER | ES_NOOLEDRAGDROP | (0 if self._word_wrap else WS_HSCROLL),
            ex_style=ES_EX_ZOOMABLE | WS_EX_ACCEPTFILES
            )

        self.rich_edit.set_font(*self._current_font)

        self.rich_edit.register_message_callback(WM_RBUTTONUP, lambda *args:
                self._show_popupmenu(self.hmenu_context))

        user32.SendMessageW(self.rich_edit.hwnd, EM_SETTEXTMODE, TM_RICHTEXT, 0)
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETEVENTMASK, 0, ENM_CHANGE | ENM_SELCHANGE | ENM_DROPFILES)

        ex_style = user32.GetWindowLongW(self.rich_edit.hwnd, GWL_EXSTYLE)
        user32.SetWindowLongA(self.rich_edit.hwnd, GWL_EXSTYLE, ex_style & ~WS_EX_CLIENTEDGE)
        user32.SetWindowPos(self.rich_edit.hwnd, 0, 0,0, 0,0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)

        user32.SendMessageW(self.rich_edit.hwnd, EM_SETTARGETDEVICE, 0, int(not self._word_wrap))

        user32.SendMessageW(self.rich_edit.hwnd, EM_SETMODIFY, FALSE, 0)

    ########################################
    #
    ########################################
    def _create_dialogs(self):
        with open(os.path.join(APP_DIR, 'resources', LANG, 'Dialog143.json'), 'rb') as f:
            dialog_dict = json.loads(f.read())

        def _dialog_proc_paragraph(hwnd, msg, wparam, lparam):

            if msg == WM_INITDIALOG:
                pf = self._current_paraformat2

                hwnd_edit = user32.GetDlgItem(hwnd, ID_LEFT)
                user32.SetWindowTextW(hwnd_edit, f'{round(twips_to_cm(pf.dxOffset + pf.dxStartIndent), 2)} cm')

                hwnd_edit = user32.GetDlgItem(hwnd, ID_RIGHT)
                user32.SetWindowTextW(hwnd_edit, f'{round(twips_to_cm(pf.dxRightIndent), 2)} cm')

                hwnd_edit = user32.GetDlgItem(hwnd, ID_FIRST_LINE)
                user32.SetWindowTextW(hwnd_edit, f'{round(-twips_to_cm(pf.dxOffset), 2)} cm')

                hwnd_edit = user32.GetDlgItem(hwnd, ID_SPACING_ABOVE)
                user32.SetWindowTextW(hwnd_edit, f'{round(twips_to_cm(pf.dySpaceBefore), 2)} cm')

                hwnd_edit = user32.GetDlgItem(hwnd, ID_SPACING_BELOW)
                user32.SetWindowTextW(hwnd_edit, f'{round(twips_to_cm(pf.dySpaceAfter), 2)} cm')

                hwnd_combo = user32.GetDlgItem(hwnd, ID_ALIGNMENT)
                user32.SendMessageW(hwnd_combo, CB_ADDSTRING, 0, _('Left'))
                user32.SendMessageW(hwnd_combo, CB_ADDSTRING, 0, _('Right'))
                user32.SendMessageW(hwnd_combo, CB_ADDSTRING, 0, _('Center'))
                user32.SendMessageW(hwnd_combo, CB_SETCURSEL, self._current_paraformat2.wAlignment - 1, 0)

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)

                if command == BN_CLICKED:
                    if control_id == ID_OK:
                        pf = self._current_paraformat2  # PARAFORMAT2()

                        pf.dwMask = PFM_OFFSET | PFM_RIGHTINDENT | PFM_STARTINDENT | PFM_ALIGNMENT | PFM_SPACEBEFORE | PFM_SPACEAFTER

                        buf = create_unicode_buffer(32)

                        hwnd_edit = user32.GetDlgItem(hwnd, ID_LEFT)
                        user32.GetWindowTextW(hwnd_edit, buf, 32)
                        v_left = float(buf.value.split(' ')[0])

                        hwnd_edit = user32.GetDlgItem(hwnd, ID_RIGHT)
                        user32.GetWindowTextW(hwnd_edit, buf, 32)
                        v_right = float(buf.value.split(' ')[0])

                        hwnd_edit = user32.GetDlgItem(hwnd, ID_FIRST_LINE)
                        user32.GetWindowTextW(hwnd_edit, buf, 32)
                        v_first = float(buf.value.split(' ')[0])

                        pf.dxOffset = -cm_to_twips(v_first)
                        pf.dxRightIndent = cm_to_twips(v_right)
                        pf.dxStartIndent = cm_to_twips(v_left + v_first)

                        hwnd_edit = user32.GetDlgItem(hwnd, ID_SPACING_ABOVE)
                        user32.GetWindowTextW(hwnd_edit, buf, 32)
                        v_before = float(buf.value.split(' ')[0])

                        hwnd_edit = user32.GetDlgItem(hwnd, ID_SPACING_BELOW)
                        user32.GetWindowTextW(hwnd_edit, buf, 32)
                        v_after = float(buf.value.split(' ')[0])

                        pf.dySpaceBefore = cm_to_twips(v_before)
                        pf.dySpaceAfter = cm_to_twips(v_after)

                        hwnd_combo = user32.GetDlgItem(hwnd, ID_ALIGNMENT)
                        idx = user32.SendMessageW(hwnd_combo, CB_GETCURSEL, 0, 0)
                        pf.wAlignment = idx + 1

                        user32.SendMessageW(self.rich_edit.hwnd, EM_SETPARAFORMAT, 0, byref(pf))

                        # for some reason pure alignment change does not trigger EN_CHANGE,
                        # so we have to update the toolbar buttons manually
                        for id in (IDM_ALIGN_LEFT, IDM_ALIGN_CENTER, IDM_ALIGN_RIGHT):
                            user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, id,
                                    TBSTATE_ENABLED | TBSTATE_CHECKED if id - IDM_ALIGN_LEFT + 1 == pf.wAlignment else TBSTATE_ENABLED)

                        user32.EndDialog(hwnd, ID_OK)

                    elif control_id == ID_CANCEL:
                        user32.EndDialog(hwnd, ID_CANCEL)

            elif msg == WM_CLOSE:
                user32.SetFocus(self.rich_edit.hwnd)

            return FALSE

        self.dialog_paragraph = Dialog(self, dialog_dict, _dialog_proc_paragraph)

        with open(os.path.join(APP_DIR, 'resources', LANG, 'Dialog145.json'), 'rb') as f:
            dialog_dict = json.loads(f.read())

        def _dialog_proc_tabs(hwnd, msg, wparam, lparam):

            if msg == WM_INITDIALOG:

                hwnd_combo = user32.GetDlgItem(hwnd, ID_TAB_STOP_POSITION)
                for t in self._current_tabs:
                    user32.SendMessageW(hwnd_combo, CB_ADDSTRING, 0, f'{round(twips_to_cm(t), 2)} cm')

                hwnd_btn_set = user32.GetDlgItem(hwnd, ID_SET)
                user32.EnableWindow(hwnd_btn_set, 0)

                hwnd_btn_clear = user32.GetDlgItem(hwnd, ID_CLEAR)
                user32.EnableWindow(hwnd_btn_clear, 0)

                hwnd_btn_clear_all = user32.GetDlgItem(hwnd, ID_CLEAR_ALL)
                user32.EnableWindow(hwnd_btn_clear_all, int(len(self._current_tabs) > 0))

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)

                if command == BN_CLICKED:
                    if control_id == ID_OK:
                        hwnd_combo = user32.GetDlgItem(hwnd, ID_TAB_STOP_POSITION)
                        cnt = user32.SendMessageW(hwnd_combo, CB_GETCOUNT, 0, 0)
                        buf = create_unicode_buffer(16)
                        self._current_tabs = []
                        for i in range(cnt):
                            user32.SendMessageW(hwnd_combo, CB_GETLBTEXT, i, buf)
                            v = float(buf.value.split(' ')[0])
                            self._current_tabs.append(cm_to_twips(v))

                        tabs = self._current_tabs + [0] * (MAX_TAB_STOPS - len(self._current_tabs))
                        pf = PARAFORMAT()
                        pf.dwMask = PFM_TABSTOPS
                        pf.cTabCount = cnt
                        pf.rgxTabs = (LONG * MAX_TAB_STOPS)(*tabs) # LONG * MAX_TAB_STOPS
                        user32.SendMessageW(self.rich_edit.hwnd, EM_SETPARAFORMAT, 0, byref(pf))
                        user32.EndDialog(hwnd, ID_OK)

                    elif control_id == ID_CANCEL:
                        user32.EndDialog(hwnd, ID_CANCEL)

                    elif control_id == ID_SET:
                        hwnd_combo = user32.GetDlgItem(hwnd, ID_TAB_STOP_POSITION)
                        buf = create_unicode_buffer(32)
                        user32.GetWindowTextW(hwnd_combo, buf, 32)

                        # TODO: check if its numeric

                        user32.SendMessageW(hwnd_combo, CB_ADDSTRING, 0, buf.value)

                    elif control_id == ID_CLEAR_ALL:
                        hwnd_combo = user32.GetDlgItem(hwnd, ID_TAB_STOP_POSITION)
                        user32.SendMessageW(hwnd_combo, CB_RESETCONTENT, 0, 0)
                        hwnd_btn_set = user32.GetDlgItem(hwnd, ID_SET)
                        user32.EnableWindow(hwnd_btn_set, 0)

                        hwnd_btn_clear = user32.GetDlgItem(hwnd, ID_CLEAR)
                        user32.EnableWindow(hwnd_btn_clear, 0)

                    elif control_id == ID_CLEAR:
                        hwnd_combo = user32.GetDlgItem(hwnd, ID_TAB_STOP_POSITION)
                        idx = user32.SendMessageW(hwnd_combo, CB_GETCURSEL, 0, 0)
                        user32.SendMessageW(hwnd_combo, CB_DELETESTRING, idx, 0)

                        if user32.SendMessageW(hwnd_combo, CB_GETCOUNT, 0, 0) == 0:
                            user32.SetWindowTextW(hwnd_combo, '')
                            hwnd_btn_set = user32.GetDlgItem(hwnd, ID_SET)
                            user32.EnableWindow(hwnd_btn_set, 0)
                            hwnd_btn_clear = user32.GetDlgItem(hwnd, ID_CLEAR)
                            user32.EnableWindow(hwnd_btn_clear, 0)
                            hwnd_btn_clear_all = user32.GetDlgItem(hwnd, ID_CLEAR_ALL)
                            user32.EnableWindow(hwnd_btn_clear_all, 0)

                elif command == CBN_EDITCHANGE:
                    buf = create_unicode_buffer(32)
                    user32.GetWindowTextW(lparam, buf, 32)
                    hwnd_btn_set = user32.GetDlgItem(hwnd, ID_SET)
                    user32.EnableWindow(hwnd_btn_set, int(len(buf.value) > 0))

                elif command == CBN_SELCHANGE:
                    hwnd_btn_clear = user32.GetDlgItem(hwnd, ID_CLEAR)
                    user32.EnableWindow(hwnd_btn_clear, 1)

            elif msg == WM_CLOSE:
                user32.SetFocus(self.rich_edit.hwnd)

            return FALSE
        self.dialog_tabs = Dialog(self, dialog_dict, _dialog_proc_tabs)

        with open(os.path.join(APP_DIR, 'resources', LANG, 'Dialog1540.json'), 'rb') as f:
            dialog_dict = json.loads(f.read())

        def _dialog_proc_find(hwnd, msg, wparam, lparam):

            if msg == WM_INITDIALOG:
                hwnd_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)

                # Limit search input to 127 chars
                user32.SendMessageW(hwnd_edit, EM_SETLIMITTEXT, 127, 0)

                # check if something is selected
                pos_start, pos_end = DWORD(), DWORD()
                user32.SendMessageW(self.rich_edit.hwnd, EM_GETSEL, byref(pos_start), byref(pos_end))
                if pos_end.value > pos_start.value:
                    text_len = user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
                    text_buf = create_unicode_buffer(text_len)
                    user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXT, text_len, text_buf)
                    self._search_term = text_buf.value[pos_start.value:pos_end.value][:127]
                elif self._search_term == '':
                    self._search_term = self._saved_search_term

                # update button states
                if self._search_term:
                    user32.SendMessageW(hwnd_edit, WM_SETTEXT, 0, create_unicode_buffer(self._search_term))
                else:
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_FIND_NEXT), 0)
                if self._search_up:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_UP), BM_SETCHECK, BST_CHECKED, 0)
                else:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_DOWN), BM_SETCHECK, BST_CHECKED, 0)
                if self._match_case:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_CASE), BM_SETCHECK, BST_CHECKED, 0)
                if self._wrap_arround:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_WRAP_AROUND), BM_SETCHECK, BST_CHECKED, 0)
                if self._whole_word:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_WHOLE_WORD_ONLY), BM_SETCHECK, BST_CHECKED, 0)

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)

                if control_id == ID_EDIT_FIND:
                    if command == EN_UPDATE:
                        text_len = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_EDIT_FIND), WM_GETTEXTLENGTH, 0, 0)
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_FIND_NEXT), int(text_len > 0))

                elif command == BN_CLICKED:
                    if control_id == ID_FIND_NEXT:
                        self._match_case = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_CASE), BM_GETCHECK, 0, 0)
                        self._wrap_arround = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_WRAP_AROUND), BM_GETCHECK, 0, 0)
                        self._search_up = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_UP), BM_GETCHECK, 0, 0)
                        self._whole_word = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_WHOLE_WORD_ONLY), BM_GETCHECK, 0, 0)
                        hwnd_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)
                        text_len = user32.SendMessageW(hwnd_edit, WM_GETTEXTLENGTH, 0, 0) + 1
                        text_buf = create_unicode_buffer(text_len)
                        user32.SendMessageW(hwnd_edit, WM_GETTEXT, text_len, text_buf)
                        self._search_term = text_buf.value
                        self._find()
                        user32.SetFocus(self.rich_edit.hwnd)

                    elif control_id == ID_CANCEL:
                        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)

            elif msg == WM_CLOSE:
                user32.SetFocus(self.rich_edit.hwnd)

            return FALSE

        self.dialog_find = Dialog(self, dialog_dict, _dialog_proc_find)

        with open(os.path.join(APP_DIR, 'resources', LANG, 'Dialog1541.json'), 'rb') as f:
            dialog_dict = json.loads(f.read())

        def _dialog_proc_replace(hwnd, msg, wparam, lparam):
            if msg == WM_INITDIALOG:
                # Limit search and replace input to 127 chars
                hwnd_search_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)
                user32.SendMessageW(hwnd_search_edit, EM_SETLIMITTEXT, 127, 0)

                hwnd_replace_edit = user32.GetDlgItem(hwnd, ID_EDIT_REPLACE)
                user32.SendMessageW(hwnd_replace_edit, EM_SETLIMITTEXT, 127, 0)

                # check if something is selected
                pos_start, pos_end = DWORD(), DWORD()
                user32.SendMessageW(self.rich_edit.hwnd, EM_GETSEL, byref(pos_start), byref(pos_end))
                if pos_end.value > pos_start.value:
                    text_len = user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
                    text_buf = create_unicode_buffer(text_len)
                    user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXT, text_len, text_buf)
                    self._search_term = text_buf.value[pos_start.value:pos_end.value][:127]
                # update button states
                if self._search_term:
                    user32.SendMessageW(hwnd_search_edit, WM_SETTEXT, 0, create_unicode_buffer(self._search_term))
                else:
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_FIND_NEXT), 0)
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE), 0)
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE_ALL), 0)
                if self._replace_term:
                    user32.SendMessageW(hwnd_replace_edit, WM_SETTEXT, 0, create_unicode_buffer(self._search_term))
                if self._match_case:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_CASE), BM_SETCHECK, BST_CHECKED, 0)
                if self._wrap_arround:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_WRAP_AROUND), BM_SETCHECK, BST_CHECKED, 0)

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)

                if control_id == ID_EDIT_FIND:
                    if command == EN_UPDATE:
                        text_len = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_EDIT_FIND), WM_GETTEXTLENGTH, 0, 0)
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_FIND_NEXT), int(text_len > 0))
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE), int(text_len > 0))
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE_ALL), int(text_len > 0))

                elif command == BN_CLICKED:
                    if control_id in (ID_FIND_NEXT, ID_REPLACE, ID_REPLACE_ALL):
                        self._match_case = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_CASE), BM_GETCHECK, 0, 0)
                        self._wrap_arround = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_WRAP_AROUND), BM_GETCHECK, 0, 0)

                        hwnd_search_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)
                        text_len = user32.SendMessageW(hwnd_search_edit, WM_GETTEXTLENGTH, 0, 0) + 1
                        text_buf = create_unicode_buffer(text_len)
                        user32.SendMessageW(hwnd_search_edit, WM_GETTEXT, text_len, text_buf)
                        self._search_term = text_buf.value

                        hwnd_replace_edit = user32.GetDlgItem(hwnd, ID_EDIT_REPLACE)
                        text_len = user32.SendMessageW(hwnd_replace_edit, WM_GETTEXTLENGTH, 0, 0) + 1
                        text_buf = create_unicode_buffer(text_len)
                        user32.SendMessageW(hwnd_replace_edit, WM_GETTEXT, text_len, text_buf)
                        self._replace_term = text_buf.value

                        if control_id == ID_FIND_NEXT:
                            self._find()
                        elif control_id == ID_REPLACE:
                            self._replace()
                        elif control_id == ID_REPLACE_ALL:
                            self._replace_all()
                        user32.SetFocus(self.rich_edit.hwnd)

                    elif control_id == ID_CANCEL:
                        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)

            elif msg == WM_CLOSE:
                user32.SetFocus(self.rich_edit.hwnd)

            return FALSE

        self.dialog_replace = Dialog(self, dialog_dict, _dialog_proc_replace)

    ########################################
    #
    ########################################
    def _show_popupmenu(self, hmenu, x=None, y=None):
        if x is None:
            pt = POINT()
            user32.GetCursorPos(byref(pt))
            x, y = pt.x, pt.y
        user32.TrackPopupMenuEx(hmenu, TPM_LEFTBUTTON, x, y, self.hwnd, 0)
        user32.PostMessageW(self.hwnd, WM_NULL, 0, 0)

    ########################################
    # Tries to detect the encoding of the specified bytes
    ########################################
    def _detect_encoding(self, data):
        if len(data) == 0:
            return UTF_8, False
        bom = self._get_bom(data)
        if bom:
            return bom, True
        if data[0] == 0:
            return UTF_16_BE, False
        if len(data) > 1 and data[1] == 0:
            return UTF_16_LE, False
        if self._is_utf8(data):
            return UTF_8, False
        return ANSI, False

    ########################################
    # Checks bytes for Byte-Order-Mark (BOM), returns BOM type or None
    ########################################
    def _get_bom(self, data):
    	if len(data) >= 3 and data[:3] == b'\xEF\xBB\xBF':
    		return UTF_8_BOM
    	elif len(data) >= 2:
            if data[:2] == b'\xFF\xFE':
                return UTF_16_LE
            elif data[:2] == b'\xFE\xFF':
                return UTF_16_BE

    ########################################
    # Checks bytes for invalid UTF-8 sequences
    # Notice: since ASCII is a UTF-8 subset, function also returns True for pure ASCII data
    ########################################
    def _is_utf8(self, data):
    	data_len = len(data)
    	i = -1
    	while True:
    		i += 1
    		if i >= data_len:
    		    break
    		o = data[i]
    		if (o < 128):
    		    continue
    		elif o & 224 == 192 and o > 193:
    		    n = 1
    		elif o & 240 == 224:
    		    n = 2
    		elif o & 248 == 240 and o < 245:
    		    n = 3
    		else:
    		    return False  # invalid UTF-8 sequence
    		for c in range(n):
    			i += 1
    			if i > data_len:
    			    return False  # invalid UTF-8 sequence
    			if data[i] & 192 != 128:
    			    return False  # invalid UTF-8 sequence
    	return True

    ########################################
    #
    ########################################
    def _load_file(self, filename):
        ext = os.path.splitext(filename)[1].lower()
        if ext not in NATIVE_DOC_EXTENSIONS:
            tmp_rtf_file = os.path.join(os.environ['TMP'], '~tmp.rtf')
            if os.path.isfile(tmp_rtf_file):
                os.unlink(tmp_rtf_file)
            args = ['pandoc.exe', '--standalone', filename, '--resource-path', os.path.dirname(filename),
                    '-o', tmp_rtf_file]
            res = subprocess.run(args, shell=True, capture_output=True)
            if res.returncode != 0 or not os.path.isfile(tmp_rtf_file):
                self.statusbar.set_text(_('Unsupported file format'))
                print(res.stderr.decode())
                return False
            with open(tmp_rtf_file, 'rb') as f:
                data = f.read()
            os.unlink(tmp_rtf_file)
        else:
            try:
                with open(filename, 'rb') as f:
                    data = f.read()
            except:
                self.show_message_box(
                        _('FILE_NOT_FOUND').format(filename),
                        APP_NAME)
                user32.SetFocus(self.rich_edit.hwnd)
                return False

        if data.startswith(b'{\\rtf1\\'):
            self._encoding_id, has_bom = ANSI, False
        else:
            self._encoding_id, has_bom = self._detect_encoding(data)

        if has_bom and self._encoding_id in (UTF_16_LE, UTF_16_BE):
            data = data[2:]

        text_buf = create_unicode_buffer(data.decode(ENCODINGS[self._encoding_id]))
        st = SETTEXTEX(0, CP_UNICODE)
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETTEXTEX, byref(st), text_buf)

        if self.is_dark:
            cf = CHARFORMATW()
            cf.dwMask = CFM_COLOR
            cf.crTextColor = COLORREF(0xffffff)
            cf.dwEffects = 0
            user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_ALL, byref(cf))

        self._update_formatting()
        self._clear_undo_buffer()
        self._is_dirty = False
        self._filename = filename

        self.set_window_text(self._get_caption())

        user32.SetFocus(self.rich_edit.hwnd)

        if filename in self._recent_files:
            self._recent_files.remove(filename)
        self._recent_files.insert(0, filename)
        self._update_recent()
        return True

    ########################################
    # TODO
    ########################################
    def _save_file(self, filename):
        ext = os.path.splitext(filename)[1].lower()

        if ext == '.txt':
            res = self.show_message_box(
                    _('WARN_TXT').format(self._filename if self._filename else _('Document')),
                    APP_NAME, MB_ICONWARNING | MB_YESNO)
            if res != IDYES:
                return

            text_len = user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
            text_buf = create_unicode_buffer(text_len)
            user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXT, text_len, text_buf)

            with open(filename, 'w', newline='') as f:
                f.write(text_buf.value)
            return True

        # get RTF
        rtf_chunks = []
        def _editstreamcallback(dwCookie, pbBuff, cb, pcb):
            rtf_chunks.append(ctypes.string_at(pbBuff, cb))
            return 0
        es = EDITSTREAM()
        es.dwError = 0
        es.pfnCallback = EDITSTREAMCALLBACK(_editstreamcallback)
        user32.SendMessageW(self.rich_edit.hwnd, EM_STREAMOUT, SF_RTF, byref(es))
        rtf = b''.join(rtf_chunks)
        if self.is_dark:
            rtf = rtf.replace(b'red255\\green255\\blue255', b'red0\\green0\\blue0')

        if ext not in NATIVE_DOC_EXTENSIONS:
            tmp_rtf_file = os.path.join(os.environ['TMP'], '~tmp.rtf')
            with open(tmp_rtf_file, 'wb') as f:
                f.write(rtf)

            args = ['pandoc.exe', tmp_rtf_file, '-o', os.path.basename(filename)]
            if ext in ('.md', '.htm', '.html'):
                image_dir = os.path.splitext(os.path.basename(filename))[0] + '_images'
                args += ['--extract-media', f'./{image_dir}']
            else:
                args += ['--standalone']

            target_dir = os.path.dirname(os.path.abspath(filename))
            res = subprocess.run(args, cwd=target_dir, shell=True, capture_output=True)
            os.unlink(tmp_rtf_file)
            if res.returncode != 0 or not os.path.isfile(filename):
                self.statusbar.set_text(_('Unknown error'))
                print(res.stderr.decode())
                return False
            #return self._load_file(filename)
        else:
            with open(filename, 'wb') as f:
                f.write(rtf)
        return True

    ########################################
    #
    ########################################
    def _handle_dirty(self):
        if not self._is_dirty:
            return True
        res = self.show_message_box(
                _('SAVE_CHANGES').format(self._filename if self._filename else _('Document')),
                APP_NAME, MB_YESNOCANCEL)
        if res == IDCANCEL:
            return False
        elif res == IDNO:
            return True
        elif res == IDYES:
            return self.action_save()

    ########################################
    #
    ########################################
    def _get_caption(self):
        return '{}{} - {}'.format(
                '*' if self._is_dirty else '',
                os.path.basename(self._filename) if self._filename else _('Document'),
                APP_NAME)

    ########################################
    # Load saved state from registry
    ########################################
    def _load_state(self):
        left, top, width, height=CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT
        hkey = HKEY()
        if advapi32.RegOpenKeyW(HKEY_CURRENT_USER, f'Software\\{APP_NAME}' , byref(hkey)) == ERROR_SUCCESS:
            data = (BYTE * sizeof(DWORD))()
            cbData = DWORD(sizeof(data))
            if advapi32.RegQueryValueExW(hkey, 'fWrap', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._word_wrap = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'DarkMode', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._is_dark_mode = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'ToolBar', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._show_toolbar = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'FormatBar', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._show_formatbar = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'StatusBar', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._show_statusbar = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'iWindowPosX', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                left = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'iWindowPosY', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                top = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'iWindowPosDX', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                width = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'iWindowPosDY', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                height = cast(data, POINTER(DWORD)).contents.value
            data = (WCHAR * MAX_PATH * MAX_RECENT_FILES)()
            cbData = DWORD(sizeof(data))
            if advapi32.RegQueryValueExW(hkey, 'RecentFiles', None, None, data, byref(cbData)) == ERROR_SUCCESS:
                self._recent_files = json.loads(cast(data, LPWSTR).value)
        else:
            advapi32.RegCreateKeyW(HKEY_CURRENT_USER, f'Software\\{APP_NAME}' , byref(hkey))
        advapi32.RegCloseKey(hkey)
        return left, top, width, height

    ########################################
    # Save state to registry
    ########################################
    def _save_state(self):
        hkey = HKEY()
        if advapi32.RegOpenKeyW(HKEY_CURRENT_USER, f'Software\\{APP_NAME}' , byref(hkey)) == ERROR_SUCCESS:
            dwsize = sizeof(DWORD)
            advapi32.RegSetValueExW(hkey, 'fWrap', 0, REG_DWORD, byref(DWORD(int(self._word_wrap))), dwsize)
            advapi32.RegSetValueExW(hkey, 'DarkMode', 0, REG_DWORD, byref(DWORD(int(self.is_dark))), dwsize)
            advapi32.RegSetValueExW(hkey, 'ToolBar', 0, REG_DWORD, byref(DWORD(int(self._show_toolbar))), dwsize)
            advapi32.RegSetValueExW(hkey, 'FormatBar', 0, REG_DWORD, byref(DWORD(int(self._show_formatbar))), dwsize)
            advapi32.RegSetValueExW(hkey, 'StatusBar', 0, REG_DWORD, byref(DWORD(int(self._show_statusbar))), dwsize)
            self.show(SW_RESTORE)
            rc = self.get_window_rect()
            advapi32.RegSetValueExW(hkey, 'iWindowPosX', 0, REG_DWORD, byref(DWORD(rc.left)), dwsize)
            advapi32.RegSetValueExW(hkey, 'iWindowPosY', 0, REG_DWORD, byref(DWORD(rc.top)), dwsize)
            advapi32.RegSetValueExW(hkey, 'iWindowPosDX', 0, REG_DWORD, byref(DWORD(rc.right - rc.left)), dwsize)
            advapi32.RegSetValueExW(hkey, 'iWindowPosDY', 0, REG_DWORD, byref(DWORD(rc.bottom - rc.top)), dwsize)
            buf = create_unicode_buffer(json.dumps(self._recent_files[:MAX_RECENT_FILES]))
            advapi32.RegSetValueExW(hkey, 'RecentFiles', 0, REG_SZ, buf, sizeof(buf))
            advapi32.RegCloseKey(hkey)

    ########################################
    #
    ########################################
    def _find(self, search_up=None):

        sel_start_pos, sel_end_pos = DWORD(), DWORD()
        user32.SendMessageW(self.rich_edit.hwnd, EM_GETSEL, byref(sel_start_pos), byref(sel_end_pos))

        # If the cpMin and cpMax members are equal, the range is empty. The range includes everything if cpMin is 0 and cpMax is –1.
        ftw = FINDTEXTW()
        ftw.lpstrText = self._search_term
        if self._search_up:
            ftw.chrg.cpMin = sel_start_pos.value
            ftw.chrg.cpMax = 0
        else:
            ftw.chrg.cpMin = sel_end_pos.value
            ftw.chrg.cpMax = -1

        flags = 0
        if not self._search_up:
            flags |= FR_DOWN
        if self._match_case:
            flags |= FR_MATCHCASE
        if self._whole_word:
            flags |= FR_WHOLEWORD

        pos = user32.SendMessageW(self.rich_edit.hwnd, EM_FINDTEXTW, flags, byref(ftw))  # -1 = not found

        if pos > -1:
            user32.SendMessageW(self.rich_edit.hwnd, EM_SETSEL, pos, pos + len(self._search_term))
        elif self._wrap_arround:
            if self._search_up:
                ftw.chrg.cpMin = user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXTLENGTH, 0, 0)
                ftw.chrg.cpMax = 0
            else:
                ftw.chrg.cpMin = 0
                ftw.chrg.cpMax = sel_start_pos.value
            pos = user32.SendMessageW(self.rich_edit.hwnd, EM_FINDTEXTW, flags, byref(ftw))
            if pos > -1:
                user32.SendMessageW(self.rich_edit.hwnd, EM_SETSEL, pos, pos + len(self._search_term))
        if pos < 0:
            self.show_message_box(_("FIND_FINISHED").format(APP_NAME), APP_NAME, dialog_width=263)
        return pos > -1

    ########################################
    # if current selection is search term, replace and find next result.
    # otherwise just find next result.
    ########################################
    def _replace(self):
        sel_start_pos, sel_end_pos = DWORD(), DWORD()
        user32.SendMessageW(self.rich_edit.hwnd, EM_GETSEL, byref(sel_start_pos), byref(sel_end_pos))
        buf = create_unicode_buffer(sel_end_pos.value - sel_start_pos.value)
        user32.SendMessageW(self.rich_edit.hwnd, EM_GETSELTEXT, 0, buf)
        if buf.value == self._search_term:
            user32.SendMessageW(self.rich_edit.hwnd, EM_REPLACESEL, TRUE, create_unicode_buffer(self._replace_term))
        self._find()

    ########################################
    #
    ########################################
    def _replace_all(self):
        while self._find():
            user32.SendMessageW(self.rich_edit.hwnd, EM_REPLACESEL, TRUE, create_unicode_buffer(self._replace_term))

    ########################################
    #
    ########################################
    def _check_if_has_text(self):
        text_len = user32.SendMessageW(self.rich_edit.hwnd, WM_GETTEXTLENGTH, 0, 0)
        # enable/disable menu items
        flag = MF_BYCOMMAND | (MF_ENABLED if text_len > 0 else MF_GRAYED)
        for item_id in (IDM_FIND, IDM_FIND_NEXT, IDM_REPLACE):
            user32.EnableMenuItem(self.hmenu, item_id, flag)

    ########################################
    #
    ########################################
    def _check_if_text_selected(self):
        char_pos_start, char_pos_end = DWORD(), DWORD()
        user32.SendMessageW(self.rich_edit.hwnd, EM_GETSEL, byref(char_pos_start), byref(char_pos_end))
        # enable/disable menu items
        flag = MF_BYCOMMAND | (MF_ENABLED if char_pos_end.value > char_pos_start.value else MF_GRAYED)
        for item_id in (IDM_CUT, IDM_COPY, IDM_CLEAR):
            user32.EnableMenuItem(self.hmenu, item_id, flag)

    ########################################
    #
    ########################################
    def _clear_undo_buffer(self):
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETUNDOLIMIT, 0, 0)
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETUNDOLIMIT, 100, 0)

    ########################################
    #
    ########################################
    def _apply_bullet_style(self):
        pf = PARAFORMAT2()
        pf.dwMask = PFM_NUMBERING | PFM_NUMBERINGSTART
        pf.wNumbering = self._current_bullet_style
        pf.wNumberingStart = 1
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETPARAFORMAT, 0, byref(pf))
        # update toolbar button
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, IDM_BULLET_STYLE,
                TBSTATE_ENABLED | TBSTATE_CHECKED if self._current_bullet_style else TBSTATE_ENABLED)
        # update menu
        for i in range(PFN_UCROMAN + 1):
            check_state = MF_CHECKED if i == self._current_bullet_style else MF_UNCHECKED
            user32.CheckMenuItem(self.hmenu, IDM_BULLET_STYLE_NONE + i, MF_BYCOMMAND | check_state)

    ########################################
    #
    ########################################
    def action_new(self):
        if not self._handle_dirty():
            user32.SetFocus(self.rich_edit.hwnd)
            return
        self._filename = None
        self._is_dirty = False
        user32.SendMessageW(self.rich_edit.hwnd, WM_SETTEXT, 0, create_unicode_buffer(''))
        self.set_window_text(self._get_caption())
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_open(self, filename=None):
        if not self._handle_dirty():
            return
        if filename is None:
            if HAS_PANDOC:
                filename = self.get_open_filename(_('Open'), '.bmp',
                        _('All Documents') + '\0*.rtf;*.txt;*.md;*.docx;*.odt;*.htm;*.html;*.epub\0' +
                        'RTF (*.rtf)\0*.rtf\0' +
                        'TXT (*.txt)\0*.txt\0' +
                        'Markdown (*.md)\0*.md\0' +
                        'DOCX (*.docx)\0*.docx\0' +
                        'ODT (*.odt)\0*.odt\0' +
                        'HTML (*.htm;*html)\0*.htm;*.html\0' +
                        'EPUB (*.epub)\0*.epub\0' +
                        '\0'
                        )
            else:
                filename = self.get_open_filename(_('Open'), '.rtf',
                        _('Rich Text Format') + ' (*.rtf)\0*.rtf\0' + _('Text Documents') + ' (*.txt)\0*.txt\0\0')
        if filename:
            self._load_file(filename)

    ########################################
    #
    ########################################
    def action_save(self):
        if not self._filename:
            return self.action_save_as()
        if not self._save_file(self._filename):
            return False
        self._is_dirty = False
        self.set_window_text(self._get_caption())
        return True

    ########################################
    #
    ########################################
    def action_save_as(self):
        if HAS_PANDOC:
            filename = self.get_save_filename(_('Save As'), '.rtf',
                    'RTF (*.rtf)\0*.rtf\0' +
                    'TXT (*.txt)\0*.txt\0' +
                    'Markdown (*.md)\0*.md\0' +
                    'DOCX (*.docx)\0*.docx\0' +
                    'ODT (*.odt)\0*.odt\0' +
                    'HTML (*.htm;*html)\0*.htm;*.html\0' +
                    'EPUB (*.epub)\0*.epub\0' +
                    '\0',
                    self._filename if self._filename else ''
                    )
        else:
            filename = self.get_save_filename(_('Save As'), '.rtf',
                    _('Rich Text Format') + ' (*.rtf)\0*.rtf\0' + _('Text Documents') + ' (*.txt)\0*.txt\0\0',
                    self._filename if self._filename else '')

        if not filename or not self._save_file(filename):
            return False
        self._filename = filename
        self._is_dirty = False
        self.set_window_text(self._get_caption())
        if filename in self._recent_files:
            self._recent_files.remove(filename)
        self._recent_files.insert(0, filename)
        self._update_recent()
        return True

    ########################################
    #
    ########################################
    def action_print(self):
        pd = PRINTDLGW()
        pd.Flags = PD_RETURNDC
        ok = comdlg32.PrintDlgW(byref(pd))
        if not ok:
            return False

        hdc = pd.hDC
        di = DOCINFOW()

        di.lpszDocName = self._filename if self._filename else _('Document')

        job_id = gdi32.StartDocW(hdc, byref(di))
        if job_id <= 0:
            return False

        fr = FORMATRANGE()
        fr.hdc       = hdc
        fr.hdcTarget = hdc

        # Set page rect to physical page size in twips.
        fr.rcPage.top    = 0
        fr.rcPage.left   = 0
        fr.rcPage.right = cm_to_twips(self._print_paper_size.x / 1000)
        fr.rcPage.bottom = cm_to_twips(self._print_paper_size.y / 1000)

        fr.rc.left = cm_to_twips(self._print_margins.left / 1000)
        fr.rc.right = fr.rcPage.right - cm_to_twips(self._print_margins.right / 1000)
        fr.rc.top = cm_to_twips(self._print_margins.top / 1000)
        fr.rc.bottom = fr.rcPage.bottom - cm_to_twips(self._print_margins.bottom / 1000)

        # Select the entire contents.
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETSEL, 0, -1)
        # Get the selection into a CHARRANGE.
        user32.SendMessageW(self.rich_edit.hwnd, EM_EXGETSEL, 0, byref(fr.chrg))

        fSuccess = True

        # print successive pages.
        while fr.chrg.cpMin < fr.chrg.cpMax and fSuccess:
            fSuccess = gdi32.StartPage(hdc) > 0
            if not fSuccess:
                break
            cpMin = user32.SendMessageW(self.rich_edit.hwnd, EM_FORMATRANGE, TRUE, byref(fr))
            if cpMin <= fr.chrg.cpMin:
                fSuccess = False
                break
            fr.chrg.cpMin = cpMin
            fSuccess = gdi32.EndPage(hdc) > 0

        user32.SendMessageW(self.rich_edit.hwnd, EM_FORMATRANGE, FALSE, 0)

        if fSuccess:
            gdi32.EndDoc(hdc)
        else:
            gdi32.AbortDoc(hdc)

        return fSuccess

    ########################################
    #
    ########################################
    def action_page_setup(self):
        psd = PAGESETUPDLGW()
        psd.hwndOwner = self.hwnd
        psd.Flags = PSD_INHUNDREDTHSOFMILLIMETERS | PSD_MARGINS
        psd.ptPaperSize = self._print_paper_size
        psd.rtMargin = self._print_margins
        ok = comdlg32.PageSetupDlgW(byref(psd))
        if ok:
            self._print_paper_size = psd.ptPaperSize
            self._print_margins = psd.rtMargin

    ########################################
    #
    ########################################
    def action_exit(self, *args):
        if not self._handle_dirty():
            user32.SetFocus(self.rich_edit.hwnd)
            return 1
        self._save_state()
        super().quit()

    ########################################
    #
    ########################################
    def action_undo(self):
        user32.SendMessageW(self.rich_edit.hwnd, EM_UNDO, 0, 0)

    ########################################
    #
    ########################################
    def action_redo(self):
        user32.SendMessageW(self.rich_edit.hwnd, EM_REDO, 0, 0)

    ########################################
    #
    ########################################
    def action_cut(self):
        user32.SendMessageW(self.rich_edit.hwnd, WM_CUT, 0, 0)

    ########################################
    #
    ########################################
    def action_copy(self):
        if self.is_dark:
            cf = CHARFORMATW()
            cf.dwMask = CFM_COLOR
            cf.crTextColor = COLORREF(0x000000)
            user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_ALL, byref(cf))
            user32.SendMessageW(self.rich_edit.hwnd, WM_COPY, 0, 0)
            cf.crTextColor = COLORREF(0xffffff)
            user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_ALL, byref(cf))
        else:
            user32.SendMessageW(self.rich_edit.hwnd, WM_COPY, 0, 0)

    ########################################
    #
    ########################################
    def action_paste(self):
        user32.SendMessageW(self.rich_edit.hwnd, WM_PASTE, 0, 0)
        if self.is_dark:
            cf = CHARFORMATW()
            cf.dwMask = CFM_COLOR
            cf.crTextColor = COLORREF(0xffffff)
            user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_ALL, byref(cf))

    ########################################
    #
    ########################################
    def action_clear(self):
        user32.SendMessageW(self.rich_edit.hwnd, WM_CLEAR, 0, 0)

    ########################################
    #
    ########################################
    def action_select_all(self):
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETSEL, 0, -1)

    ########################################
    # WS_HSCROLL can't be changed after edit control was created, so the control needs to be replaced.
    # (MS Notepad does the same thing).
    ########################################
    def action_toggle_word_wrap(self):
        self._word_wrap = not self._word_wrap
        user32.CheckMenuItem(self.hmenu, IDM_WORD_WRAP,
            MF_BYCOMMAND | (MF_CHECKED if self._word_wrap else MF_UNCHECKED))
        style = user32.GetWindowLongA(self.rich_edit.hwnd, GWL_STYLE)
        if self._word_wrap:
            style &= ~WS_HSCROLL
        else:
            style |= WS_HSCROLL
        user32.SetWindowLongA(self.rich_edit.hwnd, GWL_STYLE, style)
        user32.SetWindowPos(self.rich_edit.hwnd, 0, 0,0, 0,0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETTARGETDEVICE, 0, int(not self._word_wrap))

    ########################################
    #
    ########################################
    def action_toggle_dark(self):
        was_dirty = self._is_dirty
        self.is_dark = not self.is_dark
        user32.CheckMenuItem(self.hmenu, IDM_DARK_MODE, MF_BYCOMMAND | (MF_CHECKED if self.is_dark else MF_UNCHECKED))
        self.apply_theme(self.is_dark)
        user32.SetWindowPos(self.hwnd, 0, 0,0, 0,0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
        if not was_dirty:
            self._is_dirty = False
            self.set_window_text(self._get_caption())

    ########################################
    #
    ########################################
    def action_toggle_toolbar(self):
        self._show_toolbar = not self._show_toolbar
        user32.CheckMenuItem(self.hmenu, IDM_TOOLBAR, MF_BYCOMMAND | (MF_CHECKED if self._show_toolbar else MF_UNCHECKED))
        user32.SendMessageW(self.rebar.hwnd, RB_SHOWBAND, 0, int(self._show_toolbar))

    ########################################
    #
    ########################################
    def action_toggle_formatbar(self):
        self._show_formatbar = not self._show_formatbar
        user32.CheckMenuItem(self.hmenu, IDM_FORMAT_BAR, MF_BYCOMMAND | (MF_CHECKED if self._show_formatbar else MF_UNCHECKED))
        user32.SendMessageW(self.rebar.hwnd, RB_SHOWBAND, 1, int(self._show_formatbar))

    ########################################
    #
    ########################################
    def action_toggle_statusbar(self):
        self._show_statusbar = not self._show_statusbar
        user32.CheckMenuItem(self.hmenu, IDM_STATUS_BAR, MF_BYCOMMAND | (MF_CHECKED if self._show_statusbar else MF_UNCHECKED))
        self.statusbar.show(SW_SHOW if self._show_statusbar else SW_HIDE)
        rc = self.get_client_rect()
        width, height = rc.right - rc.left, rc.bottom - rc.top
        self.rich_edit.set_window_pos(0, self.rebar_height + 2, width, height - self.rebar_height -
                (self.statusbar.height if self._show_statusbar else 0) - 2, SWP_NOSIZE)

    ########################################
    #
    ########################################
    def action_insert_date_time(self):
        date_str = time.strftime('%X %x')
        user32.SendMessageW(self.rich_edit.hwnd, EM_REPLACESEL, 1, create_unicode_buffer(date_str))

    ########################################
    #
    ########################################
    def action_font(self):
        font = self.show_font_dialog(*self._current_font)
        if font is None:
            return
        self._current_font = font
        cf = CHARFORMATW()
        cf.dwMask = CFM_FACE | CFM_SIZE | CFM_BOLD | CFM_ITALIC
        cf.szFaceName = font[0]
        cf.yHeight = font[1] * 20  # Character height, in twips (1/1440 of an inch or 1/20 of a printer's point).
        cf.dwEffects = (CFE_BOLD if font[2] > 400 else 0) | (CFE_ITALIC if font[3] else 0)
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, byref(cf))
        user32.SendMessageW(self.combobox_font.hwnd, CB_SETCURSEL, self._fonts.index(font[0]), 0)
        if font[1] in self._font_sizes:
            user32.SendMessageW(self.combobox_fontsize.hwnd, CB_SETCURSEL, self._font_sizes.index(font[1]), 0)
        else:
            idx = sorted(self._font_sizes + [font[1]]).index(font[1])
            user32.SendMessageW(self.combobox_fontsize.hwnd, CB_INSERTSTRING, idx, str(font[1]))
            user32.SendMessageW(self.combobox_fontsize.hwnd, CB_SETCURSEL, idx, 0)
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_zoom_in(self):
        self._zoom += 10
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETZOOM, self._zoom, 100)
        self.statusbar.set_text(f'{self._zoom}%', STATUSBAR_PART_ZOOM)

    ########################################
    #
    ########################################
    def action_zoom_out(self):
        if self._zoom <= 10:
            return
        self._zoom -= 10
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETZOOM, self._zoom, 100)
        self.statusbar.set_text(f'{self._zoom}%', STATUSBAR_PART_ZOOM)

    ########################################
    #
    ########################################
    def action_restore_default_zoom(self):
        self._zoom = 100
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETZOOM, 1, 1)
        self.statusbar.set_text('100%', STATUSBAR_PART_ZOOM)

    ########################################
    #
    ########################################
    def action_about_wordpad(self):
        self.show_message_box(
                _('ABOUT_TEXT').format(APP_NAME, APP_VERSION, APP_COPYRIGHT),
                _('ABOUT_CAPTION').format(APP_NAME))
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_about_windows(self):
        shell32.ShellAboutW(self.hwnd, create_unicode_buffer('Windows'), None, None)

    ########################################
    #
    ########################################
    def action_set_font(self):
        idx = user32.SendMessageW(self.combobox_font.hwnd, CB_GETCURSEL, 0, 0)
        fn = self._fonts[idx]
        cf = CHARFORMATW()
        cf.dwMask = CFM_FACE
        cf.szFaceName = fn
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, byref(cf))
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_set_fontsize(self):
        idx = user32.SendMessageW(self.combobox_fontsize.hwnd, CB_GETCURSEL, 0, 0)
        fs = self._font_sizes[idx]
        cf = CHARFORMATW()
        cf.dwMask = CFM_SIZE
        cf.yHeight = fs * 20
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, byref(cf))
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_toggle_bold(self):
        cf = CHARFORMATW()
        cf.dwMask = CFM_BOLD
        s = user32.SendMessageW(self.toolbar_formatting.hwnd, TB_GETSTATE, IDM_BOLD, 0)
        if s & TBSTATE_CHECKED:
            s &= ~TBSTATE_CHECKED
            cf.dwEffects = 0
        else:
            s |= TBSTATE_CHECKED
            cf.dwEffects = CFE_BOLD
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, IDM_BOLD, s)
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, byref(cf))
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_toggle_italic(self):
        cf = CHARFORMATW()
        cf.dwMask = CFM_ITALIC
        s = user32.SendMessageW(self.toolbar_formatting.hwnd, TB_GETSTATE, IDM_ITALIC, 0)
        if s & TBSTATE_CHECKED:
            s &= ~TBSTATE_CHECKED
            cf.dwEffects = 0
        else:
            s |= TBSTATE_CHECKED
            cf.dwEffects = CFE_ITALIC
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, IDM_ITALIC, s)
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, byref(cf))
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_toggle_underline(self):
        cf = CHARFORMATW()
        cf.dwMask = CFM_UNDERLINE
        s = user32.SendMessageW(self.toolbar_formatting.hwnd, TB_GETSTATE, IDM_UNDERLINE, 0)
        if s & TBSTATE_CHECKED:
            s &= ~TBSTATE_CHECKED
            cf.dwEffects = 0
        else:
            s |= TBSTATE_CHECKED
            cf.dwEffects = CFE_UNDERLINE
        user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, IDM_UNDERLINE, s)
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, byref(cf))
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_set_alignment(self, alignment):
        for id in (IDM_ALIGN_LEFT, IDM_ALIGN_CENTER, IDM_ALIGN_RIGHT):
            user32.SendMessageW(self.toolbar_formatting.hwnd, TB_SETSTATE, id,
                    TBSTATE_ENABLED | TBSTATE_CHECKED if id - IDM_ALIGN_LEFT + 1 == alignment else TBSTATE_ENABLED)
        pf = PARAFORMAT()
        pf.dwMask = PFM_ALIGNMENT
        pf.wAlignment = alignment
        self._current_paraformat2.wAlignment = alignment
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETPARAFORMAT, 0, byref(pf))
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_set_line_spacing(self, line_spacing):
        pf = PARAFORMAT2()
        pf.dwMask = PFM_LINESPACING
        pf.bLineSpacingRule = 5
        pf.dyLineSpacing = [20, 23, 30, 40][line_spacing - IDM_LINE_SPACING_1]
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETPARAFORMAT, 0, byref(pf))
        # update menus
        for i in range(PFN_UCROMAN + 1):
            user32.CheckMenuItem(self.hmenu, IDM_LINE_SPACING_1 + i, MF_BYCOMMAND | (MF_CHECKED if i == line_spacing else MF_UNCHECKED))
        user32.SetFocus(self.rich_edit.hwnd)

    ########################################
    #
    ########################################
    def action_toggle_bullets(self):
        self._current_bullet_style = 0 if self._current_bullet_style else self._last_bullet_style
        self._apply_bullet_style()

    ########################################
    #
    ########################################
    def action_set_bullet_style(self, bullet_style):
        self._current_bullet_style = bullet_style
        if bullet_style:
            self._last_bullet_style = bullet_style
        self._apply_bullet_style()

    ########################################
    #
    ########################################
    def action_paragraph(self):
        self.dialog_show_sync(self.dialog_paragraph)

    ########################################
    #
    ########################################
    def action_tabs(self):
        self.dialog_show_sync(self.dialog_tabs)

    ########################################
    #
    ########################################
    def action_find(self):
        if self.dialog_find.hwnd:
            user32.SetActiveWindow(self.dialog_find.hwnd)
            return
        elif self.dialog_replace.hwnd:
            user32.SendMessageW(self.dialog_replace.hwnd, WM_CLOSE, 0, 0)
        self.dialog_show_async(self.dialog_find)

    ########################################
    #
    ########################################
    def action_find_next(self):
        if self._search_term == '':
            return self.action_find()
        self._find(False)

    ########################################
    #
    ########################################
    def action_replace(self):
        if self.dialog_replace.hwnd:
            user32.SetActiveWindow(self.dialog_replace.hwnd)
            return
        elif self.dialog_find.hwnd:
            user32.SendMessageW(self.dialog_find.hwnd, WM_CLOSE, 0, 0)
        self.dialog_show_async(self.dialog_replace)

    ########################################
    # https://learn.microsoft.com/en-us/windows/win32/controls/em-insertimage
    # https://learn.microsoft.com/en-us/windows/win32/api/richedit/ns-richedit-richedit_image_parameters
    ########################################
    def action_insert_image(self, img_file=None):
        if img_file is None:
            if HAS_MAGICK:
                img_file = self.get_open_filename(_('Open'), '.bmp',
                    _('All Images') +'\0*.bmp;*.jpg;*.jpeg;*.png;*.gif;*.tif;*.tiff\0' +
                    'BMP (*.bmp)\0*.bmp\0' +
                    'JPEG (*.jpg;*jpeg)\0*.jpg;*.jpeg\0' +
                    'PNG (*.png)\0*.png\0' +
                    'GIF (*.gif)\0*.gif\0' +
                    'TIFF (*.tif;*tiff)\0*.tif;*.tiff\0' +
                    '\0'
                    )
            else:
                img_file = self.get_open_filename(_('Open'), '.bmp', 'BMP (*.bmp)\0*.bmp\0\0')
            if not img_file:
                return

        if img_file.lower().endswith('.bmp'):
            is_tmpfile = False
            bmp_file = img_file
        elif HAS_MAGICK:
            is_tmpfile = True
            bmp_file = os.path.join(os.environ['TMP'], '~tmp.bmp')
            if os.path.isfile(bmp_file):
                os.unlink(bmp_file)
            res = subprocess.run(['magick.exe', img_file, 'BMP3:' + bmp_file], shell=True, capture_output=True)
            if res.returncode != 0 or not os.path.isfile(bmp_file):
                self.statusbar.set_text(_('Unsupported file format'))
                print(res.stderr.decode())
                return
        else:
            return
        hbitmap = user32.LoadImageW(0, bmp_file, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
        if is_tmpfile:
            os.unlink(bmp_file)
        if not hbitmap:
            self.statusbar.set_text(_('Unsupported file format'))
            return

        # force new paragraph
        user32.SendMessageW(self.rich_edit.hwnd, EM_REPLACESEL, TRUE, create_unicode_buffer('\r\n'))

        # put bitmap into clipboard
        user32.OpenClipboard(self.hwnd)
        user32.EmptyClipboard()
        user32.SetClipboardData(CF_BITMAP, hbitmap)
        user32.CloseClipboard()

        # paste bitmap
        user32.SendMessageW(self.rich_edit.hwnd, WM_PASTE, 0, 0)

        # move cursor after image
        pos_from, pos_to = DWORD(), DWORD()
        user32.SendMessageW(self.rich_edit.hwnd, EM_GETSEL, byref(pos_from), byref(pos_to))
        user32.SendMessageW(self.rich_edit.hwnd, EM_SETSEL, pos_to.value + 1, pos_to.value + 1)

        # force new paragraph
        user32.SendMessageW(self.rich_edit.hwnd, EM_REPLACESEL, TRUE, create_unicode_buffer('\r\n'))

        # clear clipboard
        user32.OpenClipboard(self.hwnd)
        user32.EmptyClipboard()
        user32.CloseClipboard()

        user32.SetFocus(self.rich_edit.hwnd)


if __name__ == "__main__":
    import traceback
    sys.excepthook = traceback.print_exception

    app = App(sys.argv[1:])
    sys.exit(app.run())

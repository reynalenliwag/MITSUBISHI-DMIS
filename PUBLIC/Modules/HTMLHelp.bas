Attribute VB_Name = "modHTMLHelp"
'**********************************************************************
'**modWinHelp
'**(c) 1997-1999, Shadow Mountain Tech. All rights reserved.
'**
'**HTMLHelp API function declarations and constant definitions.
'**
'**HTML constants and declarations were extracted from
'**clsHTMLHelp.cls by David Liske.
'**
'**********************************************************************
Option Explicit

' HTML Help Constants
Public Const HH_DISPLAY_TOPIC = &H0            '  WinHelp equivalent
Public Const HH_DISPLAY_TOC = &H1              '  WinHelp equivalent
Public Const HH_DISPLAY_INDEX = &H2            '  WinHelp equivalent
Public Const HH_DISPLAY_SEARCH = &H3           '  WinHelp equivalent
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_SYNC = &H9
Public Const HH_ADD_NAV_UI = &HA               ' not currently implemented
Public Const HH_ADD_BUTTON = &HB               ' not currently implemented
Public Const HH_GETBROWSER_APP = &HC           ' not currently implemented
Public Const HH_KEYWORD_LOOKUP = &HD           '  WinHelp equivalent
Public Const HH_DISPLAY_TEXT_POPUP = &HE       ' display string resource id
                                                ' or text in a popup window
                                                ' value in dwData
Public Const HH_HELP_CONTEXT = &HF             '  display mapped numeric
Public Const HH_CLOSE_ALL = &H12               '  WinHelp equivalent
Public Const HH_ALINK_LOOKUP = &H13            '  ALink version of
                                                '  HH_KEYWORD_LOOKUP
Public Const HH_SET_GUID = &H1A                ' For Microsoft Installer -- dwData is a pointer to the GUID string

' HTML Help window constants. These are also used
' in the window definitions in HHP files
Public Const HHWIN_PROP_ONTOP = &H2              ' Top-most window (not currently implemented)
Public Const HHWIN_PROP_NOTITLEBAR = &H4         ' no title bar
Public Const HHWIN_PROP_NODEF_STYLES = &H8       ' no default window styles (only HH_WINTYPE.dwStyles)
Public Const HHWIN_PROP_NODEF_EXSTYLES = &H10    ' no default extended window styles (only HH_WINTYPE.dwExStyles)
Public Const HHWIN_PROP_TRI_PANE = &H20          ' use a tri-pane window
Public Const HHWIN_PROP_NOTB_TEXT = &H40         ' no text on toolbar buttons
Public Const HHWIN_PROP_POST_QUIT = &H80         ' post WM_QUIT message when window closes
Public Const HHWIN_PROP_AUTO_SYNC = &H100        ' automatically ssync contents and index
Public Const HHWIN_PROP_TRACKING = &H200         ' send tracking notification messages
Public Const HHWIN_PROP_TAB_SEARCH = &H400       ' include search tab in navigation pane
Public Const HHWIN_PROP_TAB_HISTORY = &H800      ' include history tab in navigation pane
Public Const HHWIN_PROP_TAB_BOOKMARKS = &H1000   ' include bookmark tab in navigation pane
Public Const HHWIN_PROP_CHANGE_TITLE = &H2000    ' Put current HTML title in title bar
Public Const HHWIN_PROP_NAV_ONLY_WIN = &H4000    ' Only display the navigation window
Public Const HHWIN_PROP_NO_TOOLBAR = &H8000      ' Don't display a toolbar
Public Const HHWIN_PROP_MENU = &H10000           ' Menu
Public Const HHWIN_PROP_TAB_ADVSEARCH = &H20000  ' Advanced FTS UI.
Public Const HHWIN_PROP_USER_POS = &H40000       ' After initial creation, user controls window size/position

Public Const HHWIN_PARAM_PROPERTIES = &H2        ' valid fsWinProperties
Public Const HHWIN_PARAM_STYLES = &H4            ' valid dwStyles
Public Const HHWIN_PARAM_EXSTYLES = &H8          ' valid dwExStyles
Public Const HHWIN_PARAM_RECT = &H10             ' valid rcWindowPos
Public Const HHWIN_PARAM_NAV_WIDTH = &H20        ' valid iNavWidth
Public Const HHWIN_PARAM_SHOWSTATE = &H40        ' valid nShowState
Public Const HHWIN_PARAM_INFOTYPES = &H80        ' valid apInfoTypes
Public Const HHWIN_PARAM_TB_FLAGS = &H100        ' valid fsToolBarFlags
Public Const HHWIN_PARAM_EXPANSION = &H200       ' valid fNotExpanded
Public Const HHWIN_PARAM_TABPOS = &H400          ' valid tabpos
Public Const HHWIN_PARAM_TABORDER = &H800        ' valid taborder
Public Const HHWIN_PARAM_HISTORY_COUNT = &H1000  ' valid cHistory
Public Const HHWIN_PARAM_CUR_TAB = &H2000        ' valid curNavType

Public Const HHWIN_BUTTON_EXPAND = &H2           ' Expand/contract button
Public Const HHWIN_BUTTON_BACK = &H4             ' Back button
Public Const HHWIN_BUTTON_FORWARD = &H8          ' Forward button
Public Const HHWIN_BUTTON_STOP = &H10            ' Stop button
Public Const HHWIN_BUTTON_REFRESH = &H20         ' Refresh button
Public Const HHWIN_BUTTON_HOME = &H40            ' Home button
Public Const HHWIN_BUTTON_BROWSE_FWD = &H80      ' not implemented
Public Const HHWIN_BUTTON_BROWSE_BCK = &H100     ' not implemented
Public Const HHWIN_BUTTON_NOTES = &H200          ' not implemented
Public Const HHWIN_BUTTON_CONTENTS = &H400       ' not implemented
Public Const HHWIN_BUTTON_SYNC = &H800           ' Locate button
Public Const HHWIN_BUTTON_OPTIONS = &H1000       ' Options button
Public Const HHWIN_BUTTON_PRINT = &H2000         ' Print button
Public Const HHWIN_BUTTON_INDEX = &H4000         ' not implemented
Public Const HHWIN_BUTTON_SEARCH = &H8000        ' not implemented
Public Const HHWIN_BUTTON_HISTORY = &H10000      ' not implemented
Public Const HHWIN_BUTTON_BOOKMARKS = &H20000    ' not implemented
Public Const HHWIN_BUTTON_JUMP1 = &H40000        ' Jump1 button
Public Const HHWIN_BUTTON_JUMP2 = &H80000        ' Jump2 button
Public Const HHWIN_BUTTON_ZOOM = &H100000        ' Font sizing button
Public Const HHWIN_BUTTON_TOC_NEXT = &H200000    ' Browse next TOC topic button
Public Const HHWIN_BUTTON_TOC_PREV = &H400000    ' Browse previous TOC topic button

' Default button set
Public Const HHWIN_DEF_BUTTONS = (HHWIN_BUTTON_EXPAND Or HHWIN_BUTTON_BACK Or HHWIN_BUTTON_OPTIONS Or HHWIN_BUTTON_PRINT)

' Button IDs
Public Const IDTB_EXPAND = 200
Public Const IDTB_CONTRACT = 201
Public Const IDTB_STOP = 202
Public Const IDTB_REFRESH = 203
Public Const IDTB_BACK = 204
Public Const IDTB_HOME = 205
Public Const IDTB_SYNC = 206
Public Const IDTB_PRINT = 207
Public Const IDTB_OPTIONS = 208
Public Const IDTB_FORWARD = 209
Public Const IDTB_NOTES = 210             ' not implemented
Public Const IDTB_BROWSE_FWD = 211
Public Const IDTB_BROWSE_BACK = 212
Public Const IDTB_CONTENTS = 213          ' not implemented
Public Const IDTB_INDEX = 214             ' not implemented
Public Const IDTB_SEARCH = 215            ' not implemented
Public Const IDTB_HISTORY = 216           ' not implemented
Public Const IDTB_BOOKMARKS = 217         ' not implemented
Public Const IDTB_JUMP1 = 218
Public Const IDTB_JUMP2 = 219
Public Const IDTB_CUSTOMIZE = 221
Public Const IDTB_ZOOM = 222
Public Const IDTB_TOC_NEXT = 223
Public Const IDTB_TOC_PREV = 224

Public Enum HHACT_
  HHACT_TAB_CONTENTS
  HHACT_TAB_INDEX
  HHACT_TAB_SEARCH
  HHACT_TAB_HISTORY
  HHACT_TAB_FAVORITES
    
  HHACT_EXPAND
  HHACT_CONTRACT
  HHACT_BACK
  HHACT_FORWARD
  HHACT_STOP
  HHACT_REFRESH
  HHACT_HOME
  HHACT_SYNC
  HHACT_OPTIONS
  HHACT_PRINT
  HHACT_HIGHLIGHT
  HHACT_CUSTOMIZE
  HHACT_JUMP1
  HHACT_JUMP2
  HHACT_ZOOM
  HHACT_TOC_NEXT
  HHACT_TOC_PREV
  HHACT_NOTES

  HHACT_LAST_ENUM
End Enum

Public Enum HHWIN_NAVTYPE_
  HHWIN_NAVTYPE_TOC
  HHWIN_NAVTYPE_INDEX
  HHWIN_NAVTYPE_SEARCH
  HHWIN_NAVTYPE_HISTORY       ' not implemented
  HHWIN_NAVTYPE_FAVORITES     ' not implemented
End Enum

Enum HHWIN_NAVTAB_
  HHWIN_NAVTAB_TOP
  HHWIN_NAVTAB_LEFT
  HHWIN_NAVTAB_BOTTOM
End Enum

Public Const HH_MAX_TABS = 19               ' maximum number of tabs

Public Enum HH_TAB_
  HH_TAB_CONTENTS
  HH_TAB_INDEX
  HH_TAB_SEARCH
  HH_TAB_HISTORY
  HH_TAB_FAVORITES
End Enum

Public Type HH_WINTYPE
  cbStruct As Long            ' IN: size of this structure including all Information Types
  fUniCodeStrings As Long     ' IN/OUT: TRUE if all strings are in UNICODE
  pszType  As String          ' IN/OUT: Name of a type of window
  fsValidMembers As Long      ' IN: Bit flag of valid members (HHWIN_PARAM_)
  fsWinProperties As Long     ' IN/OUT: Properties/attributes of the window (HHWIN_)
  pszCaption As String        ' IN/OUT: Window title
  dwStyles  As Long           ' IN/OUT: Window styles
  dwExStyles As Long          ' IN/OUT: Extended Window styles
  rcWindowPos As RECT         ' IN: Starting position, OUT: current position
  nShowState As Long          ' IN: show state (e.g., SW_SHOW)
  hwndHelp As Long            ' OUT: window handle
  hwndCaller As Long          ' OUT: who called this window
  paInfoTypes As Long         ' IN: Pointer to an array of Information Types

  ' The following members are only valid if HHWIN_PROP_TRI_PANE is set

  hwndToolBar As Long         ' OUT: toolbar window in tri-pane window
  hwndNavigation As Long      ' OUT: navigation window in tri-pane window
  hwndHTML As Long            ' OUT: window displaying HTML in tri-pane window
  iNavWidth As Long           ' IN/OUT: width of navigation window
  rcHTML As RECT              ' OUT: HTML window coordinates

  pszToc As String            ' IN: Location of the table of contents file
  pszIndex As String          ' IN: Location of the index file
  pszFile As String           ' IN: Default location of the html file
  pszHome As String           ' IN/OUT: html file to display when Home button is clicked
  fsToolBarFlags As Long      ' IN: flags controling the appearance of the toolbar
  fNotExpanded As Long        ' IN: TRUE/FALSE to contract or expand, OUT: current state
  curNavType As Long          ' IN/OUT: UI to display in the navigational pane
  tabpos As HHWIN_NAVTAB_     ' IN/OUT: HHWIN_NAVTAB_TOP, HHWIN_NAVTAB_LEFT, or HHWIN_NAVTAB_BOTTOM
  idNotify As Long            ' IN: ID to use for WM_NOTIFY messages
  tabOrder(HH_MAX_TABS) As Byte ' IN/OUT: tab order: Contents, Index, Search, History, Favorites, Reserved 1-5, Custom tabs
  cHistory As Long            ' IN/OUT: number of history items to keep (default is 30)
  pszJump1 As String          ' Text for HHWIN_BUTTON_JUMP1
  pszJump2 As String          ' Text for HHWIN_BUTTON_JUMP2
  pszUrlJump1 As String       ' URL for HHWIN_BUTTON_JUMP1
  pszUrlJump2 As String       ' URL for HHWIN_BUTTON_JUMP2
  rcMinSize As RECT           ' Minimum size for window (ignored in version 1)
  cbInfoTypes As Long         ' size of paInfoTypes;
End Type

' UDT for text popups
Public Type HH_POPUP
  cbStruct As Long                         ' sizeof this structure
  hinst As Long                               ' instance handle for string resource
  idString As Long                            ' string resource id, or text id if pszFile
                                              ' is specified in HtmlHelp call
  pszText As String                           ' used if idString is zero
  pt As POINTAPI                              ' top center of popup window
  clrForeground As ColorConstants             ' either use VB constant or &HBBGGRR
  clrBackground As ColorConstants             ' either use VB constant or &HBBGGRR
  rcMargins As RECT                           ' amount of space between edges of window and
                                              ' text, -1 for each member to ignore
  pszFont As String                           ' facename, point size, char set, BOLD ITALIC
                                              ' UNDERLINE
End Type

' UDT for keyword and ALink searches
Public Type HH_AKLINK
  cbStruct          As Long
  fReserved         As Boolean
  pszKeywords       As String
  pszUrl            As String
  pszMsgText        As String
  pszMsgTitle       As String
  pszWindow         As String
  fIndexOnFail      As Boolean
End Type

' UDT for accessing the Search tab
Public Type HH_FTS_QUERY
  cbStruct          As Long
  fUniCodeStrings   As Long
  pszSearchQuery    As String
  iProximity        As Long
  fStemmedSearch    As Long
  fTitleOnly        As Long
  fExecute          As Long
  pszWindow         As String
End Type

Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As Any) As Long



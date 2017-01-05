Attribute VB_Name = "modWinHelp"
'**********************************************************************
'Run the Help File
Declare Function WinHelp Lib "User32" Alias "WinHelpA" (ByVal hwndApp As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As Any) As Long

'WinHelp structures
Public Const HELPINFO_WINDOW = &H1&
Public Const HELPINFO_MENUITEM = &H2&
Type POINTAPI
    X As Long
    Y As Long
End Type
Type HELPINFO
    cbSize As Long
    iContextType As Long
    iCtrlId As Long
    hItemHandle As Long
    dwContextId As Long
    MousePos As POINTAPI
End Type
Type HELPWININFO
    wStructSize As Long
    X As Long
    Y As Long
    dx As Long
    dy As Long
    wMax As Long
    rgchMember As String * 2
End Type

'WinHelp API constants
Public Const HELP_COMMAND = &H102&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1
Public Const HELP_CONTEXTMENU = &HA&
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_CONTEXTNOFOCUS = &H108&
Public Const HELP_POPUPID = &H104&
Public Const HELP_FINDER = &HB&
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_HELPONHELP = &H4
Public Const HELP_INDEX = &H3
Public Const HELP_KEY = &H101
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_CLOSEWINDOW = &H107&
Public Const HELP_QUIT = &H2
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_SETINDEX = &H5
Public Const HELP_SETWINPOS = &H203&
Public Const HELP_SETPOPUP_POS = &HD&
Public Const HELP_TCARD = &H8000&
Public Const HELP_TCARD_DATA = &H10&
Public Const HELP_TCARD_OTHER_CALLER = &H11&
Public Const HELP_WM_HELP = &HC&

Public Const HELPMSGSTRING = "commdlg_help"


Attribute VB_Name = "modApi"
Option Explicit
Public Enum AXGMD
    AfterNewLine = 1001
    AfterCancel = 1002
    AfterDoubleClick = 1003
    AfterSave = 1004
    AfterReciept = 1005
End Enum
Public Enum TriadSetting
    DefaultState
    ViewState
    AddState
    EditState
    CancelState
    CloseState
    SaveState
    PrintState
End Enum
Public Enum MoveWhere
    rec_next
    rec_prev
    rec_first
    rec_last
    rec_all
End Enum
Public Enum FormName
    Quotation
    Group
    ProspectsPersonal
    ProspectCompany
    Customer
    Company
    TestDrive
    LogCall
    LogJournal
    Master
End Enum
'''Tool Bars background from CreatePatternBitMap
Private Const GCL_HBRBACKGROUND = (-10)                       ' FOR TB LONG
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'''COMBO BOX MAXLENGTH and Width
Public Const CB_LIMITTEXT = &H141                             'Combo Length
Public Const CB_SETDROPPEDWIDTH = &H160
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'''END COMBO BOX MAXLENGTH
Public Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As Long
'''TEXT OUT
Public Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long

Public Enum enuTBType
    enuTB_STANDARD = 0
    enuTB_FLAT = 1
End Enum

Public Sub ChangeTBBack(oBar As Object, picSource As PictureBox, pType As enuTBType)
    Dim lTBWnd                               As Long
    Dim LngNew                               As Long
    LngNew = CreatePatternBrush(picSource.Picture.Handle)
    Select Case pType
        Case enuTB_FLAT                                       'FLAT Button Style Toolbar
            DeleteObject SetClassLong(oBar.hwnd, GCL_HBRBACKGROUND, LngNew)    'Its Flat, Apply directly to TB Hwnd
        Case enuTB_STANDARD                                   'STANDARD Button Style Toolbar
            lTBWnd = FindWindowEx(oBar.hwnd, 0, "msvb_lib_toolbar", vbNullString)    'Standard, find Hwnd first
            DeleteObject SetClassLong(lTBWnd, GCL_HBRBACKGROUND, LngNew)    'Set new Back
    End Select
    InvalidateRect 0&, 0&, False

End Sub
Public Sub SetComboMaxLength(ComboBox As ComboBox, ByVal lMaxLength As Long)
    SendMessage ComboBox.hwnd, CB_LIMITTEXT, lMaxLength, ByVal 0&
End Sub




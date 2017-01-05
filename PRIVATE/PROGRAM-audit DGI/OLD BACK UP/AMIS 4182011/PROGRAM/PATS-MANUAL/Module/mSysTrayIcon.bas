Attribute VB_Name = "mSysTrayIcon"
Option Explicit

'modified by:
'nenes@naga.gov.ph
'October 1999


Public Const MAX_TOOLTIP As Integer = 64
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public loadchecker As Boolean

Public Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type

Public nfIconData As NOTIFYICONDATA

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Sub AddIcon2Systray(thisForm As Form)
' Add this application's icon to the system tray.
'
' Parm 1 = Handle of the window to receive notification messages
'          associated with an icon in the taskbar status area.
' Parm 2 = Icon to display.
' Parm 3 = Handle of icon to display.
' Parm 4 = Tooltip displayed when cursor moves over system tray icon.
'
With nfIconData
    .hwnd = thisForm.hwnd
    .uID = thisForm.Icon
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = thisForm.Icon.Handle
    .szTip = thisForm.Caption & vbNullChar
    .cbSize = Len(nfIconData)
End With
loadchecker = True
Call Shell_NotifyIcon(NIM_ADD, nfIconData)

End Sub

Sub RemoveIconFromSystray()
'
' Remove this application from the System Tray.
'
loadchecker = True
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)

End Sub

Sub SystrayIconWasClicked(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single
'
' Determine the event that happened to the System Tray icon.
' Left/Right clicking the icon displays a message box.
' Should be called by MouseMove Method
' it will still work even if the form is hidden or subclassed
'

lMsg = X / Screen.TwipsPerPixelX
Select Case lMsg
    Case WM_LBUTTONUP
       
    Case WM_RBUTTONUP
       
    Case WM_MOUSEMOVE
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONDBLCLK
       frmLOGIN.Visible = Not frmLOGIN.Visible
    Case WM_RBUTTONDOWN
       
    Case WM_RBUTTONDBLCLK
       'just to provide some way to terminate app
       'u can remove this if u like
       Unload frmLOGIN
    Case Else
      
End Select

End Sub


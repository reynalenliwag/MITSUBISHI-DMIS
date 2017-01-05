Attribute VB_Name = "modWin32Api"
Option Explicit


'
'show/hide cursor
'

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


'
'hide system tray
'

Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40

Private Declare Function FindWindow Lib _
   "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long

Private Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, _
   ByVal hWndInsertAfter As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal cx As Long, ByVal cy As Long, _
   ByVal wFlags As Long) As Long


'
'hide app from process list
'

Private Const RSP_SIMPLE_SERVICE = 1
Private Const RSP_UNREGISTER_SERVICE = 0

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" _
                 (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long


'
' enable/disable ctrl-alt-del
'
Private Const SPI_SCREENSAVERRUNNING = 97&

Private Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" _
   (ByVal uAction As Long, _
    ByVal uParam As Long, _
    lpvParam As Any, _
    ByVal fuWinIni As Long) As Long

Public Sub CtrlAltDel(lbValue As Boolean)
   ' true to disable, false to enable
   Dim lngRetVal As Long
   Dim blnPrevValue As Boolean
   lngRetVal = SystemParametersInfo(SPI_SCREENSAVERRUNNING, Not lbValue, _
               blnPrevValue, 0&)
End Sub



Public Sub ViewTaskBar(lbValue As Boolean)
Dim llResult As Long
llResult = FindWindow("Shell_traywnd", "")
If llResult Then
  If lbValue Then
     llResult = SetWindowPos(llResult, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
  Else
     llResult = SetWindowPos(llResult, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
  End If
End If
End Sub



Public Sub HideApp(lbValue As Boolean)
    Dim PID As Long
    Dim lngReturn As Long
    
    PID = GetCurrentProcessId()
    
    If lbValue Then
        'lngReturn = RegisterServiceProcess(PID, RSP_SIMPLE_SERVICE)
    Else
        lngReturn = RegisterServiceProcess(PID, RSP_UNREGISTER_SERVICE)
    End If
End Sub






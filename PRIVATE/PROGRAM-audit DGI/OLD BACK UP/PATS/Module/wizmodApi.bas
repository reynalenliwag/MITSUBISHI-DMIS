Attribute VB_Name = "wizmodAPI"
Option Explicit
Public Const MAX_COMPUTERNAME_LENGTH As Long = 15&
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public oVoice As SpeechLib.SpVoice

'Form Transparent
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

'Windows API/Global Declarations for
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000
Private m_bInIDE As Boolean
'**************************************
'Change Screen Resolution
'**************************************
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
dmDeviceName As String * CCDEVICENAME
dmSpecVersion As Integer
dmDriverVersion As Integer
dmSize As Integer
dmDriverExtra As Integer
dmFields As Long
dmOrientation As Integer
dmPaperSize As Integer
dmPaperLength As Integer
dmPaperWidth As Integer
dmScale As Integer
dmCopies As Integer
dmDefaultSource As Integer
dmPrintQuality As Integer
dmColor As Integer
dmDuplex As Integer
dmYResolution As Integer
dmTTOption As Integer
dmCollate As Integer
dmFormName As String * CCFORMNAME
dmUnusedPadding As Integer
dmBitsPerPel As Integer
dmPelsWidth As Long
dmPelsHeight As Long
dmDisplayFlags As Long
dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE
    
Global ResolutionWidth As Single
Global ResolutionHeight As Single
Global ScreenResolution As String

Global CurrentWidth As Single
Global CurrentHeight As Single

'Combo AutoComplete/DropDown
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9

Public Function GetMachineName() As String
Dim plngSize As Long
Dim pstrBuffer As String
pstrBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
plngSize = Len(pstrBuffer)
If GetComputerName(pstrBuffer, plngSize) Then
   GetMachineName = Left$(pstrBuffer, plngSize)
End If
End Function

Public Sub MoveKeyPress(KeyCode As Integer)
Dim First3Letters As String
First3Letters = Mid(Screen.ActiveForm.ActiveControl.Name, 1, 3)
Select Case KeyCode
       Case 13
            If First3Letters = "cbo" Then
               If Screen.ActiveForm.ActiveControl.Text = "" Then
                  Call VBComBoBoxDroppedDown(Screen.ActiveForm.ActiveControl)
               Else
                  SendKeys "{TAB}^{HOME}+{END}"
               End If
            Else
               If First3Letters = "txt" Or First3Letters = "opt" Or First3Letters = "chk" Then
                  SendKeys "{TAB}^{HOME}+{END}"
               End If
            End If
       Case 40
            If First3Letters = "txt" Or First3Letters = "chk" Then
               SendKeys "{TAB}^{HOME}+{END}"
            End If
       Case 38
            If First3Letters = "txt" Or First3Letters = "chk" Then
               SendKeys "+{TAB}^{HOME}+{END}"
            End If
End Select
End Sub

Public Sub ShowADOErrors(gcon As ADODB.Connection)
Dim errLoop As ADODB.Error
Dim strHelp As String
For Each errLoop In gcon.Errors
    If errLoop.HelpFile = "" Then
       strHelp = " No Helpfile available"
    Else
       strHelp = " Helpfile: " & errLoop.HelpFile & "; HelpContext: " & errLoop.HelpContext
    End If
    MsgBoxXP "ADO Error #" & errLoop.Number & vbCrLf & "Source: " & errLoop.Source _
           & vbCrLf & "SQL State: " & errLoop.SQLState & "; Native Error: " & errLoop.NativeError _
           & vbCrLf & vbCrLf & "Description: " & errLoop.Description & vbCrLf & vbCrLf & strHelp, "ADO Error", XP_OKOnly, msg_Critical
Next
End Sub

Public Sub ShowVBError()
If CBool(Err) Then
   MsgBoxXP "VB Error #" & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & vbCrLf & "Description: " & Err.Description, "VB Runtime Error", XP_OKOnly, msg_Critical
   Err.Clear
End If
End Sub

Public Sub ShowNoRecord()
On Error Resume Next
oVoice.Speak "No Such Record!", SVSFlagsAsync
MsgBoxXP "No Such Record!", "No Record", XP_OKOnly, msg_Information
End Sub

Public Sub ShowCantFind(str2find As Variant)
On Error Resume Next
oVoice.Speak "Can't find " & str2find, SVSFlagsAsync
MsgBoxXP "Can't find " & str2find, "Not Found", XP_OKOnly, msg_Information
End Sub

Public Sub ShowDeletedMsg()
On Error Resume Next
oVoice.Speak "Record Successfully Deleted...", SVSFlagsAsync
MsgBoxXP "Record Successfully Deleted...", "Confirmed", XP_OKOnly, msg_Information
End Sub

Public Sub MsgSpeechBox(Msg As String)
On Error Resume Next
oVoice.Speak Msg, SVSFlagsAsync
End Sub

'Form Transparent
Public Function isTransparent(ByVal hwnd As Long) As Boolean
On Error Resume Next
Dim Msg As Long
Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
  isTransparent = True
Else
  isTransparent = False
End If
If Err Then
  isTransparent = False
End If
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If Err Then
  MakeTransparent = 2
End If
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
Dim Msg As Long
On Error Resume Next
Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
Msg = Msg And Not WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, Msg
SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If Err Then
  MakeOpaque = 2
End If
End Function

'**************************************
'Example: Call ChangeRes(800,600) to change to 800 x 600 resolution
'**************************************
    
Public Sub ChangeRes(ByVal iWidth As Single, ByVal iHeight As Single)
    Dim a As Boolean
    Dim i&
    i = 0
    Do
    a = EnumDisplaySettings(0&, i&, DevM)
    i = i + 1
    Loop Until (a = False)
    Dim b&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    b = ChangeDisplaySettings(DevM, 0)
    End Sub
    
Public Sub GetRes()

CurrentWidth = Screen.Width / 15
CurrentHeight = Screen.Height / 15

ResolutionWidth = Screen.Width / 15
ResolutionHeight = Screen.Height / 15
ScreenResolution = Str(ResolutionWidth) + ", " + Str(ResolutionHeight)
End Sub

Public Sub UnloadApp()
   SetErrorMode SEM_NOGPFAULTERRORBOX
   If ResolutionWidth <> CurrentWidth And ResolutionHeight <> CurrentHeight Then
      Call ChangeRes(CurrentWidth, CurrentHeight)
   End If
   End
End Sub

Public Sub UnloadForm(frm As Object)
   SetErrorMode SEM_NOGPFAULTERRORBOX
   Unload frm
End Sub

Public Property Get InIDE() As Boolean
   Debug.Assert (IsInIDE())
   InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
   m_bInIDE = True
   IsInIDE = m_bInIDE
End Function

'AutoMatchCBBox(cb1, KeyAscii)
Public Function AutoMatchCBBox(ByRef cbBox As ComboBox, ByVal KeyAscii As Integer) As Integer
    Dim strFindThis As String, bContinueSearch As Boolean
    Dim lResult As Long, lStart As Long, lLength As Long
    AutoMatchCBBox = 0
    bContinueSearch = True
    lStart = cbBox.SelStart
    lLength = cbBox.SelLength

    On Error GoTo ErrHandle
        
    If KeyAscii < 32 Then
        bContinueSearch = False
        cbBox.SelLength = 0
        If KeyAscii = Asc(vbBack) Then
            If lLength = 0 Then
                If Len(cbBox) > 0 Then
                    cbBox.Text = Left(cbBox.Text, Len(cbBox) - 1)
                End If
            Else
                cbBox.Text = Left(cbBox.Text, lStart)
            End If
            cbBox.SelStart = Len(cbBox)
        ElseIf KeyAscii = vbKeyReturn Then
            cbBox.SelStart = Len(cbBox)
            lResult = SendMessage(cbBox.hwnd, CBN_SELENDOK, 0, 0)
            AutoMatchCBBox = KeyAscii
        End If
    Else
        If lLength = 0 Then
            strFindThis = cbBox.Text & Chr(KeyAscii)
        Else
            strFindThis = Left(cbBox.Text, lStart) & Chr(KeyAscii)
        End If
    End If
    
    If bContinueSearch Then
        Call VBComBoBoxDroppedDown(cbBox)
        lResult = SendMessage(cbBox.hwnd, CB_SELECTSTRING, -1, ByVal strFindThis)
        If lResult = CB_ERR Then
            cbBox.Text = strFindThis
            cbBox.SelLength = 0
            cbBox.SelStart = Len(cbBox)
        Else
            cbBox.SelStart = Len(strFindThis)
            cbBox.SelLength = Len(cbBox) - cbBox.SelStart
        End If
    End If
    On Error GoTo 0
    Exit Function
    
ErrHandle:
    Debug.Print "Failed: AutoCompleteComboBox due to : " & Err.Description
    Debug.Assert False
    AutoMatchCBBox = KeyAscii
    On Error GoTo 0
End Function

Public Sub VBComBoBoxDroppedDown(ByRef cbBox As ComboBox)
    Call SendMessage(cbBox.hwnd, CB_SHOWDROPDOWN, Abs(True), 0)
End Sub

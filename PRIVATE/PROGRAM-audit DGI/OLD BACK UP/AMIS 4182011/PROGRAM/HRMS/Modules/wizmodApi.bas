Attribute VB_Name = "wizmodAPI"
Option Explicit
Public Const MAX_COMPUTERNAME_LENGTH As Long = 15&
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Global Const Pi = 3.14159265358979
Global Const strPi = "3.1415926535897932384626433832795"
Public LOGID As Long
Public LOGPASS As String
Public Const SYSTEM_OWNER_NAME1 = ""
Public Const SYSTEM_OWNER_NAME2 = ""
Public Const SYSTEM_OWNER_ADDRESS = ""
Public Const SYSTEM_OWNER_CONTACT = ""
Public Const SYSTEM_OWNER_TIN = ""
Public Const SYSTEM_POLE_TRANSITION = ""
Public Const SYSTEM_POLE_GRACE = ""
Public Const MAXIMUM_DIGIT = "###,###,##0.00"
Public Const DIGIT_FORMAT = "########0"
Public Const MOVEDOWN = "{TAB}^{HOME}+{END}"
Public Const MOVEUP = "+{TAB}^{HOME}+{END}"
Public POLEDISPLAY_EXIST As Boolean
Public Const ControlA = "^{A}"
Public Const ControlB = "^{B}"
Public Const ControlC = "^{C}"
Public Const ControlD = "^{D}"
Public Const ControlE = "^{E}"
Public Const ControlF = "^{F}"
Public Const ControlG = "^{G}"
Public Const ControlH = "^{H}"
Public Const ControlI = "^{I}"
Public Const ControlJ = "^{J}"
Public Const ControlK = "^{K}"
Public Const ControlL = "^{L}"
Public Const ControlM = "^{M}"
Public Const ControlN = "^{N}"
Public Const ControlO = "^{O}"
Public Const ControlP = "^{P}"
Public Const ControlQ = "^{Q}"
Public Const ControlR = "^R"
Public Const ControlS = "^S"
Public Const ControlT = "^T"
Public Const ControlU = "^U"
Public Const ControlV = "^V"
Public Const ControlW = "^W"
Public Const ControlX = "^X"
Public Const ControlY = "^Y"
Public Const ControlZ = "^Z"
Public Const POLE_LENGTH = 20
Public Const TOTAL_POLE_LENGTH = 40
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000
Private m_bInIDE As Boolean
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Private Type DevMode
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
Dim DevM As DevMode
    
Global ResolutionWidth As Single
Global ResolutionHeight As Single
Global ScreenResolution As String
Global CurrentWidth As Single
Global CurrentHeight As Single

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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
               If Screen.ActiveForm.ActiveControl.Text = "" Then Call VBComBoBoxDroppedDown(Screen.ActiveForm.ActiveControl) Else SendKeys MOVEDOWN
            Else
               If First3Letters = "txt" Or First3Letters = "opt" Or First3Letters = "chk" Then SendKeys MOVEDOWN
            End If
       Case 40
            If First3Letters = "txt" Or First3Letters = "chk" Then SendKeys MOVEDOWN
       Case 38
            If First3Letters = "txt" Or First3Letters = "chk" Then SendKeys MOVEUP
End Select
End Sub

Public Sub ShowADOErrors(gcon As ADODB.Connection)
Screen.MousePointer = 0
Dim errLoop As ADODB.Error
Dim strHelp As String
For Each errLoop In gcon.Errors
    If errLoop.HelpFile = "" Then strHelp = " No Helpfile available" Else strHelp = " Helpfile: " & errLoop.HelpFile & "; HelpContext: " & errLoop.HelpContext
    MsgBoxXP "ADO Error #" & errLoop.Number & vbCrLf & "Source: " & errLoop.Source _
             & vbCrLf & "SQL State: " & errLoop.SQLState & "; Native Error: " & errLoop.NativeError _
             & vbCrLf & vbCrLf & "Description: " & errLoop.Description & vbCrLf & vbCrLf & strHelp, "ADO Error", XP_OKOnly, msg_Critical
Next
End Sub

Public Sub ShowVBError()
Screen.MousePointer = 0
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
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak "Can't find " & str2find, SVSFlagsAsync
MsgBoxXP "Can't find " & str2find, "Not Found", XP_OKOnly, msg_Information
End Sub

Public Function ShowConfirmDelete() As Boolean
On Error Resume Next
oVoice.Speak "Delete selected record? are you sure?...", SVSFlagsAsync
If MsgBoxXP("Delete selected record? Are you sure...", "Confirm Delete", XP_YesNo, msg_Question) = True Then
   ShowConfirmDelete = True
Else
   ShowConfirmDelete = False
End If
End Function

Public Sub ShowDeletedMsg()
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak "Record Successfully Deleted...", SVSFlagsAsync
MsgBoxXP "Record Successfully Deleted...", "Confirmed", XP_OKOnly, msg_Information
End Sub

Public Sub ShowNothingToDeleteMsg()
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak "Nothing to Delete...", SVSFlagsAsync
MsgBoxXP "Nothing to Delete...", "Empty Record", XP_OKOnly, msg_Information
End Sub

Public Sub ShowFirstRecordMsg()
On Error Resume Next
oVoice.Speak "Beginning of Record...", SVSFlagsAsync
MsgBoxXP "Beginning of Record...", "First Record", XP_OKOnly, msg_Information
End Sub

Public Sub ShowLastRecordMsg()
On Error Resume Next
oVoice.Speak "End of Record...", SVSFlagsAsync
MsgBoxXP "End of Record...", "Last Record", XP_OKOnly, msg_Information
End Sub

Public Sub MsgSpeechBox(msg As String)
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak msg, SVSFlagsAsync
MsgBoxXP msg, "Info", XP_OKOnly, msg_Information
End Sub

Public Sub MsgSpeech(msg As String)
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak msg, SVSFlagsAsync
End Sub

Public Function MsgQuestionBox(msg As String, BoxTitle As String) As Boolean
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak msg, SVSFlagsAsync
MsgQuestionBox = MsgBoxXP(msg, BoxTitle, XP_YesNo, msg_Question)
End Function

Public Function InputSpeechBox(ByRef msg As String, Optional ByRef DefaultNumericValue As String) As Variant
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak msg, SVSFlagsAsync
InputSpeechBox = InputBoxXP(msg, "Find", DefaultNumericValue)
End Function

Public Function isTransparent(ByVal hwnd As Long) As Boolean
On Error Resume Next
Dim msg As Long
msg = GetWindowLong(hwnd, GWL_EXSTYLE)
If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then isTransparent = True Else isTransparent = False
If Err Then isTransparent = False
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  msg = msg Or WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, msg
  SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If Err Then MakeTransparent = 2
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
Dim msg As Long
On Error Resume Next
msg = GetWindowLong(hwnd, GWL_EXSTYLE)
msg = msg And Not WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, msg
SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If Err Then MakeOpaque = 2
End Function
  
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
Dim ShowCount As Integer
SetErrorMode SEM_NOGPFAULTERRORBOX
If ResolutionWidth <> CurrentWidth And ResolutionHeight <> CurrentHeight Then
   Call ChangeRes(CurrentWidth, CurrentHeight)
End If
'For ShowCount = 200 To 1 Step -50
'    MakeTransparent frm.hwnd, ShowCount
'    DoEvents
'Next
Unload frm
Set frm = Nothing
End Sub

Public Property Get InIDE() As Boolean
Debug.Assert (IsInIDE())
InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
m_bInIDE = True
IsInIDE = m_bInIDE
End Function

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

Public Function rsFindDuplicate(rs2Find As ADODB.Recordset, ByVal rsField2find, ByVal str2find) As Boolean
Screen.MousePointer = 0
On Error GoTo BFoundErr
Dim rsToFind As ADODB.Recordset
If Len(str2find) > 1 And Len(rsField2find) > 1 Then
   Set rsToFind = New ADODB.Recordset
   Set rsToFind = rs2Find.Clone
   rsToFind.Find rsField2find & " = '" & UCase(str2find) & "'"
   If Not rsToFind.EOF Then rsFindDuplicate = True Else rsFindDuplicate = False
End If
Exit Function
BFoundErr:
   MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
   rsFindDuplicate = False
End Function

Public Sub ShowAlreadyExistMsg(Ricord As Variant)
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak Ricord & " Already Exist!...", SVSFlagsAsync
MsgBoxXP Ricord & " Already Exist!...", "Duplicate Record Found", XP_OKOnly, msg_Exclamation
End Sub

Public Sub ShowIsRequiredMsg(Ricord As Variant)
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak Ricord & " is Required!...", SVSFlagsAsync
MsgBoxXP Ricord & " is Required!...", "Field must have a NumericValue", XP_OKOnly, msg_Exclamation
End Sub

Public Sub ShowSuccessFullyAdded()
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak "Data Successfully Added!...", SVSFlagsAsync
MsgBoxXP "Data Successfully Added!...", "wizweirdo's Message", XP_OKOnly, msg_Information
End Sub

Public Sub ShowSuccessFullyUpdated()
Screen.MousePointer = 0
On Error Resume Next
oVoice.Speak "Data Successfully Updated!...", SVSFlagsAsync
MsgBoxXP "Data Successfully Updated!...", "wizweirdo's Message", XP_OKOnly, msg_Information
End Sub

Public Function UpperAscii(Askey As Integer)
UpperAscii = Asc(UCase(Chr(Askey)))
End Function

Public Function OnlyNumeric(KeyCode As Integer) As Integer
If KeyCode <> vbKeyHome And KeyCode <> vbKeyEnd And KeyCode <> 8 And KeyCode <> 13 And KeyCode <> 27 And KeyCode <> 46 Then
   If KeyCode < 48 Or KeyCode > 57 Then
      If KeyCode <> 110 Or KeyCode <> 190 Then
         OnlyNumeric = 0
      Else
         OnlyNumeric = KeyCode
      End If
   Else
      OnlyNumeric = KeyCode
   End If
Else
   OnlyNumeric = KeyCode
End If
End Function

Public Function ToDoubleNumber(ByRef NumericText As Variant) As String
Dim Counter As Integer
Dim TempNumber As String
Dim FoundPeriod As Boolean
FoundPeriod = False: TempNumber = ""
For Counter = 1 To Len(NumericText)
    If Mid(NumericText, Counter, 1) = "." Then
       If FoundPeriod = False Then
          TempNumber = TempNumber & Mid(NumericText, Counter, 1)
          FoundPeriod = True
       End If
    Else
       TempNumber = TempNumber & Mid(NumericText, Counter, 1)
    End If
Next
If Trim(TempNumber) <> "" Then ToDoubleNumber = Format(Round(NumericVal(TempNumber), 2), MAXIMUM_DIGIT) Else ToDoubleNumber = "0.00"
End Function

Public Function NumericVal(NumericText As Variant) As Double
Dim Counter As Integer
Dim NumericValue As String
NumericValue = ""
If Trim(NumericText) <> "" Then
   If IsNumeric(NumericText) = True And Val(NumericText) > 0 Then
      For Counter = 1 To Len(NumericText)
          If Mid(NumericText, Counter, 1) <> "," Then
             NumericValue = NumericValue & Mid(NumericText, Counter, 1)
          End If
      Next
      NumericVal = Round(NumericValue, 2)
   Else
      NumericVal = 0
   End If
Else
   NumericVal = 0
End If
End Function

Public Function NumToSpeak(XXX As Double) As String
   If NumericVal(Right(ToDoubleNumber(XXX), 2)) > 0 Then
      NumToSpeak = "Please Pay " & Mid(NumToText(NumericVal(XXX)), 1, Len(NumToText(NumericVal(XXX))) - 11) & " pesos, and " & Right(ToDoubleNumber(XXX), 2) & " centavos"
   Else
      NumToSpeak = "Please Pay " & Mid(NumToText(NumericVal(XXX)), 1, Len(NumToText(NumericVal(XXX))) - 11) & " pesos only"
   End If
End Function

Function RoundUP(XXX As Variant)
Dim RoundedOff, Butal, StringNumber, Anne As String
StringNumber = Round(XXX, 2)
Dim Dianne, Raquel, Rommel As Integer
For Dianne = 1 To Len(StringNumber)
    Anne = Mid(StringNumber, Dianne, 1)
    If Anne = "." Then
       Butal = Mid(StringNumber, Dianne + 1, Len(StringNumber) - (Dianne))
       Exit For
    Else
       RoundedOff = RoundedOff & Anne
    End If
Next
If Len(Butal) > 1 Then
   If Right(Butal, 1) >= 1 And Right(Butal, 1) <= 5 Then Butal = Left(Butal, 1) & 5
   If Right(Butal, 1) > 5 Then Butal = (NumericVal(Left(Butal, 1)) + 1) & 0
   RoundUP = NumericVal(RoundedOff) + (NumericVal(Butal) / 100)
Else
   RoundUP = NumericVal(RoundedOff) + (NumericVal(Butal & 0) / 100)
End If
End Function

Public Sub Listview_Loadval(TisoyView As ListItems, RecSet As ADODB.Recordset)
Dim Indx As Integer
Dim i As Integer
TisoyView.Clear
If Not (RecSet.BOF And RecSet.EOF) Then
    While Not RecSet.EOF
        Indx = TisoyView.Count + 1
        TisoyView.Add Indx, , IIf(IsNull(RecSet(0)), "", Trim(RecSet(0)))
        For i = 1 To RecSet.Fields.Count - 1
            TisoyView(Indx).ListSubItems.Add i, , IIf(IsNull(RecSet(i)), "", Trim(RecSet(i)))
        Next i
        RecSet.MoveNext
    Wend
End If
Set RecSet = Nothing
End Sub

Function CommonLog(Number) As Double
    CommonLog = Log(Number) / Log(10)
End Function

Function Factorial(Number) As Double
    Dim dblLoop, dblFactorial As Double
    dblFactorial = 1
    
    For dblLoop = 1 To Number
        dblFactorial = dblFactorial * dblLoop
    Next dblLoop
    
    Factorial = dblFactorial
End Function

Function Sum(dBoundary, uBoundary) As Double
    Dim dblLoop, dblSum As Double
    dblSum = 0
    
    For dblLoop = dBoundary To uBoundary
    dblSum = dblSum + dblLoop
    Next dblLoop
    
    Sum = dblSum
End Function

Function Permutation(x, y) As Double
    Dim dblLoop1, dblLoop2, dblR1, dblR2, z As Double
    dblR1 = 1
    dblR2 = 1
    z = x - y
    
    For dblLoop1 = 1 To x
    dblR1 = dblR1 * dblLoop1
    Next dblLoop1
    
    For dblLoop2 = 1 To z
    dblR2 = dblR2 * dblLoop2
    Next dblLoop2
    
    Permutation = dblR1 / dblR2
End Function

Function Combination(n, r) As Double
    Dim dblLoop1, dblLoop2, dblLoop3 As Double
    Dim dblR1, dblR2, dblR3 As Double
    Dim z, dblCombination As Double
    
    dblR1 = 1
    dblR2 = 1
    dblR3 = 1
    z = n - r

    For dblLoop1 = 1 To n
    dblR1 = dblR1 * dblLoop1
    Next dblLoop1
    
    For dblLoop2 = 1 To r
    dblR2 = dblR2 * dblLoop2
    Next dblLoop2
        
    For dblLoop3 = 1 To z
    dblR3 = dblR3 * dblLoop3
    Next dblLoop3
    
    dblCombination = dblR1 / (dblR2 * dblR3)
    Combination = dblCombination
End Function

Function Cot(Number) As Double
    Dim dblCot As Double
    If Sin((Number * Pi) / 180) = 0 Then
    dblCot = 1 / Tan((Number * Pi) / 180)
    Else
    dblCot = Cos((Number * Pi) / 180) / Sin((Number * Pi) / 180)
    End If
    Cot = dblCot
End Function

Public Sub Listview_Loadval2(TisoyView As ListItems, RecSet As ADODB.Recordset)
Dim Indx As Long
Dim i As Long
TisoyView.Clear
If Not (RecSet.BOF And RecSet.EOF) Then
   While Not RecSet.EOF
         Indx = TisoyView.Count + 1
         TisoyView.Add Indx, , IIf(IsNull(RecSet(0)), "", Trim(RecSet(0)))
         For i = 1 To RecSet.Fields.Count - 1
            TisoyView(Indx).ListSubItems.Add i, , IIf(IsNull(RecSet(i)), "", Trim(RecSet(i)))
         Next i
         RecSet.MoveNext
    Wend
End If
Set RecSet = Nothing
End Sub

Public Sub Combo_Loadnumericval(WeirdoCombo As ComboBox, RecSet As ADODB.Recordset)
WeirdoCombo.Clear
If Not (RecSet.BOF And RecSet.EOF) Then
    While Not RecSet.EOF
          WeirdoCombo.AddItem Null2String(RecSet(0))
          RecSet.MoveNext
    Wend
End If
Set RecSet = Nothing
End Sub


Attribute VB_Name = "modRAMSMain"
Option Explicit
Public OVERWRAYT                        As Boolean
Public SMIS_REPORT_PATH, HRMS_REPORT_PATH, CRIS_REPORT_PATH, OSMS_REPORT_PATH As String
Attribute HRMS_REPORT_PATH.VB_VarUserMemId = 1073741825
Attribute CRIS_REPORT_PATH.VB_VarUserMemId = 1073741825
Attribute OSMS_REPORT_PATH.VB_VarUserMemId = 1073741825
Public PMIS_REPORT_PATH, CSMS_REPORT_PATH, CMIS_REPORT_PATH, AMIS_REPORT_PATH As String
Attribute PMIS_REPORT_PATH.VB_VarUserMemId = 1073741829
Attribute CSMS_REPORT_PATH.VB_VarUserMemId = 1073741829
Attribute CMIS_REPORT_PATH.VB_VarUserMemId = 1073741829
Attribute AMIS_REPORT_PATH.VB_VarUserMemId = 1073741829
Public SKIN_PATH                        As String
Attribute SKIN_PATH.VB_VarUserMemId = 1073741833
Public MODULENAME                       As String
Attribute MODULENAME.VB_VarUserMemId = 1073741834
Public wizVar, CryptVar                 As Object
Attribute wizVar.VB_VarUserMemId = 1073741835
Attribute CryptVar.VB_VarUserMemId = 1073741835
Public ServerName                       As String
Attribute ServerName.VB_VarUserMemId = 1073741837
Public SQLSERVERNAME                    As String
Attribute SQLSERVERNAME.VB_VarUserMemId = 1073741838
Public Database                         As String
Attribute Database.VB_VarUserMemId = 1073741839
Public Company_name                     As String
Attribute Company_name.VB_VarUserMemId = 1073741840
Public Company_Address                  As String
Attribute Company_Address.VB_VarUserMemId = 1073741841
Public LOGCODE                          As String
Attribute LOGCODE.VB_VarUserMemId = 1073741842
Public LOGNAME                          As String
Attribute LOGNAME.VB_VarUserMemId = 1073741843
Public LOGLEVEL                         As String
Attribute LOGLEVEL.VB_VarUserMemId = 1073741844
Public LOGTIME                          As String
Attribute LOGTIME.VB_VarUserMemId = 1073741845
Public LOGDATE                          As String
Attribute LOGDATE.VB_VarUserMemId = 1073741846
Public DMIS_Audit_Connection            As String
Attribute DMIS_Audit_Connection.VB_VarUserMemId = 1073741847

Public Const MAX_COMPUTERNAME_LENGTH    As Long = 15&
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public oVoice                           As SpeechLib.SpVoice
Attribute oVoice.VB_VarUserMemId = 1073741848
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public gconDMIS                         As ADODB.Connection
Attribute gconDMIS.VB_VarUserMemId = 1073741849
Public gconACCESS                       As ADODB.Connection
Attribute gconACCESS.VB_VarUserMemId = 1073741850
Public gconBIR_RELIEF                   As ADODB.Connection
Attribute gconBIR_RELIEF.VB_VarUserMemId = 1073741851

Public ConnStr                          As String
Attribute ConnStr.VB_VarUserMemId = 1073741852
Public DMIS_Connection                  As String
Attribute DMIS_Connection.VB_VarUserMemId = 1073741853
Public Const DMIS_REPORT_Connection = "DSN=DMIS;DSQ=DMIS"
Public Const DMISAUDIT_REPORT_Connection = "DSN=DMIS_AUDIT;DSQ=DMIS"
Public MODLEVEL                         As String
Attribute MODLEVEL.VB_VarUserMemId = 1073741854
Public LOGPASS                          As String
Attribute LOGPASS.VB_VarUserMemId = 1073741855
Public LOGID                            As Long
Attribute LOGID.VB_VarUserMemId = 1073741856
Public TIMER_REMIND                     As String
Attribute TIMER_REMIND.VB_VarUserMemId = 1073741857


Public Const MAXIMUM_DIGIT = "###,###,##0.00"
Public Const DIGIT_FORMAT = "###,###,##0"
Public Const ZERO = "0.00"
Public Const MOVEDOWN = "{TAB}^{HOME}+{END}"
Public Const MOVEUP = "+{TAB}^{HOME}+{END}"
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
Private m_bInIDE                        As Boolean
Attribute m_bInIDE.VB_VarUserMemId = 1073741859
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName                        As String * CCDEVICENAME
    dmSpecVersion                       As Integer
    dmDriverVersion                     As Integer
    dmSize                              As Integer
    dmDriverExtra                       As Integer
    dmFields                            As Long
    dmOrientation                       As Integer
    dmPaperSize                         As Integer
    dmPaperLength                       As Integer
    dmPaperWidth                        As Integer
    dmScale                             As Integer
    dmCopies                            As Integer
    dmDefaultSource                     As Integer
    dmPrintQuality                      As Integer
    dmColor                             As Integer
    dmDuplex                            As Integer
    dmYResolution                       As Integer
    dmTTOption                          As Integer
    dmCollate                           As Integer
    dmFormName                          As String * CCFORMNAME
    dmUnusedPadding                     As Integer
    dmBitsPerPel                        As Integer
    dmPelsWidth                         As Long
    dmPelsHeight                        As Long
    dmDisplayFlags                      As Long
    dmDisplayFrequency                  As Long
End Type
Dim DevM                                As DEVMODE
Attribute DevM.VB_VarUserMemId = 1073741860

Global ResolutionWidth As Single
Attribute ResolutionWidth.VB_VarUserMemId = 1073741861
Global ResolutionHeight As Single
Attribute ResolutionHeight.VB_VarUserMemId = 1073741862
Global ScreenResolution As String
Attribute ScreenResolution.VB_VarUserMemId = 1073741863
Global CurrentWidth As Single
Attribute CurrentWidth.VB_VarUserMemId = 1073741864
Global CurrentHeight As Single
Attribute CurrentHeight.VB_VarUserMemId = 1073741865

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public RightX, BottomY, ASpeed          As Integer
Attribute RightX.VB_VarUserMemId = 1073741866
Attribute BottomY.VB_VarUserMemId = 1073741866
Attribute ASpeed.VB_VarUserMemId = 1073741866
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName$, ByVal lpDeviceName$, ByVal lpOutput$, ByVal lpInitData&)
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare Function StretchBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal XSrc&, ByVal YSrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const NOTSRCCOPY = &H330008
Private Ind, Xo, Yo, Xs, Ys, XSrc       As Long
Attribute Ind.VB_VarUserMemId = 1073741869
Attribute Xo.VB_VarUserMemId = 1073741869
Attribute Yo.VB_VarUserMemId = 1073741869
Attribute Xs.VB_VarUserMemId = 1073741869
Attribute Ys.VB_VarUserMemId = 1073741869
Attribute XSrc.VB_VarUserMemId = 1073741869
Private YSrc, DDC, SDC, res             As Long
Attribute YSrc.VB_VarUserMemId = 1073741875
Attribute DDC.VB_VarUserMemId = 1073741875
Attribute SDC.VB_VarUserMemId = 1073741875
Attribute res.VB_VarUserMemId = 1073741875
Dim z2                                  As Long
Attribute z2.VB_VarUserMemId = 1073741879

Public Const TTLDYSIN1YR = 365

Public Sub Main()
    If App.PrevInstance = True Then
        MsgBox "There is open DSA application", vbInformation
        End
    End If
    
    Load frmMain
    On Error GoTo adder:
    If CheckServerSettings = False Then
        frmFiles_ServerSetting.intSteps = 0

        frmFiles_ServerSetting.ShowLogin = True
        frmFiles_ServerSetting.Show
        Exit Sub

        If CheckLoginSetting = False Then
            frmSecurity.Show
            Exit Sub
        End If
        If AMIS_REPORT_PATH = "" Or CMIS_REPORT_PATH = "" Or CRIS_REPORT_PATH = "" Or CSMS_REPORT_PATH = "" Or HRMS_REPORT_PATH = "" Or OSMS_REPORT_PATH = "" Or SMIS_REPORT_PATH = "" Then
            If MsgBox("Report Path has not Been Set Do you Want to Set it Now", vbOKCancel + vbInformation) = vbOK Then
                frmFiles_ServerSetting.ShowLogin = True
                frmFiles_ServerSetting.intSteps = 2
                frmFiles_ServerSetting.Show
                Exit Sub
            End If
        End If
    Else
        AMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "AMIS") & "\"
        CMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CMIS") & "\"
        CRIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CRIS") & "\"
        CSMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CSMS") & "\"
        HRMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "HRMS") & "\"
        OSMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "OSMS") & "\"
        SMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "SMIS") & "\"
        PMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "PMIS") & "\"
        frmSecurity.Show
    End If
    Exit Sub
adder:
    Err.Clear
End Sub
Sub Checkreportsettings()
    AMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "AMIS")
    CMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CMIS")
    CRIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CRIS")
    CSMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CSMS")
    HRMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "HRMS")
    OSMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "OSMS")
    SMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "SMIS")
    PMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "PMIS")
End Sub
Function CheckServerSettings() As Boolean
    ServerName = GetSetting("DMIS 2.0", "SETTINGS", "SERVERNAME")
    SQLSERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SQLSERVERNAME")
    Database = GetSetting("DMIS 2.0", "SETTINGS", "DATABASE")
    If ServerName = "" Then: CheckServerSettings = False: Exit Function
    If SQLSERVERNAME = "" Then: CheckServerSettings = False: Exit Function
    If Database = "" Then: CheckServerSettings = False: Exit Function
    DMIS_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Database & ";Data Source=" & ServerName
    DMIS_Audit_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS_AUDIT ;Data Source=" & ServerName
    CheckServerSettings = OpenConnection
End Function

Function CheckLoginSetting() As Boolean
    Dim TEMPRS                          As ADODB.Recordset
    Dim cnx                             As ADODB.Connection
    Set cnx = New ADODB.Connection
    On Error GoTo adder:
    cnx.Open ("Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Database & ";Data Source=" & SQLSERVERNAME)
    Set TEMPRS = cnx.Execute("Select * from ALL_Rams_Users WHERE USERGROUP='SDM'")
    If Not (TEMPRS.BOF Or TEMPRS.EOF) Then
        LOGNAME = Null2String(TEMPRS!Username)
        CheckLoginSetting = True
    End If
    Set cnx = Nothing
    Exit Function
adder:
    Err.Clear
    CheckLoginSetting = False
End Function

Sub FillCombo(nSQL As String, itemdatarow As Integer, ilng As Integer, cmb As ComboBox)

    Dim nrs                             As New ADODB.Recordset
    Set nrs = gconDMIS.Execute(nSQL)
    cmb.Clear
    While Not nrs.EOF
        cmb.AddItem nrs.Collect(ilng)
        If itemdatarow <> -1 Then
            cmb.ItemData(cmb.NewIndex) = nrs.Collect(itemdatarow)
        End If
        nrs.MoveNext
    Wend
    nrs.Close
    Set nrs = Nothing


End Sub
Public Sub ConfigHeaders(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                            As String
    Dim cWidth                          As Long
    Dim i                               As Integer
    Dim scwidth                         As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For i = LBound(ar) To UBound(ar)
            If i <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(i)) / 100)
                grd.ColumnHeaders(i + 1).Width = scwidth
            End If
        Next
    Else

        For i = LBound(ar) To UBound(ar)
            If i < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(i)) / 100)
                grd.Columns(i).Width = scwidth
            End If
        Next
    End If

    Erase ar
    grd.Visible = True
End Sub

Public Sub flex_FillListView(RS As Recordset, grd As ListView, Optional WithSN As Boolean = True, Optional WITHCOLUMNHEADER As Boolean)
    Dim fld                             As Field
    Dim j                               As Long
    Dim ijx                             As Integer
    Dim LST                             As ListItem
    Dim i                               As Integer


    grd.ListItems.Clear

    If WithSN = True And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        Call grd.ColumnHeaders.Add(, , "Item")
        For i = 0 To RS.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , RS.Fields(i).Name)
        Next
        While Not RS.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend

    ElseIf WithSN = True And WITHCOLUMNHEADER = False Then
        grd.ColumnHeaders.Clear
        Call grd.ColumnHeaders.Add(, , "Item")
        While Not RS.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend

    ElseIf WithSN = False And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        For i = 0 To RS.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , RS.Fields(i).Name)
        Next
        j = RS.Fields.Count
        While Not RS.EOF
            Set LST = grd.ListItems.Add(, , RS.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    Else
        j = RS.Fields.Count
        While Not RS.EOF
            Set LST = grd.ListItems.Add(, , RS.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    End If
    Set LST = Nothing
End Sub

Public Function flex_FillReportView(RS As Recordset, grd As ReportControl, Optional ByVal WithSN As Boolean = False)

    Dim fld                             As Field
    Dim j                               As Long
    Dim REC                             As XtremeReportControl.ReportRecord


    grd.Records.DeleteAll


    While Not RS.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Public Function DaysInMonth(pDate As String) As String
    Select Case pDate
        Case 1, 3, 5, 7, 8, 10, 12
            DaysInMonth = "31"
        Case 4, 6, 9, 11
            DaysInMonth = "30"
        Case 2
            If (Year(pDate) Mod 4) = 0 Then
                DaysInMonth = "29"
            Else
                DaysInMonth = "28"
            End If
    End Select
End Function
Public Sub AddColumnHeader(StringHeaders As String, lvGrid As ListView)
    Dim ar()                            As String
    Dim cWidth                          As Long
    Dim i                               As Integer

    ar = Split(StringHeaders, ",")
    cWidth = lvGrid.Width
    lvGrid.ColumnHeaders.Clear
    For i = LBound(ar) To UBound(ar)
        lvGrid.ColumnHeaders.Add , , ar(i)
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Function SelectCombo(C As ComboBox, str As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim i                               As Long
    Dim ItemDataX                       As Long
    If ByItemData = False Then
        For i = 0 To C.ListCount - 1
            If UCase(C.List(i)) = UCase(Trim(str)) Then
                SelectCombo = i
                Exit Function
            End If
        Next
    Else
        If str = vbNullString Then
            SelectCombo = -1
            Exit Function
        End If
        ItemDataX = CLng(str)
        For i = 0 To C.ListCount - 1
            If C.ItemData(i) = str Then
                SelectCombo = i
                Exit Function
            End If
        Next
    End If
    SelectCombo = -1
End Function

Sub ReportControlPaintManager(LST As ReportControl)
    With LST
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots   ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Sub ReportControlAddColumnHeader(LST As ReportControl, StringHeaders As String)
    Dim ar()                            As String
    Dim i                               As Integer

    ar = Split(StringHeaders, ",")
    LST.Columns.DeleteAll
    For i = LBound(ar) To UBound(ar)
        LST.Columns.Add i, ar(i), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Public Function GetMachineName() As String
    Dim plngSize                        As Long
    Dim pstrBuffer                      As String
    pstrBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
    plngSize = Len(pstrBuffer)
    If GetComputerName(pstrBuffer, plngSize) Then
        GetMachineName = Left$(pstrBuffer, plngSize)
    End If
End Function

Public Sub MoveKeyPress(KeyCode As Integer)

    Dim First3Letters                   As String
    If Screen.ActiveForm.ActiveControl Is Nothing Then
        Exit Sub
    End If
    First3Letters = Mid(Screen.ActiveForm.ActiveControl.Name, 1, 3)
    Select Case KeyCode
        Case 13
            If First3Letters = "cbo" Then
                If Screen.ActiveForm.ActiveControl.Text = "" Then Else SendKeys MOVEDOWN
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
    Dim errLoop                         As ADODB.Error
    Dim strHelp                         As String
    For Each errLoop In gcon.Errors
        If errLoop.HelpFile = "" Then strHelp = " No Helpfile available" Else strHelp = " Helpfile: " & errLoop.HelpFile & "; HelpContext: " & errLoop.HelpContext
        MsgBox "ADO Error '" & errLoop.Number & vbCrLf & "Source: " & errLoop.Source _
             & vbCrLf & "SQL State: " & errLoop.SQLState & "; Native Error: " & errLoop.NativeError _
             & vbCrLf & vbCrLf & "Description: " & errLoop.DESCRIPTION & vbCrLf & vbCrLf & strHelp, vbCritical, "ADO Error"
    Next
End Sub

Public Sub ShowVBError()
    Screen.MousePointer = 0
    If CBool(Err) Then
        MsgBox "VB Error '" & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & vbCrLf & "Description: " & Err.DESCRIPTION, vbCritical, "VB Runtime Error"
        Err.Clear
    End If
End Sub

Public Sub ShowNoRecord()
    On Error Resume Next
    oVoice.Speak "No Such Record!", SVSFlagsAsync
    MessagePop RecNotFound, "Empty", "No Such Record", 1000
End Sub

Public Sub ShowCantFind(str2find As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak "Can't find " & str2find, SVSFlagsAsync
    MessagePop RecNotFound, "Not Found", "Can't find" & str2find, 1000
End Sub

Public Function ShowConfirmDelete() As Boolean
    On Error Resume Next
    oVoice.Speak "Delete selected record? are you sure?...", SVSFlagsAsync
    If MsgBox("Delete selected record? Are you sure...", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
        ShowConfirmDelete = True
    Else
        ShowConfirmDelete = False
    End If
End Function

Public Sub ShowDeletedMsg()
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak "Record Successfully Deleted...", SVSFlagsAsync
    MessagePop Delete, "Confirmed", "Record Successfully Deleted..."
End Sub

Public Sub ShowNothingToDeleteMsg()
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak "Nothing to Delete...", SVSFlagsAsync
    MessagePop RecNotFound, "Empty Record", "Nothing to Delete..."
End Sub

Public Sub ShowFirstRecordMsg()
    On Error Resume Next
    oVoice.Speak "Beginning of Record...", SVSFlagsAsync
    MessagePop NaviBegin, "Beginning of Record", "First Record"
End Sub

Public Sub ShowLastRecordMsg()
    On Error Resume Next
    oVoice.Speak "End of Record...", SVSFlagsAsync
    MessagePop NaviEnd, "End of Record", "Last Record", 1500
End Sub

Public Sub MsgSpeechBox(Msg As String)
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak Msg, SVSFlagsAsync
    MsgBox Msg, vbInformation, "Info"
End Sub

Public Sub MsgSpeech(Msg As String)
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak Msg, SVSFlagsAsync
End Sub

Public Function MsgQuestionBox(Msg As String, BoxTitle As String) As Boolean
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak Msg, SVSFlagsAsync
    If MsgBox(Msg, vbQuestion + vbYesNo, BoxTitle) = vbYes Then
        MsgQuestionBox = True
    Else
        MsgQuestionBox = False
    End If
End Function

Public Sub ShowAlreadyExistMsg(Ricord As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop RecSaveError, "Duplicate Record", Ricord & " Already Exist!..."
    oVoice.Speak Ricord & " Already Exist!...", SVSFlagsAsync
End Sub

Public Sub ShowIsRequiredMsg(Ricord As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak Ricord & " is Required!...", SVSFlagsAsync
    MessagePop RecSaveError, "Missing Filelds", "Field must have a Value!..." & Ricord, 1500
End Sub

Public Sub ShowSuccessFullyAdded()
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak "Data Successfully Added!...", SVSFlagsAsync
    MessagePop RecSaveOk, "Record Added", "Data Successfully Added!..."
End Sub

Public Sub ShowSuccessFullyUpdated()
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak "Data Successfully Updated!...", SVSFlagsAsync
    MessagePop RecSaveOk, "Record Updated", "Data Successfully Updated!..."
End Sub

Public Function UpperAscii(Askey As Integer)
    UpperAscii = Asc(UCase(Chr(Askey)))
End Function

Public Function OnlyNumeric(KeyCode As Integer) As Integer
    If KeyCode <> vbKeyHome And KeyCode <> vbKeyEnd And KeyCode <> 8 And KeyCode <> 13 And KeyCode <> 27 And KeyCode <> 46 Then
        If (KeyCode < 48 Or KeyCode > 57) And KeyCode <> 110 Then
            OnlyNumeric = 0
        Else
            OnlyNumeric = KeyCode
        End If
    Else
        OnlyNumeric = KeyCode
    End If
End Function

Public Function ToDoubleNumber(ByRef NumericText As Variant) As String
    Dim Counter                         As Integer
    Dim TempNumber                      As String
    Dim FoundPeriod                     As Boolean
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
    ToDoubleNumber = Format(TempNumber, MAXIMUM_DIGIT)
End Function

Public Function NumericVal(NumericText As Variant) As Double
    Dim Counter                         As Integer
    Dim NumericValue                    As String
    NumericValue = ""
    If Trim(NumericText) <> "" Then
        If IsNumeric(NumericText) = True Then
            If Val(Abs(NumericText)) > 0 Then
                For Counter = 1 To Len(NumericText)
                    If Mid(NumericText, Counter, 1) <> "," Then
                        NumericValue = NumericValue & Mid(NumericText, Counter, 1)
                    End If
                Next
                NumericVal = NumericValue
            Else
                NumericVal = 0
            End If
        Else
            NumericVal = 0
        End If
    Else
        NumericVal = 0
    End If
End Function

Public Sub Listview_Loadval(TisoyView As ListItems, RecSet As ADODB.Recordset)
    Dim Indx                            As Long
    Dim i                               As Long
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

Public Sub Combo_Loadval(WeirdoCombo As ComboBox, RecSet As ADODB.Recordset)
    WeirdoCombo.Clear
    If Not (RecSet.BOF And RecSet.EOF) Then
        While Not RecSet.EOF
            WeirdoCombo.AddItem Null2String(RecSet(0))
            RecSet.MoveNext
        Wend
    End If
    Set RecSet = Nothing
End Sub

Public Function isTransparent(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    Dim Msg                             As Long
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then isTransparent = True Else isTransparent = False
    If Err Then isTransparent = False
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
    Dim Msg                             As Long
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
    If Err Then MakeTransparent = 2
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
    Dim Msg                             As Long
    On Error Resume Next
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
    MakeOpaque = 0
    If Err Then MakeOpaque = 2
End Function

Public Sub ChangeRes(ByVal iWidth As Single, ByVal iHeight As Single)
    Dim A                               As Boolean
    Dim i&
    i = 0
    Do
        A = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (A = False)
    Dim B&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    B = ChangeDisplaySettings(DevM, 0)
End Sub

Public Sub GetRes()
    CurrentWidth = Screen.Width / 15
    CurrentHeight = Screen.Height / 15
    ResolutionWidth = Screen.Width / 15
    ResolutionHeight = Screen.Height / 15
    ScreenResolution = str(ResolutionWidth) + ", " + str(ResolutionHeight)
End Sub

Public Sub UnloadApp()
    SetErrorMode SEM_NOGPFAULTERRORBOX
    If ResolutionWidth <> CurrentWidth And ResolutionHeight <> CurrentHeight Then
        Call ChangeRes(CurrentWidth, CurrentHeight)
    End If
    End
End Sub

Public Function OpenConnection() As Boolean
    Screen.MousePointer = 11
    On Error GoTo ConnErr
    Set gconDMIS = New ADODB.Connection
    gconDMIS.ConnectionString = DMIS_Connection
    gconDMIS.Mode = adModeReadWrite
    gconDMIS.CursorLocation = adUseClient
    gconDMIS.Open
    Dim rsProfile                       As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE ")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        Company_name = Null2String(rsProfile!CompanyName)
        Company_Address = Null2String(rsProfile!Companyaddress)
    Else


    End If
    OpenConnection = True
    Screen.MousePointer = 0
    Exit Function
ConnErr:
    Screen.MousePointer = 0
    MsgBox Err.DESCRIPTION
    OpenConnection = False

End Function

Sub CLEARSETTING()
    Call SaveSetting("DMIS 2.0", "SETTINGS", "SERVERNAME", "")
    Call SaveSetting("DMIS 2.0", "SETTINGS", "SQLSERVERNAME", "")
    Call SaveSetting("DMIS 2.0", "SETTINGS", "DATABASE", "")
    Call SaveSetting("DMIS 2.0", "REPORTS", "AMIS", "")
    Call SaveSetting("DMIS 2.0", "REPORTS", "CMIS", "")
    Call SaveSetting("DMIS 2.0", "REPORTS", "CRIS", "")
    Call SaveSetting("DMIS 2.0", "REPORTS", "CSMS", "")
    Call SaveSetting("DMIS 2.0", "REPORTS", "HRMS", "")
    Call SaveSetting("DMIS 2.0", "REPORTS", "OSMS", "")
    Call SaveSetting("DMIS 2.0", "REPORTS", "SMIS", "")
    Call SaveSetting("DMIS 2.0", "REPORTS", "PMIS", "")
End Sub

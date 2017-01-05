Attribute VB_Name = "modCRISMain"
Option Explicit

Public Sub Main()
    SERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SERVERNAME")
    SQLSERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SQLSERVERNAME")
    DATABASE = GetSetting("DMIS 2.0", "SETTINGS", "DATABASE")
    If SQLSERVERNAME = "" Or DATABASE = "" Then
        MsgBox "Application Not Yet Configured. Please Configure Server Setting From DSA.", vbCritical
        End
        Exit Sub
    End If

    ConnStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & DATABASE & " " & " ;Data Source=" & SQLSERVERNAME
    DMIS_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & DATABASE & " ;Data Source=" & SQLSERVERNAME
    DMIS_REPORT_Connection = "DSN=" & DATABASE & " ;DSQ=" & SQLSERVERNAME
    DMIS_Audit_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS_AUDIT ;Data Source=" & SQLSERVERNAME

    frmMain.Show
    frmMain.ZOrder 1
    frmSecurity.Show vbModal
    frmSecurity.ZOrder 1
    frmMain.Show
    frmMain.ZOrder 1
    frmMainMenu.Show
    ReminderModule ""



End Sub

Public Sub SetUserSettings()
    Call SetUserPathSettings
    With frmMain
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
        .StatusBar1.Panels(9).Text = "Server Name: " & SQLSERVERNAME
    End With
End Sub

Public Function OpenSQLDb() As Boolean
    Screen.MousePointer = 11
frmSecurity.Hide
    DoEvents
    ApplySecurityValidation = True

    On Error GoTo ConnErr
    Set gconDMIS = New ADODB.Connection
    gconDMIS.ConnectionString = DMIS_Connection

    DoEvents
    gconDMIS.Mode = adModeReadWrite
    gconDMIS.CursorLocation = adUseClient
    gconDMIS.Open
    OpenSQLDb = True
    SetCompanyProfile
    Screen.MousePointer = 0
    Exit Function

ConnErr:
    MsgBox Err.Description
    MsgBox "I can't open a connection!!! You may have to " & vbCrLf & _
           "LOG-IN again to connect to the server to run this program. " & vbCrLf & _
           "If you don't have an account contact your friendly " & vbCrLf & _
           "neighborhood SysAdministrator.", _
           vbOKOnly + vbCritical, "ERROR"
    End
End Function

Public Sub SetUserMenuSettings()
    If LOGLEVEL = "AUTHOR" Or LOGLEVEL = "ADMIN" Then
        'frmMain.mnuMaintenance.Enabled = True
    Else
        'frmMain.mnuMaintenance.Enabled = False
    End If
    With frmMain
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
    End With
End Sub




Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
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
    ElseIf TypeOf grd Is ReportControl Then
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

Public Sub flex_FillListView(RS As Recordset, grd As ListView, Optional WithSN As Boolean = False, Optional WITHCOLUMNHEADER As Boolean = False)
    Dim fld                             As Field
    Dim j                               As Long
    Dim ijx                             As Integer
    Dim lst                             As ListItem
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
            Set lst = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    lst.ListSubItems.Add , , vbNullString
                Else
                    lst.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend

    ElseIf WithSN = True And WITHCOLUMNHEADER = False Then

        While Not RS.EOF
            j = j + 1
            Set lst = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    lst.ListSubItems.Add , , vbNullString
                Else
                    lst.ListSubItems.Add , , fld.Value
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
            Set lst = grd.ListItems.Add(, , RS.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    lst.ListSubItems.Add , , vbNullString
                Else
                    lst.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    Else
        j = RS.Fields.Count
        While Not RS.EOF
            Set lst = grd.ListItems.Add(, , Null2String(RS.Fields(0).Value))
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    lst.ListSubItems.Add , , vbNullString
                Else
                    lst.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    End If
    Set lst = Nothing
    'Set rs = Nothing
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

Sub FillCombo(NSQL As String, ItemDataRow As Integer, ilng As Integer, cmb As ComboBox)
    Dim nrs                             As New ADODB.Recordset
    Set nrs = gconDMIS.Execute(NSQL)
    cmb.Clear
    While Not nrs.EOF
        If IsNull(nrs.Collect(ilng)) = False Then
            cmb.AddItem nrs.Collect(ilng)
            If ItemDataRow <> -1 Then
                cmb.ItemData(cmb.NewIndex) = nrs.Collect(ItemDataRow)
            End If
        End If
        nrs.MoveNext
    Wend
    nrs.Close
    Set nrs = Nothing


End Sub

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
Sub ColorIt(cntrl As Control, tmr As Timer)
    tmr.Enabled = True
    cntrl.BackColor = vbRed
    cntrl.ForeColor = vbYellow
End Sub
Function SelectCombo(C As ComboBox, STR As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim i                               As Long
    Dim ItemDataX                       As Long
    If ByItemData = False Then
        For i = 0 To C.ListCount - 1
            If UCase(C.List(i)) = UCase(Trim(STR)) Then
                SelectCombo = i
                Exit Function
            End If
        Next
    Else
        If STR = vbNullString Then
            SelectCombo = -1
            Exit Function
        End If

        ItemDataX = CLng(STR)

        For i = 0 To C.ListCount - 1
            If C.ItemData(i) = STR Then
                SelectCombo = i
                Exit Function
            End If
        Next
    End If
    SelectCombo = -1
End Function



Sub ReportControlPaintManager(lst As ReportControl)
    With lst
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

Sub ReportControlAddColumnHeader(lst As ReportControl, StringHeaders As String)
    Dim ar()                            As String
    Dim i                               As Integer


    ar = Split(StringHeaders, ",")
    lst.Columns.DeleteAll
    For i = LBound(ar) To UBound(ar)
        lst.Columns.Add i, ar(i), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString

End Sub



Function GetCustomerCode(lastname As String) As String
    Dim temprs                          As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                          As String
    lAlpha = Left(Trim(lastname), 1)
    Set temprs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        GetCustomerCode = Left(lastname, 1) & Format(Mid(temprs.Collect(0), 2, 5), "00000")
    Else
        GetCustomerCode = Left(lastname, 1) & "00001"
    End If
End Function


Sub ShowHidePictureBox2(cntl As Object, State As Boolean, Optional ByVal MasterObject As Object)
    cntl.Visible = State

    If Not (MasterObject Is Nothing) Then
        MasterObject.Enabled = Not State
    End If
    If State = True Then
        cntl.ZOrder 0
    Else
        cntl.ZOrder 1
    End If
End Sub


Sub ShadeControl(oBx As Object, ISTrue As Boolean, Optional ByVal xVal As Variant = vbNullString)
    If ISTrue Then
        oBx.Enabled = True
        oBx.BackColor = vbWhite
    Else
        oBx.Enabled = False
        oBx.BackColor = vbButtonFace
    End If
    If xVal <> vbNullString Then: oBx.Text = xVal
End Sub

Function GenerateCode(TABLENAME, FLDNAME As String, xFormat As String) As String
    Dim rsID                            As ADODB.Recordset

    Set rsID = gconDMIS.Execute("Select MAX( ISNULL(" & FLDNAME & ", 0) ) as IDFIELD from " & TABLENAME)
    If rsID.Fields(0).Value = 0 Then
        GenerateCode = Format(1, xFormat)
    Else
        GenerateCode = Format(Val(N2Str2Zero(rsID![IDFIELD])) + 1, xFormat)

    End If
    Set rsID = Nothing

End Function
'FUNCTION / FEATURE :To Check the item exists in list item or not
'DATE STARTED       :04262007
'LAST UPDATED       :04262007
'DATABASE UPDATES   :NONE
'WHO UPDATED        :AXP
'UPDATING CODE      :AXP0426200720:03

Function CheckListItem(lst As ListView, valueCode As String) As Integer
    'AXP0426200720:03
    Dim i                               As Integer
    CheckListItem = -1
    For i = 1 To lst.ListItems.Count
        If lst.ListItems(i).Text = valueCode Then
            CheckListItem = i
            Exit Function
        End If
    Next
End Function
Function FormExist(XXX As String)

    Dim FRM                             As Form

    For Each FRM In Forms
        If (UCase(FRM.Name) = UCase(XXX)) Then
            FormExist = True
        End If
    Next
    Set FRM = Nothing
End Function
Sub UPDATELOGTABLE(TABLENAME, ID)
    Dim SQL                             As String
    SQL = "UPDATE " & TABLENAME & " SET "
    SQL = SQL & " USERCODE =" & N2Str2Null(LOGCODE) & ", " & vbCrLf
    SQL = SQL & " LASTUPDATE =" & N2Str2Null(LOGDATE & " " & LOGTIME) & vbCrLf
    SQL = SQL & " WHERE ID=" & ID
    gconDMIS.Execute SQL
End Sub





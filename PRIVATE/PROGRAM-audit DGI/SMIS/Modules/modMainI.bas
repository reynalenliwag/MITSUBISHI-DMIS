Attribute VB_Name = "modMain"
Option Explicit
Public FILE_GRAPH As String
'Sub Main()
'    YearStart = "1/1/2006"
'    Set gconDMIS = New ADODB.Connection
'    WorkTimeStart = #8:00:00 AM#
'    WorkTimeEnd = #5:30:00 PM#
'    SystemInterval = 30
'    Connect
'  Call MasRecordSet.Open("Select * from CRIS_vw_Master_PullDown", gconDMIS, adOpenForwardOnly, adLockReadOnly)
'    '''''''''DELETE THIS PROC IF YOU HAVE THE FUNCTIONS IN DB
''Call TempCheck
'
'    Load frmMain
'        frmMain.Show
'        MainForm.Show
'
'
'
'End Sub
Sub CacheScripts()
'Dim strDir As String
'strDir = Environ("TEMP")
'FILE_GRAPH = strDir & "\graph.html"

'Dim fso As FileSystemObject
'Set fso = New FileSystemObject
'If fso.FileExists(FILE_GRAPH) Then: Exit Sub
 '   Call fso.CopyFolder(App.Path & "\graphs", strDir, True)
 '   Set fso = Nothing
End Sub

Sub LoadMaster(ByVal strTag As String)
    ' On Error GoTo adder:
    Dim frm                                  As Form
    For Each frm In Forms
        If frm.Tag = strTag Then
            frm.SetFocus
            Set frm = Nothing
            Exit Sub
        End If
    Next frm

    Set frm = New frmCRIS_EntryMaster
    frm.MasterType = strTag
'    frm.ShortcutCaption.caption = strTag & " List"
    frm.Tag = strTag
    frm.Show
    Set frm = Nothing
    Exit Sub
    'adder:
    '  If Err.Number = 35601 Then
    '    Err.Clear
    '      Exit Sub
    ' End If
End Sub
'''''''''DELETE THIS PROC
Sub TempCheck()
Dim SQL                                  As String
If gconDMIS.Execute("select COUNT(*) from dbo.sysobjects where id = object_id(N'[dbo].[NameOfMonth]') and xtype in (N'FN', N'IF', N'TF')").Fields(0).Value > 0 Then: Exit Sub

    SQL = SQL & "CREATE  FUNCTION NameOfMonth  (@DATE datetime) " & vbCrLf & _
          "RETURNS Varchar(30) " & vbCrLf & _
          "AS " & vbCrLf & _
        " BEGIN " & vbCrLf
    SQL = SQL & "   DECLARE @MonthName varchar(50) " & vbCrLf & _
        "   Declare @MonthNum int " & vbCrLf & _
        "   SET @monthNum=Month(@Date) " & vbCrLf & _
        "   IF (@MonthNum=1) " & vbCrLf & _
        "   SET @MonthName='JANUARY' " & vbCrLf & _
        "   IF (@MonthNum=2) " & vbCrLf & _
        "   SET @MonthName='FEBRUARY' " & vbCrLf & _
        "   IF (@MonthNum=3) " & vbCrLf & _
        "   SET @MonthName='MARCH' " & vbCrLf & _
        "   IF (@MonthNum=4) " & vbCrLf & _
        "   SET @MonthName='APRIL' " & vbCrLf & _
        "   IF (@MonthNum=5) " & vbCrLf & _
        "   SET @MonthName='MAY' " & vbCrLf & _
        "   IF (@MonthNum=6) " & vbCrLf

    SQL = SQL & "   SET @MonthName='JUNE' " & vbCrLf & _
        "   IF (@MonthNum=7) " & vbCrLf & _
        "   SET @MonthName='JULY' " & vbCrLf & _
        "   IF (@MonthNum=8) " & vbCrLf & _
        "   SET @MonthName='AUGUST' " & vbCrLf & _
        "   IF (@MonthNum=9) " & vbCrLf & _
        "   SET @MonthName='SEPTEMBER' " & vbCrLf & _
        "   IF (@MonthNum=10) " & vbCrLf & _
        "   SET @MonthName='OCTOBER' " & vbCrLf & _
        "   IF (@MonthNum=11) " & vbCrLf & _
        "   SET @MonthName='NOVEMBER' " & vbCrLf & _
        "   IF (@MonthNum=12) " & vbCrLf & _
        "   SET @MonthName='DECEMBER' " & vbCrLf & _
        "   RETURN(@MonthName)  " & vbCrLf & _
        "    " & vbCrLf & _
          "END "
          
          
On Error GoTo adder:
gconDMIS.Execute SQL
Exit Sub
adder:
Err.Clear

End Sub

Public Sub searchListView(ByRef sListView As ListView, ByVal sFindText As String)
    Dim tmp_listtview                        As ListItem
    Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem)
    If Not tmp_listtview Is Nothing Then
        tmp_listtview.EnsureVisible
        tmp_listtview.Selected = True
    End If
End Sub




Public Sub ComboList(objx As Object, oRS As Recordset)

    objx.Clear
    While Not oRS.EOF
        objx.AddItem (Null2String(oRS.Collect(1)))
        objx.ItemData(objx.NewIndex) = oRS.Collect(0)
        oRS.MoveNext
    Wend
    If objx.ListCount > 0 Then
        If TypeName(objx) = "ComboBox" Then
            objx.ListIndex = 0
        End If
    End If
End Sub

Public Sub SetMyCaps(frm As Form, strCapTag As String)
    frm.Tag = strCapTag

End Sub

Public Sub ConfigHeaders(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                 As String
    Dim cWidth                               As Long
    Dim I                                    As Integer
    Dim scwidth                              As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    Else

        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next
    End If

    Erase ar
    grd.Visible = True
End Sub

Public Sub flex_FillListView(rs As Recordset, grd As ListView, Optional WithSN As Boolean = True, Optional WITHCOLUMNHEADER As Boolean)
    Dim FLD                                  As Field
    Dim j                                    As Long
    Dim ijx                                  As Integer
    Dim LST                                  As ListItem
    Dim I                                    As Integer


    grd.ListItems.Clear

    If WithSN = True And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        Call grd.ColumnHeaders.Add(, , "Item")
        For I = 0 To rs.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , rs.Fields(I).Name)
        Next
        While Not rs.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each FLD In rs.Fields
                If IsNull(FLD.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , FLD.Value
                End If
            Next
            rs.MoveNext
        Wend

    ElseIf WithSN = True And WITHCOLUMNHEADER = False Then
        grd.ColumnHeaders.Clear
        Call grd.ColumnHeaders.Add(, , "Item")
        While Not rs.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each FLD In rs.Fields
                If IsNull(FLD.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , FLD.Value
                End If
            Next
            rs.MoveNext
        Wend

    ElseIf WithSN = False And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        For I = 0 To rs.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , rs.Fields(I).Name)
        Next
        j = rs.Fields.Count
        While Not rs.EOF
            Set LST = grd.ListItems.Add(, , rs.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(rs.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , rs.Fields(ijx).Value
                End If
            Next
            rs.MoveNext
        Wend
    Else
        j = rs.Fields.Count
        While Not rs.EOF
            Set LST = grd.ListItems.Add(, , rs.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(rs.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , rs.Fields(ijx).Value
                End If
            Next
            rs.MoveNext
        Wend
    End If
    Set LST = Nothing
    'Set rs = Nothing
End Sub

Public Function flex_FillReportView(rs As Recordset, grd As ReportControl, Optional ByVal WithSN As Boolean = False)

    Dim FLD                                  As Field
    Dim j                                    As Long
    Dim REC                                  As XtremeReportControl.ReportRecord


    grd.Records.DeleteAll


    While Not rs.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each FLD In rs.Fields
            REC.AddItem (Trim(FLD.Value))
        Next
        rs.MoveNext
    Wend
    grd.Populate
    Set FLD = Nothing
    Set REC = Nothing
    Set rs = Nothing
End Function

Sub FillCombo(nSQL As String, itemdatarow As Integer, ilng As Integer, cmb As ComboBox)
    Dim nrs                                  As New ADODB.Recordset
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



Sub FontMeIn(xBold As Boolean, xItalic As Boolean, xColor As OLE_COLOR, oObx As Object)
    oObx.ForeColor = xColor
    oObx.FontBold = xBold
    oObx.FontItalic = xItalic
End Sub
Sub PrintHeader(DataLine As String, picx As PictureBox)
    picx.Cls
    picx.CurrentY = 60
    picx.Print Space(2); DataLine
End Sub



Public Sub AddColumnHeader(StringHeaders As String, lvGrid As ListView)
    Dim ar()                                 As String
    Dim cWidth                               As Long
    Dim I                                    As Integer

    ar = Split(StringHeaders, ",")
    cWidth = lvGrid.Width
    lvGrid.ColumnHeaders.Clear
    For I = LBound(ar) To UBound(ar)
        lvGrid.ColumnHeaders.Add , , ar(I)
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub
Sub ColorIt(cntrl As Control, tmr As Timer)
    tmr.Enabled = True
    cntrl.BackColor = vbRed
    cntrl.ForeColor = vbYellow
End Sub
Function SelectCombo(C As ComboBox, str As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim I                                    As Long
    Dim ItemDataX                            As Long
    If ByItemData = False Then
        For I = 0 To C.ListCount - 1
            If UCase(C.List(I)) = UCase(Trim(str)) Then
                SelectCombo = I
                Exit Function
            End If
        Next
    Else
        If str = vbNullString Then
            SelectCombo = -1
            Exit Function
        End If
        ItemDataX = CLng(str)
        For I = 0 To C.ListCount - 1
            If C.ItemData(I) = str Then
                SelectCombo = I
                Exit Function
            End If
        Next
    End If
    SelectCombo = -1
End Function



Sub ReportControlPaintManager(LST As ReportControl)
With LST
    .PaintManager.HorizontalGridStyle = xtpGridSmallDots  ' xtpGridNoLines
    .PaintManager.HighlightBackColor = RGB(34, 133, 13)
    .PaintManager.ShadeSortColor = RGB(250, 251, 189)
    .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
    .SetCustomDraw xtpCustomBeforeDrawRow
    .PaintManager.CaptionFont.Bold = True
    .PaintManager.GroupRowTextBold = True
    .PaintManager.GroupForeColor = vbBlue
    .PaintManager.ColumnStyle = xtpColumnExplorer
End With

End Sub

Sub ReportControlAddColumnHeader(LST As ReportControl, StringHeaders As String)
    Dim ar()                                 As String
    Dim I                                    As Integer

    ar = Split(StringHeaders, ",")
    LST.Columns.DeleteAll
    For I = LBound(ar) To UBound(ar)
        LST.Columns.Add I, ar(I), 100, True
    Next
    Erase ar
        StringHeaders = vbNullString
End Sub



Function GetCustomerCode(lastname As String) As String
Dim TempRs As ADODB.Recordset
If Len(lastname) = 0 Then
    Exit Function
End If
    Dim lAlpha As String
    lAlpha = Left(Trim(lastname), 1)
        Set TempRs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (TempRs.EOF Or TempRs.BOF) Then
        GetCustomerCode = Left(lastname, 1) & Format(Mid(TempRs.Collect(0), 2, 5), "00000")
    End If
End Function
Sub SetCustomerCode(lastname As String)
Dim SQLX As String
SQLX = "Update ALL_CUSCTL SET CTLCDE='" & lastname & "'" _
            & " Where LEFT(CTLCDE,1)='@AX'"
        SQLX = Replace(SQLX, "@AX", Left(lastname, 1))
        gconDMIS.Execute SQLX
End Sub




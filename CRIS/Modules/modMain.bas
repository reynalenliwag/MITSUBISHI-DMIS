Attribute VB_Name = "modMain"
Option Explicit

Public FILE_GRAPH As String
Sub Main()
   frmCRIS_EntryProfile.Show
    
        
    
    

End Sub
''Public gconDMIS                               As ADODB.Connection
'Public Const SQLConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS;Data Source=NS3"
'Public YearStart                             As Date
'Public mySkin                                As Integer
'Public WorkTimeStart                         As Date
'Public WorkTimeEnd                           As Date
'Public SystemInterval                        As Long
'Public MasRecordSet                          As New ADODB.Recordset
'Public ActiveFormEnum                        As FormName
'Sub CacheScripts()
'Dim strDir As String
'strDir = Environ("TEMP")
'FILE_GRAPH = strDir & "\graph.html"
'
'Dim fso As FileSystemObject
'Set fso = New FileSystemObject
'If fso.FileExists(FILE_GRAPH) Then: Exit Sub
'    Call fso.CopyFolder(App.Path & "\graphs", strDir, True)
'    Set fso = Nothing
'End Sub
'
'Sub LoadMaster(ByVal strTag As String)
'    ' On Error GoTo adder:
'    Dim frm                                  As Form
'    For Each frm In Forms
'        If frm.Tag = strTag Then
'            frm.SetFocus
'            Set frm = Nothing
'            Exit Sub
'        End If
'    Next frm
'
'    Set frm = New frmCRIS_EntryMaster
'    frm.MasterType = strTag
''    frm.ShortcutCaption.caption = strTag & " List"
'    frm.Tag = strTag
'    frm.Show
'    Set frm = Nothing
'    Exit Sub
'    'adder:
'    '  If Err.Number = 35601 Then
'    '    Err.Clear
'    '      Exit Sub
'    ' End If
'End Sub
''''''''''DELETE THIS PROC
'Sub TempCheck()
'Dim SQL                                  As String
'If gconDMIS.Execute("select COUNT(*) from dbo.sysobjects where id = object_id(N'[dbo].[NameOfMonth]') and xtype in (N'FN', N'IF', N'TF')").Fields(0).Value > 0 Then: Exit Sub
'
'    SQL = SQL & "CREATE  FUNCTION NameOfMonth  (@DATE datetime) " & vbCrLf & _
'          "RETURNS Varchar(30) " & vbCrLf & _
'          "AS " & vbCrLf & _
'        " BEGIN " & vbCrLf
'    SQL = SQL & "   DECLARE @MonthName varchar(50) " & vbCrLf & _
'        "   Declare @MonthNum int " & vbCrLf & _
'        "   SET @monthNum=Month(@Date) " & vbCrLf & _
'        "   IF (@MonthNum=1) " & vbCrLf & _
'        "   SET @MonthName='JANUARY' " & vbCrLf & _
'        "   IF (@MonthNum=2) " & vbCrLf & _
'        "   SET @MonthName='FEBRUARY' " & vbCrLf & _
'        "   IF (@MonthNum=3) " & vbCrLf & _
'        "   SET @MonthName='MARCH' " & vbCrLf & _
'        "   IF (@MonthNum=4) " & vbCrLf & _
'        "   SET @MonthName='APRIL' " & vbCrLf & _
'        "   IF (@MonthNum=5) " & vbCrLf & _
'        "   SET @MonthName='MAY' " & vbCrLf & _
'        "   IF (@MonthNum=6) " & vbCrLf
'
'    SQL = SQL & "   SET @MonthName='JUNE' " & vbCrLf & _
'        "   IF (@MonthNum=7) " & vbCrLf & _
'        "   SET @MonthName='JULY' " & vbCrLf & _
'        "   IF (@MonthNum=8) " & vbCrLf & _
'        "   SET @MonthName='AUGUST' " & vbCrLf & _
'        "   IF (@MonthNum=9) " & vbCrLf & _
'        "   SET @MonthName='SEPTEMBER' " & vbCrLf & _
'        "   IF (@MonthNum=10) " & vbCrLf & _
'        "   SET @MonthName='OCTOBER' " & vbCrLf & _
'        "   IF (@MonthNum=11) " & vbCrLf & _
'        "   SET @MonthName='NOVEMBER' " & vbCrLf & _
'        "   IF (@MonthNum=12) " & vbCrLf & _
'        "   SET @MonthName='DECEMBER' " & vbCrLf & _
'        "   RETURN(@MonthName)  " & vbCrLf & _
'        "    " & vbCrLf & _
'          "END "
'
'
'On Error GoTo adder:
'gconDMIS.Execute SQL
'Exit Sub
'adder:
'Err.Clear
'
'End Sub
'
'Public Sub searchListView(ByRef sListView As ListView, ByVal sFindText As String)
'    Dim tmp_listtview                        As ListItem
'    Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem)
'    If Not tmp_listtview Is Nothing Then
'        tmp_listtview.EnsureVisible
'        tmp_listtview.Selected = True
'    End If
'End Sub
'
'
'Function SelectComboItemData(ItemDT As Long, C As ComboBox)
'    If ItemDT = 0 Or C.ListCount = 0 Then: Exit Function
'    Dim i                                    As Integer
'    For i = 0 To C.ListCount - 1
'        If C.ItemData(i) = ItemDT Then
'            SelectComboItemData = i
'            Exit Function
'        End If
'    Next
'    SelectComboItemData = 0
'End Function
'
'
'Public Sub ComboList(objx As Object, oRS As Recordset)
'
'    objx.Clear
'    While Not oRS.EOF
'        objx.AddItem (oRS.Collect(1))
'        objx.ItemData(objx.NewIndex) = oRS.Collect(0)
'        oRS.MoveNext
'    Wend
'    If objx.ListCount > 0 Then
'        If TypeName(objx) = "ComboBox" Then
'            objx.ListIndex = 0
'        End If
'    End If
'End Sub
'
'Public Sub SetMyCaps(frm As Form, strCapTag As String)
'    frm.Tag = strCapTag
'
'End Sub
'
'Public Sub ConfigHeaders(grd As Object, SizeArray As String)
'    grd.Visible = False
'
'    Dim ar()                                 As String
'    Dim cWidth                               As Long
'    Dim i                                    As Integer
'    Dim scwidth                              As Long
'    ar = Split(SizeArray, ",")
'    cWidth = grd.Width
'
'    If TypeOf grd Is ListView Then
'        For i = LBound(ar) To UBound(ar)
'            If i <= grd.ColumnHeaders.Count Then
'                scwidth = cWidth * (CDec(ar(i)) / 100)
'                grd.ColumnHeaders(i + 1).Width = scwidth
'            End If
'        Next
'    Else
'
'        For i = LBound(ar) To UBound(ar)
'            If i < grd.Columns.Count Then
'                scwidth = cWidth * (CDec(ar(i)) / 100)
'                grd.Columns(i).Width = scwidth
'            End If
'        Next
'    End If
'
'    Erase ar
'    grd.Visible = True
'End Sub
'
'Public Sub flex_FillListView(rs As Recordset, grd As ListView, Optional WithSN As Boolean = True, Optional WITHCOLUMNHEADER As Boolean)
'    Dim FLD                                  As Field
'    Dim j                                    As Long
'    Dim ijx                                  As Integer
'    Dim lst                                  As ListItem
'    Dim i                                    As Integer
'
'
'    grd.ListItems.Clear
'
'    If WithSN = True And WITHCOLUMNHEADER = True Then
'        grd.ColumnHeaders.Clear
'        Call grd.ColumnHeaders.Add(, , "Item")
'        For i = 0 To rs.Fields.Count - 1
'            Call grd.ColumnHeaders.Add(, , rs.Fields(i).Name)
'        Next
'        While Not rs.EOF
'            j = j + 1
'            Set lst = grd.ListItems.Add(, , j)
'            For Each FLD In rs.Fields
'                If IsNull(FLD.Value) Then
'                    lst.ListSubItems.Add , , vbNullString
'                Else
'                    lst.ListSubItems.Add , , FLD.Value
'                End If
'            Next
'            rs.MoveNext
'        Wend
'
'    ElseIf WithSN = True And WITHCOLUMNHEADER = False Then
'        grd.ColumnHeaders.Clear
'        Call grd.ColumnHeaders.Add(, , "Item")
'        While Not rs.EOF
'            j = j + 1
'            Set lst = grd.ListItems.Add(, , j)
'            For Each FLD In rs.Fields
'                If IsNull(FLD.Value) Then
'                    lst.ListSubItems.Add , , vbNullString
'                Else
'                    lst.ListSubItems.Add , , FLD.Value
'                End If
'            Next
'            rs.MoveNext
'        Wend
'
'    ElseIf WithSN = False And WITHCOLUMNHEADER = True Then
'        grd.ColumnHeaders.Clear
'        For i = 0 To rs.Fields.Count - 1
'            Call grd.ColumnHeaders.Add(, , rs.Fields(i).Name)
'        Next
'        j = rs.Fields.Count
'        While Not rs.EOF
'            Set lst = grd.ListItems.Add(, , rs.Fields(0).Value)
'            For ijx = 1 To j - 1
'                If IsNull(rs.Fields(ijx).Value) Then
'                    lst.ListSubItems.Add , , vbNullString
'                Else
'                    lst.ListSubItems.Add , , rs.Fields(ijx).Value
'                End If
'            Next
'            rs.MoveNext
'        Wend
'    Else
'        j = rs.Fields.Count
'        While Not rs.EOF
'            Set lst = grd.ListItems.Add(, , rs.Fields(0).Value)
'            For ijx = 1 To j - 1
'                If IsNull(rs.Fields(ijx).Value) Then
'                    lst.ListSubItems.Add , , vbNullString
'                Else
'                    lst.ListSubItems.Add , , rs.Fields(ijx).Value
'                End If
'            Next
'            rs.MoveNext
'        Wend
'    End If
'    Set lst = Nothing
'    'Set rs = Nothing
'End Sub
'
'Public Function flex_FillReportView(rs As Recordset, grd As ReportControl, Optional ByVal WithSN As Boolean = False)
'
'    Dim FLD                                  As Field
'    Dim j                                    As Long
'    Dim REC                                  As XtremeReportControl.ReportRecord
'
'
'    grd.Records.DeleteAll
'
'
'    While Not rs.EOF
'        j = j + 1
'
'        Set REC = grd.Records.Add
'        If WithSN = True Then
'            REC.AddItem j
'        End If
'        For Each FLD In rs.Fields
'            REC.AddItem (Trim(FLD.Value))
'            REC.PreviewText = "ashish pya"
'        Next
'        rs.MoveNext
'    Wend
'    grd.Populate
'    Set FLD = Nothing
'    Set REC = Nothing
'    Set rs = Nothing
'End Function
'
'Sub FillCombo(nSQL As String, itemdatarow As Integer, ilng As Integer, cmb As ComboBox)
'
'    Dim nrs                                  As New ADODB.Recordset
'    Connect
'    Set nrs = gconDMIS.Execute(nSQL)
'    'nrs.Open nSQL, , adOpenForwardOnly, adLockReadOnly
'    cmb.Clear
'    While Not nrs.EOF
'        cmb.AddItem nrs.Collect(ilng)
'        If itemdatarow <> -1 Then
'            cmb.ItemData(cmb.NewIndex) = nrs.Collect(itemdatarow)
'        End If
'        nrs.MoveNext
'    Wend
'    nrs.Close
'    Set nrs = Nothing
'
'
'End Sub
'
'Public Function DaysInMonth(pDate As String) As String
'    Select Case pDate
'        Case 1, 3, 5, 7, 8, 10, 12
'            DaysInMonth = "31"
'        Case 4, 6, 9, 11
'            DaysInMonth = "30"
'        Case 2
'            If (Year(pDate) Mod 4) = 0 Then
'                DaysInMonth = "29"
'            Else
'                DaysInMonth = "28"
'            End If
'    End Select
'End Function
'
'Sub FillSuppliersDetail(ODxTxt As TextBox, lngSuppID As Long)
'    ODxTxt.Text = vbNullString
'    ODxTxt.Locked = True
'    If lngSuppID <= 0 Then: Exit Sub
'    Dim oRsx                                 As ADODB.Recordset
'    Set oRsx = GetRS("Select ContactPerson,Address,Phone,CellPhone,OpeningBal from Supplier Where SuppliersID=" & lngSuppID)
'    With ODxTxt
'        .Text = .Text & GetString(oRsx.Fields("ContactPerson")) & vbCrLf
'        .Text = .Text & "Tel: " & GetString(oRsx.Fields("Phone")) & Space(10) & " Mobile :" & GetString(oRsx.Fields("CellPhone")) & vbCrLf
'        .Text = .Text & "OpeningBal: " & GetDouble(oRsx.Fields("OpeningBal")) & " ( Balance: " & GetDouble(oRsx.Fields("OpeningBal")) & ")"
'    End With
'    Set oRsx = Nothing
'End Sub
'
'
'Sub FontMeIn(xBold As Boolean, xItalic As Boolean, xColor As OLE_COLOR, oObx As Object)
'    oObx.ForeColor = xColor
'    oObx.FontBold = xBold
'    oObx.FontItalic = xItalic
'End Sub
'Sub PrintHeader(DataLine As String, picx As PictureBox)
'    picx.Cls
'    picx.CurrentY = 60
'    picx.Print Space(2); DataLine
'End Sub
'
'
'
'Public Sub AddColumnHeader(StringHeaders As String, lvGrid As ListView)
'    Dim ar()                                 As String
'    Dim cWidth                               As Long
'    Dim i                                    As Integer
'
'    ar = Split(StringHeaders, ",")
'    cWidth = lvGrid.Width
'    lvGrid.ColumnHeaders.Clear
'    For i = LBound(ar) To UBound(ar)
'        lvGrid.ColumnHeaders.Add , , ar(i)
'    Next
'    Erase ar
'    StringHeaders = vbNullString
'End Sub
'Sub ColorIt(cntrl As Control, tmr As Timer)
'    tmr.Enabled = True
'    cntrl.BackColor = vbRed
'    cntrl.ForeColor = vbYellow
'End Sub
'Function SelectCombo(C As ComboBox, str As String, Optional ByVal ByItemData As Boolean = False) As Integer
'    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
'    Dim i                                    As Long
'    Dim ItemDataX                            As Long
'    If ByItemData = False Then
'        For i = 0 To C.ListCount - 1
'            If UCase(C.List(i)) = UCase(Trim(str)) Then
'                SelectCombo = i
'                Exit Function
'            End If
'        Next
'    Else
'        If str = vbNullString Then
'            SelectCombo = -1
'            Exit Function
'        End If
'        ItemDataX = CLng(str)
'        For i = 0 To C.ListCount - 1
'            If C.ItemData(i) = str Then
'                SelectCombo = i
'                Exit Function
'            End If
'        Next
'    End If
'    SelectCombo = -1
'End Function
'
'
'

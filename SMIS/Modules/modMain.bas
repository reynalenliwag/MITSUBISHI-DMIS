Attribute VB_Name = "modSMISMain"
Option Explicit
Dim rsProfile                                                         As ADODB.Recordset

Public Sub Main()
    If App.PrevInstance = True Then
        MsgBox "There is open SMIS application", vbInformation
        End
    End If
    
    SERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SERVERNAME")
    SQLSERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SQLSERVERNAME")
    'SERVERNAME = "DGIDMISSVR\olddata"
    'SQLSERVERNAME = "DGIDMISSVR\olddata"
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
        .StatusBar1.Panels(9).Text = "Server Name: " & SQLSERVERNAME & "-" & DATABASE
        .StatusBar1.Panels(10).Text = "Company Code: " & COMPANY_CODE
        .StatusBar1.Panels(11).Text = "Rev.: " & App.Revision
    End With
End Sub
Public Function OpenSQLDb() As Boolean
    Screen.MousePointer = 11
    frmSecurity.Hide
    frmSplash.Show
    frmSplash.ZOrder 0
    DoEvents

    ApplySecurityValidation = True
    On Error GoTo ConnErr
    Set gconDMIS = New ADODB.Connection
    gconDMIS.ConnectionString = DMIS_Connection
    gconDMIS.CursorLocation = adUseClient
    gconDMIS.Mode = adModeReadWrite
    frmSplash.labCon.Caption = "Connecting to CMIS Database... Please wait..."
    DoEvents
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
    SetUserPathSettings
    With frmMain
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
    End With
End Sub




Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                                          As String
    Dim cWidth                                                        As Long
    Dim i                                                             As Integer
    Dim scwidth                                                       As Long
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
    Dim fld                                                           As Field
    Dim j                                                             As Long
    Dim ijx                                                           As Integer
    Dim lst                                                           As ListItem
    Dim i                                                             As Integer

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

    Dim fld                                                           As Field
    Dim j                                                             As Long
    Dim REC                                                           As XtremeReportControl.ReportRecord


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
    Dim nrs                                                           As New ADODB.Recordset
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
    Dim ar()                                                          As String
    Dim cWidth                                                        As Long
    Dim i                                                             As Integer

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
    Dim i                                                             As Long
    Dim ItemDataX                                                     As Long
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
    If C.Style = 0 Then
        C.Text = ""
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
    Dim ar()                                                          As String
    Dim i                                                             As Integer


    ar = Split(StringHeaders, ",")
    lst.Columns.DeleteAll
    For i = LBound(ar) To UBound(ar)
        lst.Columns.Add i, ar(i), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString

End Sub

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
'        oBx.Text = xVal
    End If

End Sub

Function GenerateCode(TABLENAME, FLDNAME As String, xFormat As String) As String
    Dim rsID                                                          As ADODB.Recordset

    Set rsID = gconDMIS.Execute("Select MAX( ISNULL(" & FLDNAME & "  , 0) ) as IDFIELD from " & TABLENAME & " where isnumeric(" & FLDNAME & " )=1 ")
    'JJE Prefix 02/08/2013 1:42PM   ** FOR APPROVAL **
'    Set rsID = gconDMIS.Execute("Select MAX(ISNULL(RIGHT(" & FLDNAME & ",6),0)) as IDFIELD from " & TABLENAME & "")
    
    If rsID.Fields(0).Value = 0 Then
        GenerateCode = Format(1, xFormat)
    Else
        GenerateCode = Format(Val(N2Str2Zero(rsID![IDFIELD])) + 1, xFormat)
    End If
    
'    If COMPANY_CODE = "DJM" Then       ** FOR APPROVAL **
'        If FLDNAME = "SO_NO" Then
'            GenerateCode = "SO" + GenerateCode
'        ElseIf FLDNAME = "VI_NO" Then
'            GenerateCode = "VI" + GenerateCode
'        ElseIf FLDNAME = "VDR_NO" Then
'            GenerateCode = "VD" + GenerateCode
'        ElseIf FLDNAME = "PO_NO" Then
'            GenerateCode = "VP" + GenerateCode
'        ElseIf FLDNAME = "CODE" Then
'            GenerateCode = "VR" + GenerateCode
'        End If
'    End If
    'JJE
    Set rsID = Nothing
End Function

Function CheckListItem(lst As ListView, valueCode As String) As Integer

    Dim i                                                             As Integer
    CheckListItem = -1
    For i = 1 To lst.ListItems.Count
        If lst.ListItems(i).Text = valueCode Then
            CheckListItem = i
            Exit Function
        End If
    Next
End Function
'Function FormExist(XXX As String)
'    Dim frm                                                           As Form
'    For Each frm In Forms
'        If (UCase(frm.Name) = UCase(XXX)) Then
'            FormExist = True
'        End If
'    Next
'    Set frm = Nothing
'End Function
Sub UPDATELOGTABLE(TABLENAME, ID)
    Dim SQL                                                           As String
    SQL = "UPDATE " & TABLENAME & " SET "
    SQL = SQL & " USERCODE =" & N2Str2Null(LOGCODE) & ", " & vbCrLf
    SQL = SQL & " LASTUPDATE =" & N2Str2Null(LOGDATE & " " & LOGTIME) & vbCrLf
    SQL = SQL & " WHERE ID=" & ID
    gconDMIS.Execute SQL
End Sub
Sub LoadSignatories(XXX As String)
    Dim rsSignatories                                         As ADODB.Recordset
    Set rsSignatories = gconDMIS.Execute("Select * from SMIS_Signatories where USEDIN='" & XXX & "'")
    If Not rsSignatories.BOF Or Not rsSignatories.EOF Then

        PreparedBy = Null2String(rsSignatories!PreparedBy)
        PreparedByDesig = Null2String(rsSignatories!PreparedByDesig)

        ApprovedBy = Null2String(rsSignatories!SalesApproved)
        SalesApprovedDesig = Null2String(rsSignatories!SalesApprovedDesig)

        CheckedBy = Null2String(rsSignatories!CheckedBy)
        CheckedByDesig = Null2String(rsSignatories!CheckedByDesig)

        SalesDispatcher = Null2String(rsSignatories!SalesDispatcher)
        SalesDispatcherDesig = Null2String(rsSignatories!SalesDispatcherDesig)

        GeneralManager = Null2String(rsSignatories!GeneralManager)
        GeneralManagerDesig = Null2String(rsSignatories!GeneralManagerDesig)

        DeliveredBy = Null2String(rsSignatories!DeliveredBy)
        DeliveredByDesig = Null2String(rsSignatories!DeliveredByDesig)

        FinancingManager = Null2String(rsSignatories!FinancingManager)
        FinancingManagerDesig = Null2String(rsSignatories!FinancingManagerDesig)
    Else
        MessagePop InfoStop, "Missing Section in Signatories", "Signatories Not Found for " & XXX
        
        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='TRANSACTION SLIP' AND MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO smis_signatories (USEDIN,MAINMODULENAME) VALUES('TRANSACTION SLIP','SMIS')")
        End If


        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='JOB REQUEST FORM' AND MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO smis_signatories (USEDIN,MAINMODULENAME) VALUES('JOB REQUEST FORM','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='RELEASE ORDER' AND MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO smis_signatories (USEDIN,MAINMODULENAME) VALUES('RELEASE ORDER','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='RECIEVING REPORT' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('RECIEVING REPORT','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='GATE PASS' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('GATE PASS','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='SALES INVOICE' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('SALES INVOICE','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='SALES ORDER' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('SALES ORDER','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='PURCHASE ORDER' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('PURCHASE ORDER','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='DEBIT MEMO' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('DEBIT MEMO','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='CREDIT MEMO' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('CREDIT MEMO','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='DELIVERY REPORT' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('DELIVERY REPORT','SMIS')")
        End If

        Set rsSignatories = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='STOCK TRANSFER' and MAINMODULENAME='SMIS'")
        If rsSignatories.Fields(0).Value = 0 Then
            gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('STOCK TRANSFER','SMIS')")
        End If

        PreparedBy = ""
        PreparedByDesig = ""
        ApprovedBy = ""
        SalesApprovedDesig = ""
        CheckedBy = ""
        CheckedByDesig = ""
        SalesDispatcher = ""
        SalesDispatcherDesig = ""
        GeneralManager = ""
        GeneralManagerDesig = ""
        DeliveredBy = ""
        DeliveredByDesig = ""
        FinancingManager = ""
        FinancingManagerDesig = ""
    End If
    Set rsSignatories = Nothing
End Sub


Sub USERCODE_LASTUPDATE(TABLE, fld, IDX)
    If LOGSAE <> "" Then
        gconDMIS.Execute ("UPDATE " & TABLE & " SET USERCODE=" & N2Str2Null(LOGSAE) & " , LASTUPDATED=GETDATE() WHERE " & fld & "=" & IDX)

    Else
        gconDMIS.Execute ("UPDATE " & TABLE & " SET USERCODE=" & N2Str2Null(LOGNAME) & " , LASTUPDATED=GETDATE() WHERE " & fld & "=" & IDX)
    End If
End Sub
Function GetSAECode(XXX As String)
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select saecode  from smis_vw_srep where name='" & Replace(XXX, "'", "") & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        GetSAECode = Null2String(temprs!SAECODE)
    End If
    Set temprs = Nothing
End Function

Function SetSAECode(XXX As String)
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select name  from smis_vw_srep where saecode='" & Repleys(XXX) & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        SetSAECode = Null2String(temprs!Name)
    End If
    Set temprs = Nothing
End Function

Function SelectSAE(cbo As ComboBox, XXX)
    Dim i                                                             As Integer
    For i = 0 To cbo.ListCount - 1
        If UCase(cbo.List(i)) = UCase(XXX) Then
            SelectSAE = True
            Exit Function
        End If
    Next
    SelectSAE = False
End Function

Function CheckORNum(YYY As String, xCOunterType) As String
    Dim rsCMIS_OFF_DT                                                 As ADODB.Recordset
    Set rsCMIS_OFF_DT = New ADODB.Recordset
    Set rsCMIS_OFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE TRANTYPE = '" & xCOunterType & "' AND INVOICENO = '" & YYY & "' AND ISNULL(CANCEL,0)=0 AND LEFT(OR_NUM,3) <> 'SOA'")
    If Not rsCMIS_OFF_DT.EOF And Not rsCMIS_OFF_DT.BOF Then
        CheckORNum = UCase(Null2String(rsCMIS_OFF_DT!OR_NUM))
    End If
    Set rsCMIS_OFF_DT = Nothing
End Function

Function CheckSJNum(YYY As String, xCOunterType) As String
    Dim rsAMIS_JournalSJ                                              As ADODB.Recordset
    Set rsAMIS_JournalSJ = New ADODB.Recordset
    Set rsAMIS_JournalSJ = gconDMIS.Execute("Select * from AMIS_JOURNAL_HD WHERE INVOICETYPE = '" & xCOunterType & "' AND INVOICENO = '" & YYY & "' AND STATUS<>'C' AND JTYPE = 'SJ'")
    If Not rsAMIS_JournalSJ.EOF And Not rsAMIS_JournalSJ.BOF Then
        CheckSJNum = UCase(Null2String(rsAMIS_JournalSJ!VOUCHERNO))
    End If
    Set rsAMIS_JournalSJ = Nothing
End Function


Function CheckAPJNum(YYY As String, xCOunterType) As String
    Dim rsAMIS_JournalSJ                                              As ADODB.Recordset
    Set rsAMIS_JournalSJ = New ADODB.Recordset
    Set rsAMIS_JournalSJ = gconDMIS.Execute("Select VOUCHERNO from AMIS_JOURNAL_HD WHERE JTYPE = 'APJ' AND INVOICENO = '" & YYY & "' AND INVOICETYPE = '" & xCOunterType & "' AND STATUS<>'C' ")
    If Not rsAMIS_JournalSJ.EOF And Not rsAMIS_JournalSJ.BOF Then
        CheckAPJNum = UCase(Null2String(rsAMIS_JournalSJ!VOUCHERNO))
    End If
    Set rsAMIS_JournalSJ = Nothing
End Function
Function SetINSLTOCHATTELTPL(xFIELD As String, xINVOICENO As String) As String
Dim rsInvoicing As ADODB.Recordset
Set rsInvoicing = New ADODB.Recordset
rsInvoicing.Open "SELECT " & xFIELD & " AS FIELDNAME FROM SMIS_SALESORDER WHERE VI_NO = '" & xINVOICENO & "' ", gconDMIS, adOpenForwardOnly
If Not rsInvoicing.EOF And Not rsInvoicing.BOF Then
    SetINSLTOCHATTELTPL = SetVendorName(rsInvoicing!FIELDNAME)
End If
Set rsInvoicing = Nothing
End Function
Function SetVendorName(VVV As Variant)
    Dim rsVENDOR As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select CODE,ACCOUNTNAME as nameofvendor from ALL_ENTITY where COMPLET_CODE= " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
End Function
Sub newcol()
Dim SQL As String
    SQL = ""
    SQL = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SMIS_MRRINV_TABLE' AND COLUMN_NAME='LOCATION')" & vbCrLf
    SQL = SQL & "ALTER TABLE SMIS_MRRINV_TABLE" & vbCrLf
    SQL = SQL & "ADD LOCATION   nvarchar(100)"
    gconDMIS.Execute SQL
End Sub

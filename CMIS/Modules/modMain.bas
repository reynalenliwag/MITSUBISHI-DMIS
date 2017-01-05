Attribute VB_Name = "modCMISMain"

Option Explicit
Dim rsProfile                                                       As ADODB.Recordset

Public Sub Main()
    If App.PrevInstance = True Then
        MsgBox "There is open CMIS application", vbInformation
        End
    End If

    SERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SERVERNAME")
'    SERVERNAME = "DMISSERVER\OLDDATA"
'    SQLSERVERNAME = "DMISSERVER\OLDDATA"
    SQLSERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SQLSERVERNAME")
    DATABASE = GetSetting("DMIS 2.0", "SETTINGS", "DATABASE")
    
    If SQLSERVERNAME = "" Or DATABASE = "" Then
        MsgBox "Application not yet Configured. Please Configure Server Setting from DSA.", vbCritical
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
    frmMainMenu.Show
    'UPDATING CODE      : AXP1015200720:33
    LOADSIGNATORIES
    'Upating Code       : AXP-062620071225
    ReminderModule ""
End Sub

Sub LOADSIGNATORIES()
'FUNCTION FEATURE   : TO ADD SIGNATORIES IN THE REPORTS
'DATE STARTED       : 10/15/2007
'LAST UPDATED       : 10/15/2007
'WHO UPDATED        : AXP
'UPDATING CODE      : AXP1015200720:33
'REQUEST NO         : AXP

    Dim temprs                                                      As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select * from all_profile where modulename='" & App.TITLE & "'")
    If Not temprs.EOF Or Not temprs.BOF Then
        PreparedBy = Null2String(temprs!PreparedBy)
        'IssuedBy = Null2String(temprs!IssuedBy)
        CheckedBy = Null2String(temprs!CheckedBy)
        ApprovedBy = Null2String(temprs!ApprovedBy)
        NotedBy = Null2String(temprs!notedby1)
        GeneralManager = Null2String(temprs!GeneralManager)
    End If
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
    
    Dim rsCash_Pos                                                  As ADODB.Recordset
    Set rsCash_Pos = New ADODB.Recordset
    'Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos WHERE CUTDATE = " & N2Date2Null(LOGDATE))
    'Set rsCASH_POS = gconDMIS.Execute("Select * from CMIS_Cash_Pos WHERE TAG = 0 OR TAG IS NULL")
    Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos ORDER BY CUTDATE DESC")
    If rsCash_Pos.EOF And rsCash_Pos.BOF Then
        gconDMIS.Execute ("Insert into CMIS_Cash_Pos " & _
                          "(CUTDATE, FUND, LTO)" & _
                          " values ('" & LOGDATE & _
                          "', " & MAX_PETTYFUND & ", " & MAX_LTOFUND & ")")
                          
        gconDMIS.Execute ("Insert into CMIS_Cash " & _
                          "(CUTDATE)" & _
                          " values ('" & LOGDATE & "')")
        CURRENT_CUTOFF_DATE = LOGDATE
    Else
        CURRENT_CUTOFF_DATE = Null2Date(rsCash_Pos!CUTDATE)
    End If
    Screen.MousePointer = 0
    frmSplash.Command1.Value = True
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

Public Sub SetUserSettings()
    Call SetUserPathSettings
    With frmMain
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
        .StatusBar1.Panels(9).Text = "Server Name: " & SQLSERVERNAME
        .StatusBar1.Panels(10).Text = "Rev. " & App.Revision
    End With
End Sub

Sub PostUnpostOR(xORNUM As String, xMODE_PAYMENT As String, xPOSTUNPOST As String, xRECEIPTS_AMOUNT As Variant, xTYPE_PAYMENT As String)
    Dim rsTMP                                                       As New ADODB.Recordset
    Dim xDate                                                       As String

    If COMPANY_CODE = "DJM" And OR_VAT_NONVAT = "NON-VAT" Then
        'Do nothing SOA transactions does not have to reflect in Cash Position
    Else
        If CheckIfORisDeposited(xORNUM) = True Then                 'DEPOSITED
            If xPOSTUNPOST = "POST" Then
                Set rsTMP = gconDMIS.Execute("SELECT CUTDATE FROM CMIS_OFF_HD WHERE OR_NUM = " & N2Str2Null(xORNUM) & " AND CUTDATE IS NOT NULL")
                If Not (rsTMP.BOF And rsTMP.EOF) Then
                    xDate = Null2String(rsTMP!CUTDATE)
                Else
                    xDate = CURRENT_CUTOFF_DATE
                End If
    
                If xMODE_PAYMENT = "CASH" Then
                    gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                      " CASH = ROUND(CASH,2) + " & xRECEIPTS_AMOUNT & _
                                      " where CUTDATE = '" & xDate & "'")
                ElseIf xMODE_PAYMENT = "CHECK" Then
                    gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                      " [CHECK] = ROUND([CHECK],2) + " & xRECEIPTS_AMOUNT & _
                                      " where CUTDATE = '" & xDate & "'")
                ElseIf xMODE_PAYMENT = "CARD" Then
                    gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                      " CARD = ROUND(CARD,2) + " & xRECEIPTS_AMOUNT & _
                                      " where CUTDATE = '" & xDate & "'")
                End If
            Else
                Set rsTMP = gconDMIS.Execute("SELECT CUTDATE FROM CMIS_OFF_HD WHERE OR_NUM = " & N2Str2Null(xORNUM) & " AND CUTDATE IS NOT NULL")
                If Not (rsTMP.BOF And rsTMP.EOF) Then
                    xDate = Null2String(rsTMP!CUTDATE)
                Else
                    xDate = CURRENT_CUTOFF_DATE
                End If
    
                If xMODE_PAYMENT = "CASH" Then
                    gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                      " CASHDEPO = ROUND(CASHDEPO,2) - " & xRECEIPTS_AMOUNT & _
                                      " where CUTDATE = '" & xDate & "'")
                ElseIf xMODE_PAYMENT = "CHECK" Then
                    gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                      " [CHECKDEPO] = ROUND([CHECKDEPO],2) - " & xRECEIPTS_AMOUNT & _
                                      " where CUTDATE = '" & xDate & "'")
                ElseIf xMODE_PAYMENT = "CARD" Then
                    gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                      " CARDDEPO = ROUND(CARDDEPO,2) - " & xRECEIPTS_AMOUNT & _
                                      " where CUTDATE = '" & xDate & "'")
                End If
    
                gconDMIS.Execute ("UPDATE CMIS_OFF_HD SET DEPOSIT = 0 WHERE OR_NUM = " & N2Str2Null(xORNUM) & "")
                gconDMIS.Execute ("DELETE FROM CMIS_BANKDEPO WHERE OR_NUM = " & N2Str2Null(xORNUM) & "")
            End If
        Else                                                    'NOT YET DEPOSITED
            'Update: ACL 11052009
            If xPOSTUNPOST = "POST" Then
                Set rsTMP = gconDMIS.Execute("SELECT CUTDATE FROM CMIS_OFF_HD WHERE OR_NUM = " & N2Str2Null(xORNUM) & " AND CUTDATE IS NOT NULL")
                If Not (rsTMP.BOF And rsTMP.EOF) Then
                    xDate = Null2String(rsTMP!CUTDATE)
                Else
                    xDate = CURRENT_CUTOFF_DATE
                End If
                If COMPANY_CODE = M_COMPANY_CODE Then
                    If xTYPE_PAYMENT = "FULL" Then
                        If xMODE_PAYMENT = "CASH" Then
                            gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                              " CASH = ROUND(CASH,2) + " & xRECEIPTS_AMOUNT & _
                                              " where CUTDATE = '" & xDate & "'")
                        ElseIf xMODE_PAYMENT = "CHECK" Then
                            gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                              " [CHECK] = ROUND([CHECK],2) + " & xRECEIPTS_AMOUNT & _
                                              " where CUTDATE = '" & xDate & "'")
                        ElseIf xMODE_PAYMENT = "CARD" Then
                            gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                              " CARD = ROUND(CARD,2) + " & xRECEIPTS_AMOUNT & _
                                              " where CUTDATE = '" & xDate & "'")
                        End If
                    ElseIf xTYPE_PAYMENT = "PARTIAL" Then
                        If xMODE_PAYMENT = "CASH" Then
                            gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                              " CASH = ROUND(CASH,2) + " & NumericVal(FINAL_CASH) & _
                                              " where CUTDATE = '" & xDate & "'")
                        ElseIf xMODE_PAYMENT = "CHECK" Then
                            gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                              " [CHECK] = ROUND([CHECK],2) + " & AMOUNT_TENDERED & _
                                              " where CUTDATE = '" & xDate & "'")
                        ElseIf xMODE_PAYMENT = "CARD" Then
                            gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                              " CARD = ROUND(CARD,2) + " & AMOUNT_TENDERED & _
                                              " where CUTDATE = '" & xDate & "'")
                        End If
                    End If
                Else
                    If xMODE_PAYMENT = "CASH" Then
                        gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                          " CASH = ROUND(CASH,2) + " & xRECEIPTS_AMOUNT & _
                                          " where CUTDATE = '" & xDate & "'")
                    ElseIf xMODE_PAYMENT = "CHECK" Then
                        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                          " [CHECK] = ROUND([CHECK],2) + " & xRECEIPTS_AMOUNT & _
                                          " where CUTDATE = '" & xDate & "'")
                    ElseIf xMODE_PAYMENT = "CARD" Then
                        gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                                          " CARD = ROUND(CARD,2) + " & xRECEIPTS_AMOUNT & _
                                          " where CUTDATE = '" & xDate & "'")
                    End If
                End If
            End If
        End If
    End If
End Sub

Function CheckIfORisDeposited(xORNO As String) As Boolean
    Dim rsTMP                                                       As New ADODB.Recordset
    Set rsTMP = gconDMIS.Execute("SELECT OR_NUM FROM CMIS_BANKDEPO WHERE OR_NUM = " & N2Str2Null(xORNO) & "")
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        CheckIfORisDeposited = True
    End If
    Set rsTMP = Nothing
End Function

Function CheckTotalPayment(xOR_Num As String, xVAT_OR As Integer) As Double
    Dim rsCheckPayment                                              As ADODB.Recordset
    Set rsCheckPayment = New ADODB.Recordset
    rsCheckPayment.Open "Select OR_AMT,CASHAMOUNT,CHKAMOUNT,CARDAMOUNT from CMIS_OFF_HD where OR_NUM = '" & xOR_Num & "' AND VAT = '" & xVAT_OR & "'", gconDMIS, adOpenKeyset
    If Not rsCheckPayment.EOF And Not rsCheckPayment.BOF Then
        CheckTotalPayment = NumericVal(rsCheckPayment!CashAmount) + NumericVal(rsCheckPayment!CHKAMOUNT) + NumericVal(rsCheckPayment!CardAmount)
    End If
End Function

Function CheckIfBank(xCusCde As String) As Boolean
    Dim rsCheckCode                                                 As ADODB.Recordset
    Set rsCheckCode = New ADODB.Recordset
    rsCheckCode.Open "Select Cuscde from All_Customer_Table where CusCde = " & N2Str2Null(xCusCde) & "", gconDMIS, adOpenForwardOnly
    If Not rsCheckCode.EOF And Not rsCheckCode.BOF Then
        Do While Not rsCheckCode.EOF
            Dim rsCheckBank                                         As ADODB.Recordset
            Set rsCheckBank = New ADODB.Recordset
            rsCheckBank.Open "Select CusCde from CMIS_CardBank where CusCde = " & N2Str2Null(rsCheckCode!CUSCDE) & "", gconDMIS, adOpenForwardOnly
            If Not rsCheckBank.EOF And Not rsCheckBank.BOF Then
                CheckIfBank = True
            Else
                CheckIfBank = False
            End If
            rsCheckCode.MoveNext
        Loop
    End If
    Set rsCheckCode = Nothing
    Set rsCheckBank = Nothing
End Function

Function CheckBankName(xCusCde As String) As String
    Dim rsCheckCode                                                 As ADODB.Recordset
    Set rsCheckCode = New ADODB.Recordset
    rsCheckCode.Open "Select Cuscde from All_Customer_Table where CusCde = " & N2Str2Null(xCusCde) & "", gconDMIS, adOpenForwardOnly
    If Not rsCheckCode.EOF And Not rsCheckCode.BOF Then
        Do While Not rsCheckCode.EOF
            Dim rsCheckBank                                         As ADODB.Recordset
            Set rsCheckBank = New ADODB.Recordset
            rsCheckBank.Open "Select CusCde from CMIS_CardBank where CusCde = " & N2Str2Null(rsCheckCode!CUSCDE) & "", gconDMIS, adOpenForwardOnly
            If Not rsCheckBank.EOF And Not rsCheckBank.BOF Then
                CheckBankName = N2Str2Null(rsCheckBank!CUSCDE)
            End If
            rsCheckCode.MoveNext
        Loop
    End If
    Set rsCheckCode = Nothing
    Set rsCheckBank = Nothing
End Function

Sub SaveCashPosition(XTYPE As String, xAMOUNT As Currency)
    If XTYPE = "CASH" Then
        gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                          " CASH = CASH - " & xAMOUNT & "," & _
                          " CASHDEPO = CASHDEPO + " & xAMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    ElseIf XTYPE = "CHECK" Then
        gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                          " [CHECK] = [CHECK] - " & xAMOUNT & "," & _
                          " CHECKDEPO = CHECKDEPO + " & xAMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    ElseIf XTYPE = "CARD" Then
        gconDMIS.Execute ("Update CMIS_Cash_Pos Set" & _
                          " CARD = CARD - " & xAMOUNT & "," & _
                          " CARDDEPO = CARDDEPO + " & xAMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    End If
End Sub

Function CheckCutoff(xCutoffDate) As Boolean
    Dim rsProcessCutOff                                             As ADODB.Recordset
    Set rsProcessCutOff = New ADODB.Recordset
    rsProcessCutOff.Open "SELECT DISTINCT CUTDATE FROM CMIS_OFF_HD WHERE CUTDATE IN (SELECT CUTDATE FROM CMIS_CASH_POS WHERE CUTDATE='" & CDate(xCutoffDate) & "') and CUTDATE IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsProcessCutOff.EOF And Not rsProcessCutOff.BOF Then
        CheckCutoff = True
    End If
End Function

Sub ReportControlAddColumnHeader(LST As ReportControl, StringHeaders As String)
    Dim ar()                                                        As String
    Dim i                                                           As Integer

    ar = Split(StringHeaders, ",")
    LST.Columns.DeleteAll
    For i = LBound(ar) To UBound(ar)
        LST.Columns.Add i, ar(i), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False
    Dim ar()                                                        As String
    Dim cWidth                                                      As Long
    Dim i                                                           As Integer
    Dim scwidth                                                     As Long
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

Public Sub AddColumnHeader(StringHeaders As String, lvGrid As ListView)
    Dim ar()                                                        As String
    Dim cWidth                                                      As Long
    Dim i                                                           As Integer

    ar = Split(StringHeaders, ",")
    cWidth = lvGrid.Width
    lvGrid.ColumnHeaders.Clear
    For i = LBound(ar) To UBound(ar)
        lvGrid.ColumnHeaders.Add , , ar(i)
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

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

Function SelectCombo(C As ComboBox, STR As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim i                                                           As Long
    Dim ItemDataX                                                   As Long
    
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

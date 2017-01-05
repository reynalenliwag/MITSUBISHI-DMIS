VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMS_Report_SubletSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sublet Sales Report"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_Report_SubletSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   3480
   Begin VB.ComboBox cboType 
      Height          =   330
      ItemData        =   "frmCSMS_Report_SubletSales.frx":1082
      Left            =   810
      List            =   "frmCSMS_Report_SubletSales.frx":108C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   90
      Width           =   2565
   End
   Begin VB.ComboBox cboYEar 
      Height          =   330
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   900
      Width           =   2085
   End
   Begin VB.ComboBox cboMOnth 
      Height          =   330
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2565
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1875
      MouseIcon       =   "frmCSMS_Report_SubletSales.frx":10AD
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_Report_SubletSales.frx":11FF
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1350
      Width           =   735
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   2820
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1155
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_Report_SubletSales.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1350
      Width           =   735
   End
   Begin VB.Label labCap 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   300
      TabIndex        =   7
      Top             =   180
      Width           =   405
   End
   Begin VB.Label labCap 
      AutoSize        =   -1  'True
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   330
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label labCap 
      AutoSize        =   -1  'True
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   570
      Width           =   525
   End
End
Attribute VB_Name = "frmCSMS_Report_SubletSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xCUST                                           As Double
Dim xWARR                                           As Double
Dim xINSU                                           As Double
Dim xSALE                                           As Double
Dim xCOMP                                           As Double
Dim TOT_PARTS                                       As Double
Dim TOT_MATER                                       As Double
Dim TOT_LABOR                                       As Double
Dim TOT_CUSTO                                       As Double
Dim TOT_WARRA                                       As Double
Dim TOT_INSUR                                       As Double
Dim TOT_SALES                                       As Double
Dim TOT_COMPA                                       As Double
Dim TOT_ROAMT                                       As Double
Dim TOT_SUBLE                                       As Double
Dim TEMP_INS                                        As Double
Dim xACCT_CODE                                      As String
Dim xACCT_DESC                                      As String
Dim TOTGJ_SUBPARTS                                  As Double
Dim TOTGJ_SUBMAT                                    As Double
Dim TOTGJ_SUBLABOR                                  As Double
Dim TOTBP_SUBPARTS                                  As Double
Dim TOTBP_SUBMAT                                    As Double
Dim TOTBP_SUBLABOR                                  As Double

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call PrintInExcel
    Exit Sub
    
'    Dim RSTMP                   As New ADODB.Recordset
'    Set RSTMP = gconDMIS.Execute("SELECT DTE_COMP FROM CSMS_REPOR WHERE " & _
'        " DTE_COMP IS NOT NULL " & _
'        " AND DTE_COMP BETWEEN " & N2Str2Null(txtFROM) & " AND " & N2Str2Null(txtTO) & "")
'    If Not (RSTMP.BOF And RSTMP.EOF) Then
'        On Error GoTo ERROR_MSG
'        rpt.WindowTitle = "Sublet Sales Report"
'        rpt.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
'        rpt.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
'        rpt.Formulas(2) = "Printedby = '" & LOGNAME & "'"
'        rpt.Formulas(3) = "RANGEDATE = '" & txtFROM & "-" & txtTO & "'"
'
'        PrintSQLReport rpt, CSMS_REPORT_PATH & "Sublet Sales Report.RPT", "{REPOR.DTE_COMP} >= DATE(" & Year(txtFROM.Value) & "," & Month(txtFROM.Value) & "," & Day(txtFROM.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
'
'        'NEW LOG AUDIT-----------------------------------------------------
'            Call NEW_LogAudit("V", "SUBLET SALES REPORTS", "", "", "", "SUMMARY - " & txtFROM & " TO " & txtTO, "", "")
'        'NEW LOG AUDIT-----------------------------------------------------
'    Else
'        Call ShowNoRecord
'    End If
'    Set RSTMP = Nothing
    
    Exit Sub
ERROR_MSG:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Call fillcbomonth(cboMOnth)
    Call FillCboMoreYear(cboYEar)
    
    cboType.ListIndex = 0
    cboMOnth.Text = MonthName(Month(Date))
    'TXTFrom.Value = firstDay(Date)
    'txtTO.Value = Date
    
    Screen.MousePointer = 0
End Sub

Sub PrintInExcel()
    Screen.MousePointer = 11
    Dim rstmp                                               As New ADODB.Recordset
    Dim xlApp                                               As Excel.Application
    Dim xlBook                                              As Excel.Workbook
    Dim xlSheet                                             As Excel.Worksheet
    Dim cnt                                                 As Integer
    Dim rsHEAD                                              As New ADODB.Recordset
    Dim rsDet                                               As New ADODB.Recordset
    
    cnt = 7
'
    Set rsHEAD = gconDMIS.Execute("SELECT DTE_REL, INSAMT, INVOICE, REP_OR, AMOUNT, RO_AMOUNT FROM CSMS_REPOR " & _
        " WHERE MONTH(DTE_REL) = " & What_month(cboMOnth) & _
        " AND YEAR(DTE_REL) = " & cboYEar & "")
        
    'R-00075430

'    Set rsHEAD = gconDMIS.Execute("SELECT DTE_COMP, INSAMT, INVOICE, REP_OR, AMOUNT, RO_AMOUNT FROM CSMS_REPOR " & _
        " WHERE REP_OR = 'R-00075428'")
        
    TOT_ROAMT = 0
    TOT_SUBLE = 0
    
    TOT_PARTS = 0
    TOT_MATER = 0
    TOT_LABOR = 0
    
'updated by:     IEBV 08162010_0113pm
    TOTGJ_SUBPARTS = 0
    TOTGJ_SUBMAT = 0
    TOTGJ_SUBLABOR = 0
    TOTBP_SUBPARTS = 0
    TOTBP_SUBMAT = 0
    TOTBP_SUBLABOR = 0
'-------------------------------------------------------------------------------
    
    TOT_CUSTO = 0
    TOT_WARRA = 0
    TOT_INSUR = 0
    TOT_SALES = 0
    TOT_COMPA = 0
    
    If (rsHEAD.BOF And rsHEAD.EOF) Then
        Call ShowNoRecord
        Set rsHEAD = Nothing
        Exit Sub
    Else
        On Error Resume Next
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Sublet Sales.xlt")
        Set xlSheet = xlBook.Worksheets(1)
        
        xlSheet.Cells(1, "A") = COMPANY_NAME
        xlSheet.Cells(2, "A") = COMPANY_ADDRESS
        If cboType.ListIndex = 0 Then
            xlSheet.Cells(3, "A") = "Report of Sublet Sales"
        Else
            xlSheet.Cells(3, "A") = "Report of Sublet COS"
        End If
        xlSheet.Cells(4, "A") = "For the Month of " & cboMOnth & " " & cboYEar
        Do While Not rsHEAD.EOF
            TEMP_INS = NumericVal(rsHEAD!INSAMT)
            If CheckIfDetailsHaveSublet(Null2String(rsHEAD!REP_OR)) = False Then
                GoTo NEXT_RECORD
            End If
    
            xlSheet.Cells(cnt, "A") = Null2String(rsHEAD!dte_rel)
            xlSheet.Cells(cnt, "B") = Null2String(rsHEAD!invoice)
            xlSheet.Cells(cnt, "C") = Null2String(rsHEAD!REP_OR)
            xlSheet.Cells(cnt, "D") = GetRepairOrderAmount(Null2String(rsHEAD!REP_OR))
            TOT_ROAMT = TOT_ROAMT + xlSheet.Cells(cnt, "D")
            xlSheet.Cells(cnt, "E") = ComputeTotalSubletAmount(Null2String(rsHEAD!REP_OR))
            TOT_SUBLE = TOT_SUBLE + xlSheet.Cells(cnt, "E")
            
            xlSheet.Cells(cnt, "F") = GetSubletAmountPerType("2", Null2String(rsHEAD!REP_OR))
            TOT_PARTS = TOT_PARTS + xlSheet.Cells(cnt, "F")
            xlSheet.Cells(cnt, "G") = GetSubletAmountPerType("3", Null2String(rsHEAD!REP_OR))
            TOT_MATER = TOT_MATER + xlSheet.Cells(cnt, "G")
            xlSheet.Cells(cnt, "H") = GetSubletAmountPerType("1", Null2String(rsHEAD!REP_OR))
            TOT_LABOR = TOT_LABOR + xlSheet.Cells(cnt, "H")
'updated by:    IEBV 08162010_0113pm
'description:   To display the sublet amount of the PO
'----------------------------------------------------------------------------------------------------
            xlSheet.Cells(cnt, "I") = getSubletdetails_BIlling("1", "GJ", Null2String(rsHEAD!REP_OR))
            TOTGJ_SUBLABOR = TOTGJ_SUBLABOR + xlSheet.Cells(cnt, "I")
            xlSheet.Cells(cnt, "J") = getSubletdetails_BIlling("3", "GJ", Null2String(rsHEAD!REP_OR))
            TOTGJ_SUBMAT = TOTGJ_SUBMAT + xlSheet.Cells(cnt, "J")
            xlSheet.Cells(cnt, "K") = getSubletdetails_BIlling("2", "GJ", Null2String(rsHEAD!REP_OR))
            TOTGJ_SUBPARTS = TOTGJ_SUBPARTS + xlSheet.Cells(cnt, "K")
            xlSheet.Cells(cnt, "L") = getSubletdetails_BIlling("1", "BP", Null2String(rsHEAD!REP_OR))
            TOTBP_SUBLABOR = TOTBP_SUBLABOR + xlSheet.Cells(cnt, "L")
            xlSheet.Cells(cnt, "M") = getSubletdetails_BIlling("3", "BP", Null2String(rsHEAD!REP_OR))
            TOTBP_SUBMAT = TOTBP_SUBMAT + xlSheet.Cells(cnt, "M")
            xlSheet.Cells(cnt, "N") = getSubletdetails_BIlling("2", "BP", Null2String(rsHEAD!REP_OR))
            TOTBP_SUBPARTS = TOTBP_SUBPARTS + xlSheet.Cells(cnt, "N")
 '----------------------------------------------------------------------------------------------------
            
            Call ComputeChargeTo(Null2String(rsHEAD!REP_OR), NumericVal(rsHEAD!INSAMT))
            xlSheet.Cells(cnt, "O") = xCUST
            xlSheet.Cells(cnt, "P") = xWARR
            xlSheet.Cells(cnt, "Q") = xINSU
            xlSheet.Cells(cnt, "R") = xSALE
            xlSheet.Cells(cnt, "S") = xCOMP
            
            TOT_CUSTO = TOT_CUSTO + xCUST
            TOT_WARRA = TOT_WARRA + xWARR
            TOT_INSUR = TOT_INSUR + xINSU
            TOT_SALES = TOT_SALES + xSALE
            TOT_COMPA = TOT_COMPA + xCOMP
            xlSheet.Cells(cnt, "T") = GetInternalDescription(Null2String(rsHEAD!REP_OR))
            xACCT_CODE = "":            xACCT_DESC = ""
            Call GetInternalDescription_ACCT(Null2String(rsHEAD!REP_OR))
            
            xlSheet.Cells(cnt, "U") = xACCT_CODE
            xlSheet.Cells(cnt, "V") = xACCT_DESC
            cnt = cnt + 1
            
NEXT_RECORD:
            rsHEAD.MoveNext
        Loop
        
        xlSheet.Cells(cnt, "D") = TOT_ROAMT
        xlSheet.Cells(cnt, "E") = TOT_SUBLE
        xlSheet.Cells(cnt, "F") = TOT_PARTS
        xlSheet.Cells(cnt, "G") = TOT_MATER
        xlSheet.Cells(cnt, "H") = TOT_LABOR
'updated by:    IEBV 08162010_0113pm
'description:   To display the sublet total amount of the PO
'------------------------------------------------------------------------------------------
        xlSheet.Cells(cnt, "I") = TOTGJ_SUBLABOR
        xlSheet.Cells(cnt, "J") = TOTGJ_SUBMAT
        xlSheet.Cells(cnt, "K") = TOTGJ_SUBPARTS
        xlSheet.Cells(cnt, "L") = TOTBP_SUBLABOR
        xlSheet.Cells(cnt, "M") = TOTBP_SUBMAT
        xlSheet.Cells(cnt, "N") = TOTBP_SUBPARTS
'------------------------------------------------------------------------------------------
        xlSheet.Cells(cnt, "O") = TOT_CUSTO
        xlSheet.Cells(cnt, "P") = TOT_WARRA
        xlSheet.Cells(cnt, "Q") = TOT_INSUR
        xlSheet.Cells(cnt, "R") = TOT_SALES
        xlSheet.Cells(cnt, "S") = TOT_COMPA
        
        xlSheet.Range("D" & cnt & ":" & "S" & cnt).Font.Bold = True
        xlSheet.Range("D" & cnt & ":" & "S" & cnt).Borders(xlBottom).LineStyle = xlDouble
        xlSheet.Range("D" & cnt & ":" & "S" & cnt).Borders(xlTop).LineStyle = xlDouble
        
    End If
    xlApp.Windows.ITEM(1).Caption = "Sublet Sales Report"
    xlApp.Visible = True
    
    Set xlApp = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    
    Screen.MousePointer = 0
    Exit Sub
    
ERROR_MSG:
    If Err.Number = 1004 Then
        MsgBox "Report File not found", vbCritical, "Error"
        'MsgBox Err.Number & " - " & Err.Description
    Else
        MsgBox "Unknown Error occured", vbCritical, "Error"
    End If
    
    Err.Clear
End Sub

Function CheckIfDetailsHaveSublet(xREPOR As String) As Boolean
    Dim rstmp                                               As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("selecT ROTYPE FROM CSMS_RO_DET WHERE ISNULL(ROTYPE,'') = 'SR' AND REP_OR = " & N2Str2Null(xREPOR) & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        CheckIfDetailsHaveSublet = True
    Else
        CheckIfDetailsHaveSublet = False
    End If
    Set rstmp = Nothing
End Function

Sub GetInternalDescription_ACCT(xREPOR As String)
    Dim rstmp                                               As New ADODB.Recordset
    Dim RSREP                                               As New ADODB.Recordset
    Dim rsChartAccount                                      As New ADODB.Recordset
    Dim xDESC                                               As String
    
    Set RSREP = gconDMIS.Execute("SELECT CODE FROM CSMS_RO_DET " & _
        " WHERE REP_OR = " & N2Str2Null(xREPOR) & _
        " AND ROTYPE = 'SR'")
    If Not (RSREP.BOF And RSREP.EOF) Then
        Do While Not RSREP.EOF
            Set rstmp = gconDMIS.Execute("SELECT CHARTCODES FROM CMIS_CBOOK WHERE BOOK = 'S' " & _
                " AND CODE = " & N2Str2Null(RSREP!Code) & "")
            If Not (rstmp.BOF And rstmp.EOF) Then
                Set rsChartAccount = gconDMIS.Execute("Select ACCTCODE, DESCRIPTION from AMIS_ChartAccount Where " & _
                    " AcctCode = " & N2Str2Null(rstmp!CHARTCODES) & "")
                If Not (rsChartAccount.BOF And rsChartAccount.EOF) Then
                    xACCT_CODE = xACCT_CODE & Null2String(rsChartAccount!ACCTCODE) & vbCrLf
                    xACCT_DESC = xACCT_DESC & Null2String(rsChartAccount!Description) & vbCrLf
                End If
                Set rsChartAccount = Nothing
            End If
            RSREP.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Function GetInternalDescription(xREPOR As String) As String
    Dim rstmp                                               As New ADODB.Recordset
    Dim RSREP                                               As New ADODB.Recordset
    Dim xDESC                                               As String
    
    Set RSREP = gconDMIS.Execute("SELECT CODE FROM CSMS_RO_DET " & _
        " WHERE REP_OR = " & N2Str2Null(xREPOR) & _
        " AND ROTYPE = 'SR'")
    If Not (RSREP.BOF And RSREP.EOF) Then
        Do While Not RSREP.EOF
            Set rstmp = gconDMIS.Execute("SELECT DESCNAME FROM CMIS_CBOOK WHERE BOOK = 'S' " & _
                " AND CODE = " & N2Str2Null(RSREP!Code) & "")
            If Not (rstmp.BOF And rstmp.EOF) Then
                xDESC = xDESC & Null2String(rstmp!DESCNAME) & vbCrLf
            End If
            RSREP.MoveNext
        Loop
    End If
    GetInternalDescription = xDESC
    Set rstmp = Nothing
End Function

Function GetSubletAmountPerType(XTYPE As String, xREPOR As String) As Double
    Dim rstmp                                               As New ADODB.Recordset
    Dim RSHD                                                As New ADODB.Recordset
    Dim RSDT                                                As New ADODB.Recordset
    Dim RSRO                                                As New ADODB.Recordset
    Dim RESULT                                              As Double
    
    If cboType.ListIndex = 0 Then
        Set RSHD = gconDMIS.Execute("SELECT SUM(ISNULL(DET_AMT,0)) AS RESULT FROM CSMS_RO_DET WHERE " & _
            " LIVIL = " & XTYPE & _
            " AND REP_OR = " & N2Str2Null(xREPOR) & _
            " AND ROTYPE = 'SR'")
        If Not (RSHD.BOF And RSHD.EOF) Then
            GetSubletAmountPerType = NumericVal(RSHD!RESULT)
        End If
    Else
        Set RSDT = gconDMIS.Execute("SELECT PO_NO FROM CSMS_PO_HD WHERE " & _
            " RO_NO = " & N2Str2Null(xREPOR) & _
            " AND STATUS = 'P'")
        If Not (RSDT.BOF And RSDT.EOF) Then
            Do While Not RSDT.EOF
                Set RSRO = gconDMIS.Execute("SELECT LINE_NO FROM CSMS_RO_DET WHERE " & _
                    " REP_OR = " & N2Str2Null(xREPOR) & _
                    " AND SUBPOCODE = " & N2Str2Null(RSDT!PO_NO) & _
                    " AND ROTYPE = 'SR'")
                If Not (RSRO.BOF And RSRO.EOF) Then
                    Do While Not RSRO.EOF
                        Set rstmp = gconDMIS.Execute("SELECT SUM(ISNULL(CONTRACTAMOUNT,0)) AS RESULT FROM CSMS_PO_DT WHERE " & _
                            " PO_NO = " & N2Str2Null(RSDT!PO_NO) & _
                            " AND LIVIL = " & N2Str2Null(XTYPE) & _
                            " AND LINE_NO = " & N2Str2Null(RSRO!LINE_NO) & "")
                        If Not (rstmp.BOF And rstmp.EOF) Then
                            RESULT = RESULT + NumericVal(rstmp!RESULT)
                        End If
                        Set rstmp = Nothing
                        RSRO.MoveNext
                    Loop
                End If
                
                RSDT.MoveNext
            Loop
        End If
        GetSubletAmountPerType = RESULT
    End If
    Set RSHD = Nothing
End Function

Sub ComputeChargeTo(xREPOR As String, xINS_AMT As Double)
    Dim rstmp                                               As New ADODB.Recordset
    Dim rsDet                                               As New ADODB.Recordset
    Dim RSHD                                                As New ADODB.Recordset
    Dim INST                                                As Double
    
    INST = xINS_AMT
    xCUST = 0:    xWARR = 0
    xSALE = 0:    xCOMP = 0:    xINSU = 0
    
    If cboType.ListIndex = 0 Then
        Set rstmp = gconDMIS.Execute("SELECT ISNULL(WCODE,'') AS WCODE, ISNULL(DET_AMT,0) AS DET_AMT FROM CSMS_RO_DET " & _
            " WHERE REP_OR = " & N2Str2Null(xREPOR) & _
            " AND ROTYPE = 'SR'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            Do While Not rstmp.EOF
                If INST > 0 Then
                    If INST > NumericVal(rstmp!DET_AMT) Then
                        INST = INST - NumericVal(rstmp!DET_AMT)
                        xINSU = xINSU + NumericVal(rstmp!DET_AMT)
                    Else
                        xINSU = xINSU + INST
                        xCUST = xCUST + (NumericVal(rstmp!DET_AMT) - INST)
                        INST = 0
                    End If
                Else
                    If Null2String(rstmp!wCode) = "" Then
                        xCUST = xCUST + NumericVal(rstmp!DET_AMT)
                    ElseIf Null2String(rstmp!wCode) = "W" Then
                        xWARR = xWARR + NumericVal(rstmp!DET_AMT)
                    ElseIf Null2String(rstmp!wCode) = "S" Then
                        xSALE = xSALE + NumericVal(rstmp!DET_AMT)
                    Else
                        xCOMP = xCOMP + NumericVal(rstmp!DET_AMT)
                    End If
                End If
                
                rstmp.MoveNext
            Loop
        End If
    Else
        Set RSHD = gconDMIS.Execute("SELECT PO_NO FROM CSMS_PO_HD WHERE " & _
            " RO_NO = " & N2Str2Null(xREPOR) & " AND STATUS = 'P'")
        If Not (RSHD.BOF And RSHD.EOF) Then
            Do While Not RSHD.EOF
                Set rstmp = gconDMIS.Execute("SELECT LINE_NO, ISNULL(WCODE,'') AS WCODE, ISNULL(DET_AMT,0) AS DET_AMT FROM CSMS_RO_DET " & _
                    " WHERE REP_OR = " & N2Str2Null(xREPOR) & _
                    " AND ROTYPE = 'SR' " & _
                    " AND SUBPOCODE = " & N2Str2Null(RSHD!PO_NO) & "")
                If Not (rstmp.BOF And rstmp.EOF) Then
                    Do While Not rstmp.EOF
                        Set rsDet = gconDMIS.Execute("SELECT ISNULL(CONTRACTAMOUNT,0) AS DET_AMT FROM CSMS_PO_DT " & _
                            " WHERE PO_NO = " & N2Str2Null(RSHD!PO_NO) & _
                            " AND STATUS = 'P' " & _
                            " AND LINE_NO = " & N2Str2Null(rstmp!LINE_NO) & "")
                        If Not (rsDet.BOF And rsDet.EOF) Then
                            If INST > 0 Then
                                If INST > NumericVal(rsDet!DET_AMT) Then
                                    INST = INST - NumericVal(rsDet!DET_AMT)
                                    xINSU = xINSU + NumericVal(rsDet!DET_AMT)
                                Else
                                    xINSU = xINSU + INST
                                    xCUST = xCUST + (NumericVal(rsDet!DET_AMT) - INST)
                                    INST = 0
                                End If
                            Else
                                If Null2String(rstmp!wCode) = "" Then
                                    xCUST = xCUST + NumericVal(rsDet!DET_AMT)
                                ElseIf Null2String(rstmp!wCode) = "W" Then
                                    xWARR = xWARR + NumericVal(rsDet!DET_AMT)
                                ElseIf Null2String(rstmp!wCode) = "S" Then
                                    xSALE = xSALE + NumericVal(rsDet!DET_AMT)
                                Else
                                    xCOMP = xCOMP + NumericVal(rsDet!DET_AMT)
                                End If
                            End If
                        End If
                        Set rsDet = Nothing
                                        
                        rstmp.MoveNext
                        
                    Loop
                End If
                Set rstmp = Nothing
                
                RSHD.MoveNext
            Loop
        End If
    End If

    Set RSHD = Nothing
End Sub

Function ComputeTotalSubletAmount(xREPOR As String) As Double
    Dim rstmp                                           As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT SUM(ISNULL(DET_AMT,0)) AS RESULT FROM CSMS_RO_DET " & _
        " WHERE REP_OR = " & N2Str2Null(xREPOR) & _
        " AND ROTYPE = 'SR'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        ComputeTotalSubletAmount = NumericVal(rstmp!RESULT)
    End If
    Set rstmp = Nothing
End Function

Function GetRepairOrderAmount(xREPOR As String) As Double
    Dim rstmp                                           As New ADODB.Recordset
    
    Set rstmp = gconDMIS.Execute("SELECT SUM(ISNULL(DET_AMT,0)) AS RESULT FROM CSMS_RO_DET WHERE " & _
        " REP_OR = " & N2Str2Null(xREPOR) & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GetRepairOrderAmount = NumericVal(rstmp!RESULT)
    End If
    Set rstmp = Nothing
End Function
'updated by:    IEBV 08162010_0113pm
'description:   To get the sublet amount of the PO
Function getSubletdetails(xLIVIL As Integer, xClassification As String, xREPOR As Variant) As Double
    Dim rssubletdetail                                  As New ADODB.Recordset
    Set rssubletdetail = New ADODB.Recordset
        If cboType.ListIndex = 0 Then
        Set rssubletdetail = gconDMIS.Execute("SELECT SUM(ISNULL(DET_AMT,0)) AS RESULT FROM csms_po_dt WHERE " & _
            " LIVIL = " & xLIVIL & _
            " AND REP_OR = " & N2Str2Null(xREPOR) & _
            " AND status =  'P' " & _
            " AND Jobtype = " & N2Str2Null(xClassification))
            If Not (rssubletdetail.BOF And rssubletdetail.EOF) Then
                rssubletdetail.MoveFirst
                Do While Not rssubletdetail.EOF
                    getSubletdetails = getSubletdetails + NumericVal(rssubletdetail!RESULT)
                    rssubletdetail.MoveNext
                Loop
            End If
        End If
    Set rssubletdetail = Nothing
End Function
'-------------------------------------------------------------------------------------------------------------------


'updated by:    IEBV 08162010_0113pm
'description:   To get the sublet amount of the PO
Function getSubletdetails_BIlling(xLIVIL As Integer, xClassification As String, xREPOR As Variant) As Double
    Dim rssubletdetail                                  As New ADODB.Recordset
    Set rssubletdetail = New ADODB.Recordset
        If cboType.ListIndex = 0 Then
        Set rssubletdetail = gconDMIS.Execute("SELECT SUM(ISNULL(DET_AMT,0)) AS RESULT FROM csms_ro_det WHERE " & _
            " LIVIL = " & xLIVIL & _
            " AND REP_OR = " & N2Str2Null(xREPOR) & _
            " AND status in ('R') " & _
            " AND ROTYPE =  'SR' " & _
            " AND Jobtype = " & N2Str2Null(xClassification))
            If Not (rssubletdetail.BOF And rssubletdetail.EOF) Then
                rssubletdetail.MoveFirst
                Do While Not rssubletdetail.EOF
                    getSubletdetails_BIlling = getSubletdetails_BIlling + NumericVal(rssubletdetail!RESULT)
                    rssubletdetail.MoveNext
                Loop
            End If
        End If
    Set rssubletdetail = Nothing
End Function
'-------------------------------------------------------------------------------------------------------------------





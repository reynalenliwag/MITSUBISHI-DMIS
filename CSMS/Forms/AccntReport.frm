VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSServiceReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Report"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   FillColor       =   &H8000000D&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AccntReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   4530
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   345
      Left            =   150
      TabIndex        =   1
      Top             =   900
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   98697217
      CurrentDate     =   39484
   End
   Begin VB.ComboBox cboReportType 
      Appearance      =   0  'Flat
      Height          =   345
      ItemData        =   "AccntReport.frx":1082
      Left            =   150
      List            =   "AccntReport.frx":1084
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   4305
   End
   Begin Crystal.CrystalReport rptReports 
      Left            =   3960
      Top             =   2010
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2220
      MouseIcon       =   "AccntReport.frx":1086
      MousePointer    =   99  'Custom
      Picture         =   "AccntReport.frx":11D8
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   1650
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1500
      MouseIcon       =   "AccntReport.frx":1623
      MousePointer    =   99  'Custom
      Picture         =   "AccntReport.frx":1775
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   1650
      Width           =   735
   End
   Begin VB.PictureBox picCSHCHG 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   4335
      TabIndex        =   15
      Top             =   1260
      Width           =   4335
      Begin VB.OptionButton Option1 
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3690
         TabIndex        =   6
         Top             =   60
         Width           =   1335
      End
      Begin VB.OptionButton optINT 
         Caption         =   "INTERNAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2280
         TabIndex        =   5
         Top             =   60
         Width           =   1335
      End
      Begin VB.OptionButton optCharge 
         Caption         =   "CHARGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   4
         Top             =   60
         Width           =   1155
      End
      Begin VB.OptionButton optCash 
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.CheckBox chkSummary 
      Caption         =   "Summary Only"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1290
      TabIndex        =   9
      Top             =   1260
      Width           =   2025
   End
   Begin MSComCtl2.DTPicker txtTo 
      Height          =   345
      Left            =   2460
      TabIndex        =   2
      Top             =   900
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   98697217
      CurrentDate     =   39484
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "CHOOSE A REPORT TYPE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   180
      TabIndex        =   16
      Top             =   60
      Width           =   1950
   End
   Begin VB.Label labRep_or 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Top             =   1080
      Width           =   4425
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Completed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   10770
      TabIndex        =   13
      Top             =   1110
      Width           =   825
   End
   Begin VB.Label labProgress 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   10200
      TabIndex        =   12
      Top             =   1110
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   150
      TabIndex        =   11
      Top             =   690
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2460
      TabIndex        =   10
      Top             =   690
      Width           =   225
   End
End
Attribute VB_Name = "frmCSMSServiceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPOR                                            As ADODB.Recordset

Private Sub cboReportType_Click()
    Dim A
    A = UCase(cboReportType.Text)

    If A = "SERVICE INVOICE SUMMARY REPORT" Then
        chkSummary.Visible = False: picCSHCHG.Visible = True
    ElseIf A = "SERVICE COST OF SALES - PARTS" Or A = "REPAIR ORDER ON PROCESS" Or A = "SERVICE COST OF SALES - ACCESSORIES" Or A = "SERVICE COST OF SALES - MATERIALS" Then
        chkSummary.Visible = False: picCSHCHG.Visible = False
    Else
        chkSummary.Visible = True: picCSHCHG.Visible = False
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "SERVICE REPORT") = False Then Exit Sub

    On Error GoTo ErrorCode

    Dim Filter                                         As String

    rptReports.Reset
    rptReports.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptReports.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptReports.Formulas(2) = "printedby = '" & LOGNAME & "'"

    Screen.MousePointer = 11
    Select Case UCase(LTrim(RTrim(cboReportType.Text)))
        Case "BILLED OUT REPORT"
            If chkSummary.Value = 1 Then
                rptReports.ReportTitle = "BILLED OUT REPORT-SUMMARY"
                rptReports.WindowTitle = "BILLED OUT REPORT-SUMMARY"
                'PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUTSUM.RPT", "{REPOR.DTE_COMP} >= DATE(" & Year(txtFrom.Value) & "," & Month(txtFrom.Value) & "," & Day(txtFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTo.Value) & "," & Month(txtTo.Value) & "," & Day(txtTo.Value) & ") AND isnull({REPOR.DTE_REL}) = false ", CSMS_REPORT_CONNECTION, 1
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUTSUM.RPT", "{REPOR.DTE_COMP} >= DATE(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                'PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUTSUM.RPT", "MONTH({REPOR.DTE_COMP}) = 9 AND YEAR({REPOR.DTE_COMP}) = 2008 AND DAY({REPOR.DTE_COMP}) = 2", CSMS_REPORT_CONNECTION, 1
                'PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUTSUM.RPT", "{REPOR.REP_OR} = 'R-00001147'", CSMS_REPORT_CONNECTION, 1

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "BILLED OUT REPORT SUMMARY " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            Else
                If COMPANY_CODE <> "CMC" And COMPANY_CODE <> "DSSC" Then
                    rptReports.Formulas(3) = "RANGEDATE = '" & "FROM " & TXTFrom.Value & " TO " & txtTO.Value & "'"
                    rptReports.ReportTitle = "BILLED OUT REPORT"
                    rptReports.WindowTitle = "BILLED OUT REPORT"
                    'PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUT.RPT", "{REPOR.DTE_COMP} >= DATE(" & Year(txtFrom.Value) & "," & Month(txtFrom.Value) & "," & Day(txtFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTo.Value) & "," & Month(txtTo.Value) & "," & Day(txtTo.Value) & ") AND isnull({REPOR.DTE_REL}) = false ", CSMS_REPORT_CONNECTION, 1
                    PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUT.RPT", "{REPOR.DTE_COMP} >= DATE(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                    'PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUT.RPT", "{REPOR.REP_OR} = 'R-00001147'", CSMS_REPORT_CONNECTION, 1
               
                Else
                    PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUT_ServiceSales.RPT", "{REPOR.DTE_COMP} >= DATE(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                    PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDOUT.RPT", "{REPOR.DTE_COMP} >= DATE(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                    'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "BILLED OUT REPORT DETAIL " & TXTFrom.Value & "-" & txtTO, "", "")
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
            End If

        Case "DAILY RELEASE REPORT"
            If chkSummary.Value = 1 Then
                rptReports.ReportTitle = "DAILY RELEASE REPORT-SUMMARY"
                rptReports.WindowTitle = "DAILY RELEASE REPORT-SUMMARY"
                'PrintSQLReport rptReports, CSMS_REPORT_PATH & "DAILYSALESSUM.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(txtFrom.Value) & "," & Month(txtFrom.Value) & "," & Day(txtFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTo.Value) & "," & Month(txtTo.Value) & "," & Day(txtTo.Value) & ")", CSMS_REPORT_CONNECTION, 1
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "DAILYSALESSUM.RPT", "{REPOR.DTE_REL} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_REL} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "DAILY RELEASE REPORT - SUMMARY " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            Else
                If COMPANY_CODE <> "CMC" Then
                    rptReports.ReportTitle = "DAILY RELEASE REPORT"
                    rptReports.WindowTitle = "DAILY RELEASE REPORT"
                    'PrintSQLReport rptReports, CSMS_REPORT_PATH & "DAILYSALES.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(txtFrom.Value) & "," & Month(txtFrom.Value) & "," & Day(txtFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTo.Value) & "," & Month(txtTo.Value) & "," & Day(txtTo.Value) & ")", CSMS_REPORT_CONNECTION, 1
                    PrintSQLReport rptReports, CSMS_REPORT_PATH & "DAILYSALES.RPT", "{REPOR.DTE_REL} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_REL} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                
                Else
                    rptReports.ReportTitle = "DAILY RELEASE REPORT"
                    rptReports.WindowTitle = "DAILY RELEASE REPORT"
                    'PrintSQLReport rptReports, CSMS_REPORT_PATH & "DAILYSALES.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(txtFrom.Value) & "," & Month(txtFrom.Value) & "," & Day(txtFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTo.Value) & "," & Month(txtTo.Value) & "," & Day(txtTo.Value) & ")", CSMS_REPORT_CONNECTION, 1
                    PrintSQLReport rptReports, CSMS_REPORT_PATH & "DAILYSALES.RPT", "{REPOR.DTE_REL} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_REL} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                    
                    Dim RSRODET As ADODB.Recordset
                    Dim rsCusmas As ADODB.Recordset
                    
                    Set RSRODET = New ADODB.Recordset
                    Set RSRODET = gconDMIS.Execute("select * from Csms_repor inner join ALL_CUSTOMER ON ALL_CUSTOMER.CUSCDE=CSMS_Repor.ACCT_NO  WHERE DTE_REL >= '" & TXTFrom.Value & "' AND DTE_REL <= '" & txtTO.Value & "'")
                   
                   ' RSRODET.Open "select * from Csms_repor WHERE {REPOR.DTE_REL} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_REL} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", gconDMIS, adOpenForwardOnly, adLockReadOnly
                    
                    Dim xlApp
                    Dim xlBook
                    Dim xlSheet1
                    Set xlApp = CreateObject("Excel.Application")
                
                    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "SALES TRANSACTION.xlt")
                    Set xlSheet1 = xlBook.Worksheets(1)
                
                    Dim i                               As Integer
                    Dim j                               As Integer
                    
                    xlSheet1.Cells(5, 1) = "TIN : " & COMPANY_TIN
                    xlSheet1.Cells(6, 1) = "OWNER'S NAME : " & COMPANY_NAME
                    xlSheet1.Cells(7, 1) = "OWNER'S TRADE NAME : " & COMPANY_NAME
                    xlSheet1.Cells(8, 1) = "OWNER'S ADDRESS : " & COMPANY_ADDRESS
                     
                     If Not RSRODET.EOF And Not RSRODET.EOF Then
                        Do While Not RSRODET.EOF
'                        Set rscusmas = New ADODB.Recordset
'                        Set rscusmas = gconDMIS.Execute("select COMPANYADD,ACCTNAME,TIN,CUSTYPE from ALL_Customer_Table INNER JOIN ALL_CUSTOMER ON ALL_CUSTOMER.CUSCDE=CSMS_Repor.ACCT_NO")
'                                xlSheet1.Cells(11 + j, 1) = j + 1
                                xlSheet1.Cells(11 + j, 1) = Null2String(RSRODET!dte_rel)
                                xlSheet1.Cells(11 + j, 2) = Null2String(RSRODET!TIN)
                                If RSRODET!CUSTYPE = "P" Then
                                xlSheet1.Cells(11 + j, 4) = Null2String(RSRODET!AcctName)
                                Else
                                xlSheet1.Cells(11 + j, 3) = Null2String(RSRODET!AcctName)
                                End If
                                xlSheet1.Cells(11 + j, 5) = Null2String(RSRODET!CUSTOMERADD)
                                xlSheet1.Cells(11 + j, 6) = Null2String(RSRODET!amount / 1.12)
                                xlSheet1.Cells(11 + j, 7) = 0#
                                xlSheet1.Cells(11 + j, 8) = 0#
                                xlSheet1.Cells(11 + j, 9) = Null2String(RSRODET!amount / 1.12)
                                xlSheet1.Cells(11 + j, 10) = Null2String((RSRODET!amount / 1.12) * 0.12)
                                xlSheet1.Cells(11 + j, 11) = Null2String(RSRODET!amount)
                                j = j + 1
                            
                            RSRODET.MoveNext
                        Loop
                     End If
                    xlApp.Visible = True
                    Set xlBook = Nothing
                    Set xlSheet1 = Nothing
                    Set xlApp = Nothing
                             
                 End If

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "DAILY RELEASE REPORT - DETAIL " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            End If

        Case "BILLED OUT BUT UNRELEASED"
            If chkSummary.Value = 1 Then
                rptReports.ReportTitle = "BILLED OUT BUT UNRELEASED-SUMMARY"
                rptReports.WindowTitle = "BILLED OUT BUT UNRELEASED-SUMMARY"
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDUNRELEASEDSUM.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "BILLED OUT BUT UNRELEASED - SUMMARY " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            Else
                rptReports.ReportTitle = "BILLED OUT BUT UNRELEASED"
                rptReports.WindowTitle = "BILLED OUT BUT UNRELEASED"
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "BILLEDUNRELEASED.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "BILLED OUT BUT UNRELEASED - DETAIL " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            End If

        Case "SERVICE INVOICE SUMMARY REPORT"
            If optCash.Value = True Then
                rptReports.ReportTitle = "SERVICE INVOICE SUMMARY REPORT - CASH"
                rptReports.WindowTitle = "SERVICE INVOICE SUMMARY REPORT - CASH"
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "ACCTSR.RPT", "{REPOR.TERM} = 'CSH' AND {REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATESERIAL(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1    ' (({REPOR.TERM} = 'CSH' or isnull({repor.term})=true)) AND

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "SERVICE INVOICE SUMMARY REPORT - CASH " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            ElseIf optINT.Value = True Then
                rptReports.ReportTitle = "SERVICE INVOICE SUMMARY REPORT - INTERNAL"
                rptReports.WindowTitle = "SERVICE INVOICE SUMMARY REPORT - INTERNAL"
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "ACCTSR.RPT", "ISNULL({REPOR.TERM}) = TRuE AND {REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATESERIAL(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1    ' (({REPOR.TERM} = 'CSH' or isnull({repor.term})=true)) AND

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "SERVICE INVOICE SUMMARY REPORT - INTERNAL " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            ElseIf optCharge.Value = True Then
                rptReports.ReportTitle = "SERVICE INVOICE SUMMARY REPORT - CHARGE"
                rptReports.WindowTitle = "SERVICE INVOICE SUMMARY REPORT - CHARGE"
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "ACCTSR.RPT", "{REPOR.TERM} = 'CHG' AND {REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATESERIAL(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "SERVICE INVOICE SUMMARY REPORT - CHARGE " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            Else
                rptReports.ReportTitle = "SERVICE INVOICE SUMMARY REPORT - ALL"
                rptReports.WindowTitle = "SERVICE INVOICE SUMMARY REPORT - ALL"
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "ACCTSR.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATESERIAL(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "SERVICE INVOICE SUMMARY REPORT - ALL " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            End If

        Case "REPAIR ORDER ON PROGRESS"
            If chkSummary.Value = 1 Then
                rptReports.ReportTitle = "REPAIR ORDER - ON-PROGRESS SUMMARY"
                rptReports.WindowTitle = "REPAIR ORDER - ON-PROGRESS SUMMARY"
                'PrintSQLReport rptReports, CSMS_REPORT_PATH & "RO_ONPROCESS.RPT", "{REPOR.DTE_RECD} >= DATESERIAL(" & Year(txtfrom.Text) & "," & Month(txtfrom.Text) & "," & Day(txtfrom.Text) & ") AND {REPOR.DTE_RECD} <= DATE(" & Year(txtto.Text) & "," & Month(txtto.Text) & "," & Day(txtto.Text) & ") AND ISNULL({REPOR.DTE_COMP}) = TRUE AND ISNULL({REPOR.DTE_REL}) = TRUE ", CSMS_REPORT_CONNECTION, 1 'comment by JUN 01/22/2008
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "RO_ONPROCESSSUM.RPT", "{REPOR.DTE_RECD} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_RECD} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1    'JUN 01/22/2008

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "REPAIR ORDER WORK IN PROGRESS - SUMMARY " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            Else
                rptReports.ReportTitle = "REPAIR ORDER - ON-PROGRESS"
                rptReports.WindowTitle = "REPAIR ORDER - ON-PROGRESS"
                'PrintSQLReport rptReports, CSMS_REPORT_PATH & "RO_ONPROCESS.RPT", "{REPOR.DTE_RECD} >= DATESERIAL(" & Year(txtfrom.Text) & "," & Month(txtfrom.Text) & "," & Day(txtfrom.Text) & ") AND {REPOR.DTE_RECD} <= DATE(" & Year(txtto.Text) & "," & Month(txtto.Text) & "," & Day(txtto.Text) & ") AND ISNULL({REPOR.DTE_COMP}) = TRUE AND ISNULL({REPOR.DTE_REL}) = TRUE ", CSMS_REPORT_CONNECTION, 1 'comment by JUN 01/22/2008
                PrintSQLReport rptReports, CSMS_REPORT_PATH & "RO_ONPROCESS.RPT", "{REPOR.DTE_RECD} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_RECD} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1    'JUN 01/22/2008

                'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "REPAIR ORDER WORK IN PROGRESS -DETAIL " & TXTFrom.Value & "-" & txtTO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            End If
        Case "SERVICE COST OF SALES - PARTS"
            rptReports.ReportTitle = "SERVICE COST OF SALES - PARTS"
            rptReports.WindowTitle = "SERVICE COST OF SALES - PARTS"
            PrintSQLReport rptReports, CSMS_REPORT_PATH & "SI_COSTOFSALES.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "SERVICE COST OF SALES - PARTS " & TXTFrom.Value & "-" & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------

        Case "SERVICE COST OF SALES - ACCESSORIES"
            rptReports.ReportTitle = "SERVICE COST OF SALES - ACCESSORIES"
            rptReports.WindowTitle = "SERVICE COST OF SALES - ACCESSORIES"
            PrintSQLReport rptReports, CSMS_REPORT_PATH & "SI_COSTOFSALES_ACC.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "SERVICE COST OF SALES - ACCESSORIES " & TXTFrom.Value & "-" & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Case "SERVICE COST OF SALES - MATERIALS"
            rptReports.ReportTitle = "SERVICE COST OF SALES - MATERIALS"
            rptReports.WindowTitle = "SERVICE COST OF SALES - MATERIALS"
            PrintSQLReport rptReports, CSMS_REPORT_PATH & "SI_COSTOFSALES_MAT.RPT", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "SERVICE REPORT", "", "", "", "SERVICE COST OF SALES - MATERIALS " & TXTFrom.Value & "-" & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
    
    End Select

    Screen.MousePointer = 0

    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SERVICE REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SERVICE REPORT", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    TXTFrom.Value = firstDay(LOGDATE): txtTO.Value = LOGDATE: chkSummary.Value = 1

    cboReportType.Clear
    cboReportType.AddItem "Billed Out Report"
    cboReportType.AddItem "Daily Release Report"
    cboReportType.AddItem "Repair Order On Progress"
    cboReportType.AddItem "Billed Out But Unreleased"
    cboReportType.AddItem "Service Invoice Summary Report"
    cboReportType.AddItem "Service Cost of Sales - Parts"
    cboReportType.AddItem "Service Cost of Sales - Accessories"
    cboReportType.AddItem "Service Cost of Sales - Materials"

    cboReportType.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSServiceReport = Nothing
End Sub

Private Sub txtFrom_LostFocus()
    If TXTFrom.Value <> "" Then TXTFrom.Value = Format(TXTFrom.Value, "Short Date")
End Sub

Private Sub txtTo_LostFocus()
    If txtTO.Value <> "" Then txtTO.Value = Format(txtTO.Value, "Short Date")
End Sub

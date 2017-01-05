VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Report_Log 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Report"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   3915
   Icon            =   "Report_LogReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optProspectLog 
      Caption         =   "Prospect Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   3225
   End
   Begin VB.OptionButton optCustomerLog 
      Caption         =   "Customer Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Value           =   -1  'True
      Width           =   3225
   End
   Begin Crystal.CrystalReport rptLogs 
      Left            =   135
      Top             =   2190
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Height          =   825
      Left            =   1905
      MouseIcon       =   "Report_LogReport.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "Report_LogReport.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   1950
      Width           =   885
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
      Height          =   825
      Left            =   1035
      MouseIcon       =   "Report_LogReport.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "Report_LogReport.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   1950
      Width           =   885
   End
   Begin MSComCtl2.DTPicker dtpToDateLog 
      Height          =   435
      Left            =   1290
      TabIndex        =   4
      Top             =   1485
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   52494337
      CurrentDate     =   39232
   End
   Begin MSComCtl2.DTPicker dtpFromDateLog 
      Height          =   435
      Left            =   1275
      TabIndex        =   5
      Top             =   945
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   52494337
      CurrentDate     =   39203
   End
   Begin VB.Label Label1 
      Caption         =   "From Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   1035
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "To Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   390
      TabIndex        =   6
      Top             =   1575
      Width           =   885
   End
End
Attribute VB_Name = "frmCRIS_Report_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    Dim FDate                                                         As Date
    Dim TDate                                                         As Date
    Dim rsLogs                                                        As ADODB.Recordset
    Dim RecordSelection                                               As String
    Set rsLogs = New ADODB.Recordset

    FDate = CDate(dtpFromDateLog.Value)
    TDate = CDate(dtpToDateLog.Value)

    Set rsLogs = gconDMIS.Execute("SELECT * from CRIS_ViewLog")

    If Not rsLogs.BOF And Not rsLogs.EOF Then
        If optCustomerLog.Value = True Then
            Dim rsC_Log                                               As ADODB.Recordset
            Set rsC_Log = New ADODB.Recordset

            Set rsC_Log = gconDMIS.Execute("Select * from CRIS_ViewLog where cscde is not NULL")
            If Not rsC_Log.EOF And Not rsC_Log.BOF Then

                rptLogs.Formulas(0) = "datefrom = '" & FDate & "'"
                rptLogs.Formulas(1) = "dateto = '" & TDate & "'"

                rptLogs.Formulas(2) = "CompanyName = '" & COMPANY_NAME & "'"
                rptLogs.Formulas(3) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

                RecordSelection = "YEAR({CRIS_ViewLog.Deyt}) =" & Year(FDate)
                RecordSelection = "{CRIS_ViewLog.Deyt} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CRIS_ViewLog.Deyt} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")"
                PrintSQLReport rptLogs, CRIS_REPORT_PATH & "CustomerLogReport.rpt", RecordSelection, CRIS_REPORT_PATH, 1
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "CUSTOMER LOG REPORT", "", "", "", "CUSTOMER LOG SUMMARY - " & dtpFromDateLog & " " & dtpToDateLog, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

                'LogAudit "V", "CUSTOMER LOG SUMMARY ", "FROM " & FDate & " TO " & TDate
            Else
                ShowNoRecord
            End If
        Else
            Dim rsP_Log                                               As ADODB.Recordset
            Set rsP_Log = New ADODB.Recordset
            Set rsP_Log = gconDMIS.Execute("Select * from CRIS_ViewLog where cscde is NULL")
            If Not rsP_Log.EOF And Not rsP_Log.BOF Then
                rptLogs.Formulas(0) = "datefrom = '" & FDate & "'"
                rptLogs.Formulas(1) = "dateto = '" & TDate & "'"
                rptLogs.Formulas(2) = "CompanyName = '" & COMPANY_NAME & "'"
                rptLogs.Formulas(3) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

                PrintSQLReport rptLogs, CRIS_REPORT_PATH & "ProspectLogReport.rpt", "{CRIS_ViewLog.Deyt} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CRIS_ViewLog.Deyt} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", CRIS_REPORT_PATH, 1
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "CUSTOMER LOG REPORT", "", "", "", "PROSPECT LOG SUMMARY - " & dtpFromDateLog & " " & dtpToDateLog, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

                'LogAudit "V", "PROSPECT LOG SUMMARY ", "FROM " & FDate & " TO " & TDate
            End If
        End If
    Else
        ShowNoRecord
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CUSTOMER LOG REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CUSTOMER LOG REPORT", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    dtpFromDateLog.Value = firstDay(LOGDATE)
    dtpToDateLog.Value = LOGDATE
End Sub

Private Sub optCustomerLog_Click()
    Label1.Enabled = True
    Label2.Enabled = True
    dtpFromDateLog.Enabled = True
    dtpToDateLog.Enabled = True
    cmdPrint.Enabled = True
End Sub

Private Sub optProspectLog_Click()
    Label1.Enabled = True
    Label2.Enabled = True
    dtpFromDateLog.Enabled = True
    dtpToDateLog.Enabled = True
    cmdPrint.Enabled = True
End Sub


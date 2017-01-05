VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSWorkInProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Work In Progress Report"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSWorkInProgress.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   3720
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmCSMSWorkInProgress.frx":1082
      Left            =   1260
      List            =   "frmCSMSWorkInProgress.frx":108C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   900
      Width           =   2325
   End
   Begin Crystal.CrystalReport rptWork_In_Progress 
      Left            =   270
      Top             =   1590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Work In Progress Monitoring Report"
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpToDateWorkInProgress 
      Height          =   345
      Left            =   1260
      TabIndex        =   1
      Top             =   510
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   609
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
      Format          =   59572225
      CurrentDate     =   39232
   End
   Begin MSComCtl2.DTPicker dtpFromDateWorkInProgress 
      Height          =   345
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   609
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
      Format          =   59572225
      CurrentDate     =   39203
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
      Height          =   855
      Left            =   2820
      MouseIcon       =   "frmCSMSWorkInProgress.frx":109E
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSWorkInProgress.frx":11F0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   1320
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
      Height          =   855
      Left            =   2100
      MouseIcon       =   "frmCSMSWorkInProgress.frx":163B
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSWorkInProgress.frx":178D
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   570
      TabIndex        =   7
      Top             =   960
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   450
      TabIndex        =   6
      Top             =   570
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   180
      Width           =   870
   End
End
Attribute VB_Name = "frmCSMSWorkInProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "WORKING IN PROGRRESS") = False Then Exit Sub
    Screen.MousePointer = 11
    'On Error GoTo Errorcode

    Dim rsJStatus                                      As ADODB.Recordset
    Dim FDate                                          As Date
    Dim TDate                                          As Date

    FDate = CDate(dtpFromDateWorkInProgress.Value)
    TDate = CDate(dtpToDateWorkInProgress.Value)
    rptWork_In_Progress.Formulas(2) = "FromDate = '" & FDate & "'"
    rptWork_In_Progress.Formulas(3) = "ToDate = '" & TDate & "'"
    'PrintSQLReport rptWork_In_Progress, CSMS_REPORT_PATH & "Work_In_Progress_Monitoring.rpt", "{CSMS_Repor.Dte_Recd} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_Repor.Dte_Recd} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ") and {CSMS_Repairorder.status}<>'Billed' and {CSMS_Repairorder.status}<>'Finish Job' and {CSMS_Repairorder.status}<>'Released' and {CSMS_Repairorder.status}<>'Park' and {CSMS_RO_DET.livil}='1' ", CSMS_REPORT_CONNECTION, 1 'comment by jun 01/22/2008

    'JUN 02/05/2008
    rptWork_In_Progress.Formulas(4) = "Company Name = '" & COMPANY_NAME & "'"
    rptWork_In_Progress.Formulas(5) = "Company Address = '" & COMPANY_ADDRESS & "'"
    rptWork_In_Progress.Formulas(6) = "Printedby = '" & LOGNAME & "'"

    'PrintSQLReport rptWork_In_Progress, CSMS_REPORT_PATH & "Work_In_Progress_Monitoring.rpt", "{CSMS_REPOR.DTE_RECD} >= Date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_REPOR.DTE_RECD} <= Date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ") AND {CSMS_Repairorder.status}<>'Billed' AND {CSMS_Repairorder.status}<>'Finish Job' AND {CSMS_Repairorder.status}<>'Released' AND {CSMS_Repairorder.status}<>'Park' AND ISNULL({CSMS_REPOR.DTE_COMP}) = TRUE AND ISNULL({CSMS_REPOR.DTE_REL}) = TRUE ", CSMS_REPORT_CONNECTION, 1 ' jun 01/22/2008
    If Combo1.Text = "Park" Then
        PrintSQLReport rptWork_In_Progress, CSMS_REPORT_PATH & "Work_In_Progress_Monitoring.rpt", "{CSMS_REPOR.DTE_RECD} >= Date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_REPOR.DTE_RECD} <= Date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ") AND {CSMS_Repairorder.status} = 'Park' AND ISNULL({CSMS_REPOR.DTE_COMP}) = TRUE AND ISNULL({CSMS_REPOR.DTE_REL}) = TRUE ", CSMS_REPORT_CONNECTION, 1    ' jun 01/22/2008
    Else
        PrintSQLReport rptWork_In_Progress, CSMS_REPORT_PATH & "Work_In_Progress_Monitoring.rpt", "{CSMS_REPOR.DTE_RECD} >= Date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_REPOR.DTE_RECD} <= Date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ") AND {CSMS_Repairorder.status}<>'Billed' AND {CSMS_Repairorder.status}<>'Finish Job' AND {CSMS_Repairorder.status}<>'Released' AND ISNULL({CSMS_REPOR.DTE_COMP}) = TRUE AND ISNULL({CSMS_REPOR.DTE_REL}) = TRUE ", CSMS_REPORT_CONNECTION, 1    ' jun 01/22/2008
    End If
    'LogAudit "V", "WORKING IN PROGRRESS", dtpFromDateWorkInProgress & "-" & dtpToDateWorkInProgress

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "WORKING IN PROGRESS", "", "", "", dtpFromDateWorkInProgress & " TO " & dtpToDateWorkInProgress & " - " & Combo1, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (WORKING IN PROGRESS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "WORKING IN PROGRESS", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    dtpFromDateWorkInProgress.Value = firstDay(LOGDATE)
    dtpToDateWorkInProgress.Value = LOGDATE
    Combo1.ListIndex = 0
End Sub


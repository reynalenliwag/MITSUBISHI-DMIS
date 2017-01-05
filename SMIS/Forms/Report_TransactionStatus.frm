VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_TransactionStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Status Report"
   ClientHeight    =   4185
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_TransactionStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   3930
   Begin VB.Frame Frame2 
      Caption         =   "Select Transaction Status"
      ForeColor       =   &H00800000&
      Height          =   1395
      Left            =   113
      TabIndex        =   7
      Top             =   1710
      Width           =   3675
      Begin VB.OptionButton chkALL 
         Caption         =   "All"
         Height          =   285
         Left            =   2370
         TabIndex        =   11
         Top             =   540
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.OptionButton chkCancelled 
         Caption         =   "Cancelled"
         Height          =   285
         Left            =   330
         TabIndex        =   10
         Top             =   870
         Width           =   2715
      End
      Begin VB.OptionButton chkUnPosted 
         Caption         =   "Un Posted"
         Height          =   285
         Left            =   330
         TabIndex        =   9
         Top             =   570
         Value           =   -1  'True
         Width           =   2715
      End
      Begin VB.OptionButton chkPosted 
         Caption         =   "Posted"
         Height          =   285
         Left            =   330
         TabIndex        =   8
         Top             =   270
         Width           =   2715
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please Select Your Transaction Type"
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   113
      TabIndex        =   2
      Top             =   120
      Width           =   3705
      Begin VB.OptionButton optRR 
         Caption         =   "Vehicle Recieving Entry"
         Height          =   225
         Left            =   330
         TabIndex        =   6
         Top             =   1140
         Width           =   2625
      End
      Begin VB.OptionButton optPO 
         Caption         =   "Vehicle Purchase Order"
         Height          =   225
         Left            =   330
         TabIndex        =   5
         Top             =   840
         Width           =   2625
      End
      Begin VB.OptionButton optVI 
         Caption         =   "Vehicle Sales Invoice"
         Height          =   225
         Left            =   330
         TabIndex        =   4
         Top             =   540
         Width           =   2625
      End
      Begin VB.OptionButton optSO 
         Caption         =   "Sales Order"
         Height          =   225
         Left            =   330
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2625
      End
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
      Height          =   825
      Left            =   1890
      MouseIcon       =   "Report_TransactionStatus.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Report_TransactionStatus.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   3210
      Width           =   885
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
      Height          =   825
      Left            =   1020
      MouseIcon       =   "Report_TransactionStatus.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "Report_TransactionStatus.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   3210
      Width           =   885
   End
   Begin Crystal.CrystalReport rptInvDate 
      Left            =   2580
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Vehicle Inventory"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmSMIS_Report_TransactionStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    Dim REPORTNAME, FILTER                                            As String
    Dim Audit_Value1                                                  As String
    
    Screen.MousePointer = 11
    rptInvDate.Reset
    rptInvDate.SelectionFormula = ""
    FILTER = "": REPORTNAME = ""
    rptInvDate.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInvDate.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    If optSO.Value = True Then
        REPORTNAME = "TransactionStatus-SO.rpt"
        rptInvDate.ReportTitle = "TRANSACTION STATUS REPORT SALES ORDER (VSO)"

        If chkCancelled.Value = True Then
            FILTER = "{DET.SOSTATUS}='C'"
        ElseIf chkPosted.Value = True Then
            FILTER = "({DET.SOSTATUS}='P')"
        ElseIf chkUnPosted.Value = True Then
            'FILTER = "({DET.SOSTATUS})='U' AND ISNULL(({DET.SOSTATUS}))=FALSE"
            'FILTER = "({DET.SOSTATUS}='' AND ISNULL({DET.SOSTATUS}) = False AND {DET.SOSTATUS}='U')"
            FILTER = "ISNULL({DET.SOSTATUS})=FALSE AND ({DET.SOSTATUS})='U'"
        End If


    ElseIf optVI.Value = True Then
        REPORTNAME = "TransactionStatus-VI.rpt"
        rptInvDate.ReportTitle = "TRANSACTION STATUS REPORT VEHICLE INVOICE (VI)"
        If chkCancelled.Value = True Then
            FILTER = "{DET.STATUS}='C'"
        ElseIf chkPosted.Value = True Then
            FILTER = "{DET.STATUS}='P'"
        ElseIf chkUnPosted.Value = True Then
            FILTER = "({DET.STATUS}='' OR ISNULL({DET.STATUS}) = TRUE or {DET.STATUS}='U' )"
        End If
    ElseIf optPO.Value = True Then
        REPORTNAME = "TransactionStatus-PO.rpt"
        rptInvDate.ReportTitle = "TRANSACTION STATUS REPORT VEHICLE PURCHASE ORDER(VPO)"
        If chkCancelled.Value = True Then
            FILTER = "{DET.STATUS}='C'"
        ElseIf chkPosted.Value = True Then
            FILTER = "{DET.STATUS}='P'"
        ElseIf chkUnPosted.Value = True Then
            FILTER = "({DET.STATUS}='' OR ISNULL({DET.STATUS}) = TRUE or {DET.STATUS}='U' )"
        End If
    ElseIf optRR.Value = True Then
        REPORTNAME = "TransactionStatus-MRR.rpt"
        rptInvDate.ReportTitle = "TRANSACTION STATUS REPORT RECIEVING (MRR)"
        If chkCancelled.Value = True Then
            FILTER = "{DET.STATUS}='C'"
        ElseIf chkPosted.Value = True Then
            FILTER = "{DET.STATUS}='P'"
        ElseIf chkUnPosted.Value = True Then
            FILTER = "({DET.STATUS}='' OR ISNULL({DET.STATUS}) = TRUE or {DET.STATUS}='U' )"
        End If
    End If
    PrintSQLReport rptInvDate, SMIS_REPORT_PATH & "LISTING/" & REPORTNAME, FILTER, DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0

    'UPDATED BY: JUN
    'DATE UPDATED: 09032008 5:00
    'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
     If chkCancelled.Value = True Then
        Audit_Value1 = "CANCELLED"
     ElseIf chkPosted.Value = True Then
        Audit_Value1 = "POSTED"
     Else
        Audit_Value1 = "UNPOSTED"
     End If

     Call NEW_LogAudit("V", "TRANSACTION STATUS REPORT", "", "", "", rptInvDate.ReportTitle & "-" & Audit_Value1, "", "")
    'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'LogAudit "V", rptInvDate.ReportTitle
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTION STATUS REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "TRANSACTION STATUS REPORT", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
End Sub


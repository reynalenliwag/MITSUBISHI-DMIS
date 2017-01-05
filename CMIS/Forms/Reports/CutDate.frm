VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCMISCutDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Tally Report"
   ClientHeight    =   1530
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   3300
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CutDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   3300
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
      Left            =   2355
      MouseIcon       =   "CutDate.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "CutDate.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   630
      Width           =   765
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
      Left            =   1605
      MouseIcon       =   "CutDate.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "CutDate.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   630
      Width           =   765
   End
   Begin Crystal.CrystalReport rptCMISReportRange 
      Left            =   180
      Top             =   810
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpCutDate 
      Height          =   405
      Left            =   1440
      TabIndex        =   0
      Top             =   90
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   49610753
      CurrentDate     =   38216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Cut-Off Date"
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
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Width           =   1245
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   3
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmCMISCutDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "REPORT CASH TALLY REPORT") = False Then
        'rptCMISReportRange.PrintReport = 0
        Exit Sub
    End If

    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    Screen.MousePointer = 11
    With rptCMISReportRange
        .Formulas(0) = "DEALER_NAME = '" & COMPANY_NAME & "'"
        .Formulas(1) = "DEALER_ADDRESS = '" & COMPANY_ADDRESS & "'"
        .Formulas(2) = "PREPAREDBY='" & PreparedBy & "'"
        .Formulas(3) = "NOTEDBY='" & NotedBy & "'"
        .Formulas(4) = "CHECKEDBY='" & CheckedBy & "'"
        .Formulas(5) = "PRINTEDBY=" & N2Str2Null(LOGNAME)
    End With

    PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Cash_Tally_Sheet_Report.rpt", "{CASH_POS.CUTDATE} = Date(" & Year(dtpCutDate) & "," & Month(dtpCutDate) & "," & Day(dtpCutDate) & ")", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0

    'NEW LOG AUDIT----------------------------------------------------------
        Call NEW_LogAudit("V", "REPORT CASH TALLY REPORT", "", "", "", "CUT OFF DATE: " & dtpCutDate.Value, "", "")
    'NEW LOG AUDIT----------------------------------------------------------
    Exit Sub

ErrorCode:
    ShowVBError
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CASH TALLY REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CASH TALLY REPORT", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpCutDate = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCMISCutDate = Nothing
End Sub

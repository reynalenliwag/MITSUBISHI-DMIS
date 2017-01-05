VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISInquiry_UnPostedRCRange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unposted Receipts"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   ForeColor       =   &H00DEDFDE&
   Icon            =   "UnPostedRCRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3045
   Begin Crystal.CrystalReport rptReceipts 
      Left            =   90
      Top             =   1260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Transaction Listings - Receipts"
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
   Begin MSMask.MaskEdBox txtFrom 
      Height          =   345
      Left            =   810
      TabIndex        =   0
      ToolTipText     =   "Input starting date of the report"
      Top             =   90
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   609
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTo 
      Height          =   345
      Left            =   810
      TabIndex        =   1
      ToolTipText     =   "Input end date of the report"
      Top             =   480
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   609
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      CausesValidation=   0   'False
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
      Left            =   1620
      MouseIcon       =   "UnPostedRCRange.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "UnPostedRCRange.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   960
      Width           =   735
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
      Left            =   900
      MouseIcon       =   "UnPostedRCRange.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "UnPostedRCRange.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Report"
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
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
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Left            =   0
      TabIndex        =   3
      Top             =   510
      Width           =   765
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   2
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISInquiry_UnPostedRCRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRR_HD                                            As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "REPORTS INTERNAL UNPOSTED RECEIPTS TRANSACTION") = False Then Exit Sub

    On Error GoTo Errorcode:

    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select * from PMIS_RR_Hd where (rrdate >= '" & txtFrom.Text & "' AND rrdate <= '" & txtTo.Text & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
        Screen.MousePointer = 11
        rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "Unpostedrcrange.rpt", "{rr_hd.rrdate} >= date(" & Year(txtFrom.Text) & "," & Month(txtFrom.Text) & "," & Day(txtFrom.Text) & ") AND {rr_hd.rrdate} <= date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
        cmdPrint.Enabled = False
        LogAudit "V", "UNPOSTED RECEIPTS", txtFrom & "-" & txtTo
    Else
        ShowNoRecord
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtFrom.Text = Format(firstDay(LOGDATE), "DD-MMM-YY")
    txtTo.Text = Format(LOGDATE, "DD-MMM-YY")
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_RCRange = Nothing
    UnloadForm Me
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtFrom_Validate(Cancel As Boolean)
    Cancel = Not IsDate(txtFrom)
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtTo_Validate(Cancel As Boolean)
    Cancel = Not IsDate(txtTo)
End Sub

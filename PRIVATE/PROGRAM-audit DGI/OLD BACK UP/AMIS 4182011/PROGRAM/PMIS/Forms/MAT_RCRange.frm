VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMISReports_RCRange_MAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Receipts"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_RCRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   3165
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   345
      Left            =   810
      TabIndex        =   1
      Top             =   480
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   50331649
      CurrentDate     =   39232
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   345
      Left            =   810
      TabIndex        =   0
      Top             =   90
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   50331649
      CurrentDate     =   39203
   End
   Begin VB.CheckBox chkHistReceipts 
      Caption         =   "Look in History File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   810
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin Crystal.CrystalReport rptReceipts 
      Left            =   90
      Top             =   1500
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
      TabIndex        =   5
      ToolTipText     =   "Input starting date of the report"
      Top             =   90
      Visible         =   0   'False
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
      TabIndex        =   6
      ToolTipText     =   "Input end date of the report"
      Top             =   480
      Visible         =   0   'False
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
      Left            =   1770
      MouseIcon       =   "MAT_RCRange.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "MAT_RCRange.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   1170
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
      Left            =   1050
      MouseIcon       =   "MAT_RCRange.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "MAT_RCRange.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1170
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   510
      Width           =   765
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   7
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISReports_RCRange_MAT"
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
    On Error GoTo Errorcode:

    Dim FDate                                          As Date
    Dim TDate                                          As Date

    FDate = CDate(dtpFromDate.Value)
    TDate = CDate(dtpToDate.Value)

    If chkHistReceipts.Value = 1 Then
        Set rsRR_HD = New ADODB.Recordset
        rsRR_HD.Open "select * from PMIS_Rec_Hist where TYPE = 'M' AND (rrdate >= '" & FDate & "' AND rrdate <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
            Screen.MousePointer = 11
            rptReceipts.ReportTitle = "TRANSACTION LISTING - RECEIPTS (MATERIALS)"
            rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptReceipts.Formulas(12) = "mindate = '" & FDate & "'"
            rptReceipts.Formulas(11) = "maxdate = '" & TDate & "'"
            PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "rchistrange.rpt", "{rr_hd.TYPE} = 'M' AND {rr_hd.rrdate} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {rr_hd.rrdate} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            Call NEW_LogAudit("V", "MATERIALS TRANSACTION LISTING", "", "", "", dtpFromDate & " " & dtpToDate & " HISTORY", "", "")
        Else
            ShowNoRecord
        End If
    Else
        Set rsRR_HD = New ADODB.Recordset
        rsRR_HD.Open "select * from PMIS_RR_Hd where TYPE = 'M' AND (rrdate >= '" & FDate & "' AND rrdate <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
            Screen.MousePointer = 11
            rptReceipts.ReportTitle = "TRANSACTION LISTING - RECEIPTS (MATERIALS)"
            rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptReceipts.Formulas(12) = "mindate = '" & FDate & "'"
            rptReceipts.Formulas(11) = "maxdate = '" & TDate & "'"
            PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "rcrange.rpt", "{rr_hd.TYPE} = 'M' AND {rr_hd.rrdate} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {rr_hd.rrdate} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            Call NEW_LogAudit("V", "MATERIALS TRANSACTION LISTING", "", "", "", dtpFromDate & " " & dtpToDate, "", "")
        Else
            ShowNoRecord
        End If
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
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    'txtFrom.Text = Format(firstDay(LOGDATE), "DD-MMM-YY")
    'txtTo.Text = Format(LOGDATE, "DD-MMM-YY")
    dtpFromDate.Value = firstDay(LOGDATE)
    dtpToDate.Value = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_RCRange = Nothing
    UnloadForm Me
End Sub

Private Sub txtFrom_GotFocus()
    txtFrom.Text = Format(txtFrom.Text, "Short Date")
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtTo_GotFocus()
    txtTo.Text = Format(txtTo.Text, "Short Date")
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub


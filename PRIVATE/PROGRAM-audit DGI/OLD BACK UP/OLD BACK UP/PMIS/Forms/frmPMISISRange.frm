VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMISReports_ISRange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Issuance"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   ForeColor       =   &H00DEDFDE&
   Icon            =   "frmPMISISRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   3135
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   345
      Left            =   810
      TabIndex        =   1
      Top             =   480
      Width           =   2250
      _ExtentX        =   3969
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
      Format          =   52822017
      CurrentDate     =   39232
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   345
      Left            =   810
      TabIndex        =   0
      Top             =   90
      Width           =   2250
      _ExtentX        =   3969
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
      Format          =   52822017
      CurrentDate     =   39203
   End
   Begin VB.TextBox txtTo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   810
      TabIndex        =   6
      Text            =   "Text1"
      ToolTipText     =   "Input the end date of the report "
      Top             =   480
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox txtFrom 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   810
      TabIndex        =   5
      Text            =   "Text1"
      ToolTipText     =   "Input starting date of the report"
      Top             =   90
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CheckBox chkHistIssuance 
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
      Top             =   870
      Width           =   2415
   End
   Begin Crystal.CrystalReport rptIssuance 
      Left            =   60
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Transaction Listing - Issuances"
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2370
      MouseIcon       =   "frmPMISISRange.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "frmPMISISRange.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   1170
      Width           =   675
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
      Left            =   1710
      MouseIcon       =   "frmPMISISRange.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "frmPMISISRange.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print this Record"
      Top             =   1170
      Width           =   675
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
Attribute VB_Name = "frmPMISReports_ISRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOrd_Hd                                                          As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    'If (txtFrom.Text > txtTo.Text) Or IsDate(txtFrom.Text) = False Or IsDate(txtTo.Text) = False Then
    '    MsgSpeechBox "Error In From and To date"
    '    Exit Sub
    'End If
    Dim FDate                                                         As Date
    Dim TDate                                                         As Date

    FDate = CDate(dtpFromDate.Value)
    TDate = CDate(dtpToDate.Value)

    If chkHistIssuance.Value = 1 Then
        Set rsOrd_Hd = New ADODB.Recordset
        rsOrd_Hd.Open "select trandate from PMIS_Ord_Hist where TYPE = 'P' AND (trandate >= '" & FDate & "' AND trandate <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsOrd_Hd.EOF And Not rsOrd_Hd.EOF Then
            Screen.MousePointer = 11
            rptIssuance.ReportTitle = "TRANSACTION LISTING - ISSUANCES (PARTS)"
            rptIssuance.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptIssuance.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptIssuance.Formulas(12) = "mindate = '" & FDate & "'"
            rptIssuance.Formulas(11) = "maxdate = '" & TDate & "'"
            PrintSQLReport rptIssuance, PMIS_REPORT_PATH & "ishistrange.rpt", "{ORD_hd.TYPE} = 'P' AND {ORD_hd.trandate} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_hd.trandate} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
            rptIssuance.ReportTitle = "TRANSACTION LISTING - ISSUANCES COST (PARTS)"
            PrintSQLReport rptIssuance, PMIS_REPORT_PATH & "ishistrangeCost.rpt", "{ORD_hd.TYPE} = 'P' AND {ORD_hd.trandate} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_hd.trandate} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
            NEW_LogAudit "V", "PARTS TRANSACTION LISTING", "", "", "", dtpFromDate & " - " & dtpToDate, "HISTORY", ""
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            Exit Sub
        End If
    Else
        Set rsOrd_Hd = New ADODB.Recordset
        rsOrd_Hd.Open "select trandate from PMIS_Ord_Hd where TYPE = 'P' AND (trandate >= '" & FDate & "' AND trandate <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsOrd_Hd.EOF And Not rsOrd_Hd.EOF Then
            Screen.MousePointer = 11
            rptIssuance.ReportTitle = "TRANSACTION LISTING - ISSUANCES (PARTS)"
            rptIssuance.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptIssuance.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptIssuance.Formulas(12) = "mindate = '" & FDate & "'"
            rptIssuance.Formulas(11) = "maxdate = '" & TDate & "'"
            PrintSQLReport rptIssuance, PMIS_REPORT_PATH & "isrange.rpt", "{ORD_hd.TYPE} = 'P' AND {ORD_hd.trandate} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_hd.trandate} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
         
            rptIssuance.ReportTitle = "TRANSACTION LISTING - ISSUANCES COST (PARTS)"
            PrintSQLReport rptIssuance, PMIS_REPORT_PATH & "israngeCost.rpt", "{ORD_hd.TYPE} = 'P' AND {ORD_hd.trandate} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_hd.trandate} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
            NEW_LogAudit "V", "PARTS TRANSACTION LISTING", "", "", "Parts", dtpFromDate & " - " & dtpToDate, "", ""
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            Exit Sub
        End If
    End If
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PARTS TRANSACTION LISTING)"
            Call frmALL_AuditInquiry.DisplayHistory("", "PARTS TRANSACTION LISTING", "PRINTING")
    End Select
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
    UnloadForm Me
End Sub

Private Sub txtFrom_GotFocus()
    txtFrom.Text = Format(txtFrom.Text, "Short Date")
End Sub

Private Sub txtFrom_LostFocus()
    txtFrom.Text = Format(txtFrom.Text, "DD-MMM-YY")
End Sub

Private Sub txtTo_GotFocus()
    txtTo.Text = Format(txtTo.Text, "Short Date")
End Sub

Private Sub txtTo_LostFocus()
    txtTo.Text = Format(txtTo.Text, "DD-MMM-YY")
End Sub


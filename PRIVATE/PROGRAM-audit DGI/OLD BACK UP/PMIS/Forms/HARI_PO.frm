VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMIS_HARI_PO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Report: HARI"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   ForeColor       =   &H00DEDFDE&
   Icon            =   "HARI_PO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3045
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
      Left            =   540
      TabIndex        =   2
      Top             =   900
      Width           =   2415
   End
   Begin Crystal.CrystalReport rptHariPO 
      Left            =   90
      Top             =   1620
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
      Left            =   1560
      MouseIcon       =   "HARI_PO.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "HARI_PO.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   1260
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
      Left            =   840
      MouseIcon       =   "HARI_PO.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "HARI_PO.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   1260
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   510
      Width           =   765
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
Attribute VB_Name = "frmPMIS_HARI_PO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTdayTran                                            As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "PURCHASE ORDER FROM HARI REPORTS") = False Then Exit Sub
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    If (txtFrom.Text > txtTo.Text) Or IsDate(txtFrom.Text) = False Or IsDate(txtTo.Text) = False Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If chkHistReceipts.Value = 1 Then
        Set rsTdayTran = New ADODB.Recordset
        rsTdayTran.Open "select * from PMIS_daytran where TYPE = 'P' AND  TRANTYPE = 'PO' and (TRANDATE >= '" & txtFrom.Text & "' AND TRANDATE <= '" & txtTo.Text & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTdayTran.EOF And Not rsTdayTran.EOF Then
            Screen.MousePointer = 11
            rptHariPO.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptHariPO.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptHariPO.Formulas(12) = "FromDate = '" & txtFrom.Text & "'"
            rptHariPO.Formulas(11) = "ToDate = '" & txtTo.Text & "'"
            PrintSQLReport rptHariPO, PMIS_REPORT_PATH & "OrderReport_HARI_Hist.rpt", "{PMIS_Tdaytran.TRANDATE} >= date(" & Year(txtFrom.Text) & "," & Month(txtFrom.Text) & "," & Day(txtFrom.Text) & ") AND {PMIS_Tdaytran.TRANDATE} <= date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            'cmdPrint.Enabled = False
            'LogAudit "V", "MONTHLY RECEIPT HISTORY"
        Else
            ShowNoRecord
        End If
    Else
        Set rsTdayTran = New ADODB.Recordset
        rsTdayTran.Open "select * from PMIS_Tdaytran where TYPE = 'P' AND TRANTYPE = 'PO' and (TRANDATE >= '" & txtFrom.Text & "' AND TRANDATE <= '" & txtTo.Text & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTdayTran.EOF And Not rsTdayTran.EOF Then
            Screen.MousePointer = 11
            rptHariPO.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptHariPO.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptHariPO.Formulas(12) = "FromDate = '" & txtFrom.Text & "'"
            rptHariPO.Formulas(11) = "ToDate = '" & txtTo.Text & "'"
            PrintSQLReport rptHariPO, PMIS_REPORT_PATH & "OrderReport_HARI.rpt", "{PMIS_Tdaytran.TRANDATE} >= date(" & Year(txtFrom.Text) & "," & Month(txtFrom.Text) & "," & Day(txtFrom.Text) & ") AND {PMIS_Tdaytran.TRANDATE} <= date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            'cmdPrint.Enabled = False
            'LogAudit "V", "MONTHLY RECEIPT HISTORY"
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
    txtFrom.Text = Format(firstDay(LOGDATE), "DD-MMM-YY")
    txtTo.Text = Format(LOGDATE, "DD-MMM-YY")
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISRCRange = Nothing
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

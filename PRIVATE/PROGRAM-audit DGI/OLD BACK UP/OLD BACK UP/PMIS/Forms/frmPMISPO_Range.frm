VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISPO_Range 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Purchase"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   ForeColor       =   &H00DEDFDE&
   Icon            =   "frmPMISPO_Range.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   2955
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
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Input the end date of the report "
      Top             =   480
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
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "Input starting date of the report"
      Top             =   90
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptPOHist 
      Left            =   30
      Top             =   1470
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Transactions Listing - Issuances"
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
      Left            =   1590
      MouseIcon       =   "frmPMISPO_Range.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "frmPMISPO_Range.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   900
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
      Left            =   870
      MouseIcon       =   "frmPMISPO_Range.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "frmPMISPO_Range.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Report"
      Top             =   900
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
      Top             =   3240
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISPO_Range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    
    On Error GoTo Errorcode
    If frmPMISMainMenu.optPO_AC = True Then
    If Function_Access(LOGID, "Acess_Print", "ACCESSORIES TRANSACTION LISTING") = False Then Exit Sub
            'Accessories Transaction History
            Dim rsPO_HD_Ac As ADODB.Recordset
            Set rsPO_HD_Ac = New ADODB.Recordset
            rptPOHist.Formulas(0) = "FromDate = '" & txtFrom.Text & "'"
            rptPOHist.Formulas(1) = "ToDate = '" & txtTo.Text & "'"
            rsPO_HD_Ac.Open "select trandate from PMIS_Tdaytran where (trandate >= '" & CDate(txtFrom.Text) & "' AND trandate <= '" & CDate(txtTo.Text) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
            
            If Not rsPO_HD_Ac.EOF And Not rsPO_HD_Ac.EOF Then
                Screen.MousePointer = 11
                rptPOHist.Formulas(12) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPOHist.Formulas(11) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPOHist, PMIS_REPORT_PATH & "PORange_Acc.rpt", "{Tdaytran.trandate} >= date(" & Year(txtFrom.Text) & "," & Month(txtFrom.Text) & "," & Day(txtFrom.Text) & ") AND {Tdaytran.trandate} <= date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            End If
    ElseIf frmPMISMainMenu.optPO_Mat = True Then
    If Function_Access(LOGID, "Acess_Print", "MATERIALS TRANSACTION LISTING") = False Then Exit Sub
            'Materials Transaction History
            Dim rsPO_HD_Mat As ADODB.Recordset
            Set rsPO_HD_Mat = New ADODB.Recordset
            rptPOHist.Formulas(0) = "FromDate = '" & txtFrom.Text & "'"
            rptPOHist.Formulas(1) = "ToDate = '" & txtTo.Text & "'"
            rsPO_HD_Mat.Open "select trandate from PMIS_Tdaytran where (trandate >= '" & CDate(txtFrom.Text) & "' AND trandate <= '" & CDate(txtTo.Text) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
            
            If Not rsPO_HD_Mat.EOF And Not rsPO_HD_Mat.EOF Then
                Screen.MousePointer = 11
                rptPOHist.Formulas(12) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPOHist.Formulas(11) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPOHist, PMIS_REPORT_PATH & "PORange_Mat.rpt", "{Tdaytran.trandate} >= date(" & Year(txtFrom.Text) & "," & Month(txtFrom.Text) & "," & Day(txtFrom.Text) & ") AND {Tdaytran.trandate} <= date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            End If
    Else
        If Function_Access(LOGID, "Acess_Print", "PARTS TRANSACTION LISTING") = False Then Exit Sub
            'Parts Transaction History
            Dim rsPO_HD_Parts As ADODB.Recordset
            Set rsPO_HD_Parts = New ADODB.Recordset
            rptPOHist.Formulas(0) = "FromDate = '" & txtFrom.Text & "'"
            rptPOHist.Formulas(1) = "ToDate = '" & txtTo.Text & "'"
            rsPO_HD_Parts.Open "select trandate from PMIS_Tdaytran where (trandate >= '" & CDate(txtFrom.Text) & "' AND trandate <= '" & CDate(txtTo.Text) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
            
            If Not rsPO_HD_Parts.EOF And Not rsPO_HD_Parts.EOF Then
                Screen.MousePointer = 11
                rptPOHist.Formulas(12) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPOHist.Formulas(11) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPOHist, PMIS_REPORT_PATH & "PORange_Parts.rpt", "{Tdaytran.trandate} >= date(" & Year(txtFrom.Text) & "," & Month(txtFrom.Text) & "," & Day(txtFrom.Text) & ") AND {Tdaytran.trandate} <= date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            End If
    End If
            'LogAudit "V", "UNPOSTED ISSUANCE", txtFrom & "-" & txtTo
    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
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
    UnloadForm Me
End Sub

Private Sub txtFrom_LostFocus()
    txtFrom.Text = Format(txtFrom.Text, "SHORT DATE")
End Sub

Private Sub txtTo_LostFocus()
    txtTo.Text = Format(txtTo.Text, "SHORT DATE")
End Sub

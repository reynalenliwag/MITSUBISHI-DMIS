VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISISRange1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Issuance"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   ForeColor       =   &H00DEDFDE&
   Icon            =   "ISRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3285
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
      Left            =   540
      TabIndex        =   2
      Top             =   870
      Width           =   2415
   End
   Begin Crystal.CrystalReport rptIssuance 
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
      Left            =   1680
      MouseIcon       =   "ISRange.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ISRange.frx":0F94
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
      Left            =   960
      MouseIcon       =   "ISRange.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ISRange.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print this Record"
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
Attribute VB_Name = "frmPMISISRange1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOrd_Hd                                           As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "PARTS TRANSACTION LISTING") = False Then Exit Sub
    
    On Error GoTo ErrorCode

    If (txtFrom.Text > txtTo.Text) Or IsDate(txtFrom.Text) = False Or IsDate(txtTo.Text) = False Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If chkHistIssuance.Value = 1 Then
        Set rsOrd_Hd = New ADODB.Recordset
        rsOrd_Hd.Open "select trandate from PMIS_Ord_Hist where TYPE = 'P' AND (trandate >= '" & CDate(txtFrom.Text) & "' AND trandate <= '" & CDate(txtTo.Text) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsOrd_Hd.EOF And Not rsOrd_Hd.EOF Then
            Screen.MousePointer = 11
            rptIssuance.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptIssuance.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptIssuance.Formulas(12) = "mindate = '" & txtFrom.Text & "'"
            rptIssuance.Formulas(11) = "maxdate = '" & txtTo.Text & "'"
            PrintSQLReport rptIssuance, PMIS_REPORT_PATH & "ishistrange.rpt", "{ORD_hd.TYPE} = 'P' AND {ORD_hd.trandate} >= date(" & Year(txtFrom.Text) & "," & Month(txtFrom.Text) & "," & Day(txtFrom.Text) & ") AND {ORD_hd.trandate} <= date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "MONTHLY ISSUANCE - History"
            Screen.MousePointer = 0
            'cmdPrint.Enabled = False
        Else
            ShowNoRecord
            Exit Sub
        End If
    Else
        Set rsOrd_Hd = New ADODB.Recordset
        rsOrd_Hd.Open "select trandate from PMIS_Ord_Hd where TYPE = 'P' AND (trandate >= '" & CDate(txtFrom.Text) & "' AND trandate <= '" & CDate(txtTo.Text) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsOrd_Hd.EOF And Not rsOrd_Hd.EOF Then
            Screen.MousePointer = 11
            rptIssuance.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptIssuance.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptIssuance.Formulas(12) = "mindate = '" & txtFrom.Text & "'"
            rptIssuance.Formulas(11) = "maxdate = '" & txtTo.Text & "'"
            PrintSQLReport rptIssuance, PMIS_REPORT_PATH & "isrange.rpt", "{ORD_hd.TYPE} = 'P' AND {ORD_hd.trandate} >= date(" & Year(txtFrom.Text) & "," & Month(txtFrom.Text) & "," & Day(txtFrom.Text) & ") AND {ORD_hd.trandate} <= date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "MONTHLY ISSUANCE - Range"
            Screen.MousePointer = 0
            'cmdPrint.Enabled = False
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

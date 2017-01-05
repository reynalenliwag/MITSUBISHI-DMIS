VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmSMIS_Report_YearlyGrossProfit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yearly Gross Profile"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_YearlyGrossProfit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1590
   ScaleWidth      =   3375
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
      Left            =   1785
      MouseIcon       =   "Report_YearlyGrossProfit.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Report_YearlyGrossProfit.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   675
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
      Left            =   795
      MouseIcon       =   "Report_YearlyGrossProfit.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "Report_YearlyGrossProfit.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   675
      Width           =   885
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   2250
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Yearly Ending Inventory"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   555
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "9999"
      Top             =   60
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   330
      TabIndex        =   1
      Top             =   90
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_YearlyGrossProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                           As ADODB.Recordset
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode

    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE YEAR(datereleased) >= " & txtYear, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11
        rptGenREP.Formulas(0) = "CompanyName = '" & Company_name & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
        rptGenREP.WindowTitle = "Yearly Gross Profit"
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "GPRREP3.rpt", "year({purchagree.datereleased}) = " & txtYear, DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Year of " & txtYear
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtYear.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

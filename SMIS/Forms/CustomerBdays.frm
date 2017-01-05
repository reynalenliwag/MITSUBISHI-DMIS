VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMISCustomerBDays 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Birthday Celebrants of the Month"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   ForeColor       =   &H00FCFCFC&
   Icon            =   "CustomerBdays.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1650
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
      Height          =   825
      Left            =   1665
      MouseIcon       =   "CustomerBdays.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "CustomerBdays.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   750
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
      Left            =   675
      MouseIcon       =   "CustomerBdays.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "CustomerBdays.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   750
      Width           =   885
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   465
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3105
   End
   Begin Crystal.CrystalReport rptCelebrants 
      Left            =   2670
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Customers Birthday Celebrants of the Month"
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
End
Attribute VB_Name = "frmSMISCustomerBDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                  As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_PurchAgree where Month(Birthdate) = " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
        rptCelebrants.Reset
        rptCelebrants.Formulas(0) = "YEER = " & Year(LOGDATE)
        rptCelebrants.Formulas(1) = "CURRENT_DAY = DATE(" & Year(LOGDATE) & "," & Month(LOGDATE) & "," & Day(LOGDATE) & ")"
        '   rptCelebrants.Formulas(2) = "CompanyName " & SetCompName("1")
        '  rptCelebrants.Formulas(3) = "CompanyAddress " & SetCompAdress("1")
        rptCelebrants.Formulas(0) = "CompanyName = '" & Company_name & "'"
        rptCelebrants.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
        PrintSQLReport rptCelebrants, SMIS_REPORT_PATH & "CustomerBday.rpt", "Month({Purchagree.BirthDate}) = " & What_month(cboMonth), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Celebrants for the month of " & cboMonth.Text
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    Screen.MousePointer = 0
End Sub

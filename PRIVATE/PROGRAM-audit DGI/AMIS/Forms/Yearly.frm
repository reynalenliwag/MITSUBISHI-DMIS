VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAMISYearly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a Year"
   ClientHeight    =   1500
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   2865
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Yearly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   2865
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
      Left            =   1515
      MouseIcon       =   "Yearly.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Yearly.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   585
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
      Left            =   645
      MouseIcon       =   "Yearly.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "Yearly.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   585
      Width           =   885
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   330
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptYearly 
      Left            =   30
      Top             =   570
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   -90
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   4
      Top             =   2940
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISYearly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsProfile                                               As ADODB.Recordset

Sub FillcboYear2()
    FillcboNewYear cboYear
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:33
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:



    On Error GoTo ErrorCode
    rptYearly.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptYearly.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptYearly.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
    End If
    rptYearly.WindowTitle = "SCHEDULE OF INCOME TAXES W/HELD FROM SUPPLIER FOR THE YEAR " & cboYear.Text
    rptYearly.ReportTitle = "SCHEDULE OF INCOME TAXES W/HELD FROM SUPPLIER FOR THE YEAR " & cboYear.Text
    PrintSQLReport rptYearly, AMIS_REPORT_PATH & "Schedules\SchedIncomeTaxesWheldFromSuppliers.Rpt", "year({Journal_Hd.InvoiceDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    LogAudit "V", "SCHEDULES OF INCOME TAX W/HELD FROM SUPPLIERS", cboYear
    Exit Sub

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillcboYear2
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub


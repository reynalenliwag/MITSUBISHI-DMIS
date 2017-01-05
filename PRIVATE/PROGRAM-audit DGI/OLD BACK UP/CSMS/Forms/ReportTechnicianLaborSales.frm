VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSTechnicianLaborSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Labor Sales"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReportTechnicianLaborSales.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   3825
   Begin VB.ComboBox cboTech 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "ReportTechnicianLaborSales.frx":0E42
      Left            =   1290
      List            =   "ReportTechnicianLaborSales.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   2475
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   540
      Width           =   2475
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   960
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   810
      Left            =   3015
      MouseIcon       =   "ReportTechnicianLaborSales.frx":0E46
      MousePointer    =   99  'Custom
      Picture         =   "ReportTechnicianLaborSales.frx":0F98
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   1365
      Width           =   735
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   30
      Top             =   1620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Technician Attendance Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   810
      Left            =   2295
      MouseIcon       =   "ReportTechnicianLaborSales.frx":13E3
      MousePointer    =   99  'Custom
      Picture         =   "ReportTechnicianLaborSales.frx":1535
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   1365
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Technician"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   7
      Top             =   210
      Width           =   930
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   645
      TabIndex        =   5
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   780
      TabIndex        =   4
      Top             =   1020
      Width           =   405
   End
End
Attribute VB_Name = "frmCSMSTechnicianLaborSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub FillCombo()
    Dim rstmp                                          As New ADODB.Recordset
    Dim NEYM                                           As String

    cboTECH.AddItem "All"
    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_vw_Technician order by tech_name")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            'NEYM = Null2String(RSTMP!lastname) & ", " & Null2String(RSTMP!Firstname) & " " & Left(Null2String(RSTMP!MIDDLENAME), 1) & "."
            'NEYM = Null2String(RSTMP!lastname) & ", " & Null2String(RSTMP!Firstname) & " " & Left(Null2String(RSTMP!MIDDLENAME), 1) & "."

            cboTECH.AddItem Null2String(rstmp!TECH_NAME)
            rstmp.MoveNext
        Loop
    End If

    Set rstmp = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    On Error GoTo Errorcode
    Dim Filter                                         As String

    RPT.Reset
    RPT.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    RPT.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    RPT.WindowTitle = "Technician Labor Sales"

    If UCase(cboTECH) <> "ALL" Then
        Filter = " AND {TECH.TECH_NAME}='" & Replace(cboTECH, "'", "") & "'"
    End If

    PrintSQLReport RPT, CSMS_REPORT_PATH & "TechnicianLaborSales.rpt", "Month({RO.DTE_COMP}) = " & What_month(cboMonth.Text) & " AND Year({RO.DTE_COMP}) = " & cboYear.Text & Filter, DMIS_REPORT_Connection, 1
    'LogAudit "V", "TECHNICIAN LABOR SALES - REPORTS ", cboTech & cboMonth & cboYEAR
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "TECHNICIAN REPORT", "", "", "", "TECHNICIAN LABOR SALES : " & cboTECH & " " & cboMonth & " " & cboYear, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    Screen.MousePointer = 0
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
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    FillCombo
    cboTECH.Text = "All"
End Sub


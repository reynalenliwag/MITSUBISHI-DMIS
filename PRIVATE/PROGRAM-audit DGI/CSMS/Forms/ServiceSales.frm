VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSServiceSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Advisor Sales"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   Icon            =   "ServiceSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboTech 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   1500
      TabIndex        =   6
      Top             =   90
      Width           =   2475
   End
   Begin VB.ComboBox cboMonth 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1500
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
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   930
      Width           =   2475
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
      Height          =   750
      Left            =   2625
      MouseIcon       =   "ServiceSales.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ServiceSales.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   1455
      Width           =   735
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   420
      Top             =   1095
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
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1905
      MouseIcon       =   "ServiceSales.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ServiceSales.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   1455
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Technician :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   180
      Width           =   2490
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
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
      Left            =   630
      TabIndex        =   5
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   660
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSServiceSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
'Function/Feature: Technician Attendance Report
'Date Started: 05/24/2007 2:47pm
'Last Update:
'Database Updates:
'Who Updated: Jonathan
'Updating Code: JAA - 05242007
'==============================================================

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "TECHNICIAN LABOR SALES") = False Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo Errorcode
    Dim Filter As String
    'Updating Code: JAA - 05242007
    '==========================================================================================
    
    
        rpt.Reset
        rpt.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rpt.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        If UCase(cboTech) <> "ALL" Then
        Filter = " AND {TECH.TECH_NAME}='" & Replace(cboTech, "'", "") & "'"
        
        End If
        PrintSQLReport rpt, CSMS_REPORT_PATH & "TechnicianLaborSales.rpt", "Month({RO.DTE_COMP}) = " & What_month(cboMonth.Text) & " AND Year({RO.DTE_COMP}) = " & cboYear.Text & Filter, DMIS_REPORT_Connection, 1
    '==========================================================================================
    'End of update

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
    FillcboYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    FillCombo
End Sub
Sub FillCombo()
    Dim tmp_value                                      As String
    Dim rsTechnician_Performance_Report                As ADODB.Recordset
    tmp_value = ""
    Set rsTechnician_Performance_Report = New ADODB.Recordset
    rsTechnician_Performance_Report.Open "Select Tech_Name from CSMS_JobClock order by Tech_Name asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTechnician_Performance_Report.EOF And Not rsTechnician_Performance_Report.BOF Then
        rsTechnician_Performance_Report.MoveFirst
        cboTech.Clear
        cboTech.AddItem "All"
        Do While Not rsTechnician_Performance_Report.EOF
            If tmp_value = Null2String(rsTechnician_Performance_Report!Tech_Name) Then
                rsTechnician_Performance_Report.MoveNext
            Else
                cboTech.AddItem Null2String(rsTechnician_Performance_Report!Tech_Name)
                tmp_value = Null2String(rsTechnician_Performance_Report!Tech_Name)
                rsTechnician_Performance_Report.MoveNext
            End If
        Loop
    End If
    Set rsTechnician_Performance_Report = Nothing
End Sub



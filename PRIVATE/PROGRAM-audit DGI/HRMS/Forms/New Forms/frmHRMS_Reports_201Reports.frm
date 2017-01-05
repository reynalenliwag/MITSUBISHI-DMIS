VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmHRMS_Reports_201Reports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Define 201 Reports"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14775
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_Reports_201Reports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   14775
   Begin XtremeReportControl.ReportControl rptLIST 
      Height          =   5655
      Left            =   30
      TabIndex        =   18
      Top             =   1440
      Width           =   14715
      _Version        =   655364
      _ExtentX        =   25956
      _ExtentY        =   9975
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   30
      ScaleHeight     =   5565
      ScaleWidth      =   2625
      TabIndex        =   21
      Top             =   7110
      Width           =   2655
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emp no."
         Enabled         =   0   'False
         Height          =   210
         Left            =   180
         TabIndex        =   38
         Top             =   480
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Employee Name"
         Enabled         =   0   'False
         Height          =   210
         Left            =   180
         TabIndex        =   37
         Top             =   796
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Address"
         Height          =   210
         Left            =   180
         TabIndex        =   36
         Top             =   1112
         Width           =   1005
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contact No."
         Height          =   240
         Left            =   180
         TabIndex        =   35
         Top             =   1428
         Width           =   1185
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Birth Date"
         Height          =   210
         Left            =   180
         TabIndex        =   34
         Top             =   1774
         Width           =   1035
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Religion"
         Height          =   210
         Left            =   180
         TabIndex        =   33
         Top             =   2090
         Width           =   915
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Citizenship"
         Height          =   210
         Left            =   180
         TabIndex        =   32
         Top             =   2406
         Width           =   1125
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SSS no"
         Height          =   210
         Left            =   180
         TabIndex        =   31
         Top             =   2722
         Width           =   885
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tin No."
         Height          =   210
         Left            =   180
         TabIndex        =   30
         Top             =   3038
         Width           =   855
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PhilHealth No."
         Height          =   210
         Left            =   180
         TabIndex        =   29
         Top             =   3354
         Width           =   1275
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pagibig No."
         Height          =   210
         Left            =   180
         TabIndex        =   28
         Top             =   3670
         Width           =   1155
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Position"
         Height          =   210
         Left            =   180
         TabIndex        =   27
         Top             =   3986
         Width           =   1065
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date Hired"
         Height          =   210
         Left            =   180
         TabIndex        =   26
         Top             =   4302
         Width           =   1125
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acct No."
         Height          =   210
         Left            =   180
         TabIndex        =   25
         Top             =   4618
         Width           =   975
      End
      Begin VB.CheckBox Check15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Allowance"
         Height          =   210
         Left            =   180
         TabIndex        =   24
         Top             =   4934
         Width           =   1155
      End
      Begin VB.CheckBox Check16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Basic Pay"
         Height          =   210
         Left            =   180
         TabIndex        =   23
         Top             =   5250
         Width           =   1095
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   14805
         _Version        =   655364
         _ExtentX        =   26114
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   " Choose Field(s) to Display"
         ForeColor       =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   4194304
      End
   End
   Begin VB.CheckBox Check17 
      Caption         =   "Print in Excel"
      Height          =   210
      Left            =   13410
      TabIndex        =   19
      Top             =   360
      Width           =   1275
   End
   Begin VB.ComboBox cboStat 
      Height          =   330
      ItemData        =   "frmHRMS_Reports_201Reports.frx":058A
      Left            =   7770
      List            =   "frmHRMS_Reports_201Reports.frx":0597
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   300
      Width           =   1425
   End
   Begin VB.ComboBox cboEXM 
      Height          =   330
      ItemData        =   "frmHRMS_Reports_201Reports.frx":05B1
      Left            =   9210
      List            =   "frmHRMS_Reports_201Reports.frx":05BE
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   300
      Width           =   1635
   End
   Begin VB.ComboBox cboInc 
      Height          =   330
      ItemData        =   "frmHRMS_Reports_201Reports.frx":05D5
      Left            =   10890
      List            =   "frmHRMS_Reports_201Reports.frx":05DF
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   300
      Width           =   1485
   End
   Begin VB.ComboBox cboDept 
      Height          =   330
      ItemData        =   "frmHRMS_Reports_201Reports.frx":05F5
      Left            =   2520
      List            =   "frmHRMS_Reports_201Reports.frx":05FF
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   300
      Width           =   3405
   End
   Begin VB.ComboBox cboSEX 
      Height          =   330
      ItemData        =   "frmHRMS_Reports_201Reports.frx":0611
      Left            =   1470
      List            =   "frmHRMS_Reports_201Reports.frx":061E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   300
      Width           =   1035
   End
   Begin VB.ComboBox cboLevel 
      Height          =   330
      ItemData        =   "frmHRMS_Reports_201Reports.frx":0635
      Left            =   120
      List            =   "frmHRMS_Reports_201Reports.frx":0642
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   1305
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9585
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   525
      Left            =   12870
      MouseIcon       =   "frmHRMS_Reports_201Reports.frx":065F
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Reports_201Reports.frx":07B1
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print this Record"
      Top             =   180
      Width           =   465
   End
   Begin VB.CommandButton cmdFind 
      Height          =   525
      Left            =   12420
      MouseIcon       =   "frmHRMS_Reports_201Reports.frx":0B17
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Reports_201Reports.frx":0C69
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Find a Record"
      Top             =   180
      Width           =   465
   End
   Begin VB.ComboBox cboTYPE 
      Height          =   330
      ItemData        =   "frmHRMS_Reports_201Reports.frx":0F63
      Left            =   6000
      List            =   "frmHRMS_Reports_201Reports.frx":0F70
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   300
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   8
      Left            =   7830
      TabIndex        =   17
      Top             =   60
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exm. Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   7
      Left            =   9270
      TabIndex        =   16
      Top             =   60
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ACTIVE/ INACTIVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   5
      Left            =   10920
      TabIndex        =   15
      Top             =   60
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   4
      Left            =   6030
      TabIndex        =   10
      Top             =   60
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   3
      Left            =   2610
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   2
      Left            =   1530
      TabIndex        =   6
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Level"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1035
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   1425
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   14805
      _Version        =   655364
      _ExtentX        =   26114
      _ExtentY        =   2514
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmHRMS_Reports_201Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                           As Excel.Application
Dim xlbook                                          As Excel.Workbook
Dim xlsheet                                         As Excel.Worksheet
Dim SQL_QUERY As String

Sub PRINTEXCEL()
    If rptLIST.Records.count = 0 Then
        MessagePop InfoFriend, "No Record", "Kindly search for Record first"
        Exit Sub
    End If
    
    Dim RSTMP                                       As New ADODB.Recordset
    Dim SQL As String
    Dim xcnt                                        As Integer
    Set xlApp = CreateObject("Excel.Application")
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "UserDefine201Reports.xls")
    Set xlsheet = xlbook.Worksheets(1)
    
    SQL = Replace(SQL_QUERY, "LASTNAME + ', ' + FIRSTNAME + ' ' + MIDDLENAME + '.' AS FULLNAME         ", "LASTNAME, FIRSTNAME, MIDDLENAME, SEX, STATUS, EXSTATUS, PERSON, RELATION, RELTELNO")
    SQL = Replace(SQL, "FULLNAME", "LASTNAME")
    Set RSTMP = gconDMIS.Execute(SQL)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        xlsheet.Range("A3").CopyFromRecordset RSTMP
    End If

    xlApp.Windows.ITEM(1).Caption = "User Define 201 Reports"
    xlApp.Visible = True
    Set xlApp = Nothing
    Set xlsheet = Nothing
    Set xlbook = Nothing
End Sub

Private Sub cmdFind_Click()
    Dim SQL As String
    Dim QUERY As String
    Dim C1 As String:    Dim c2 As String:    Dim c3 As String:    Dim c4 As String
    Dim c5 As String:    Dim c6 As String:    Dim C7 As String:    Dim C8 As String
    Dim C9 As String:    Dim C10 As String:    Dim C11 As String:    Dim C12 As String
    Dim C13 As String:    Dim C14 As String:    Dim C15 As String:    Dim C16 As String
    Dim S1 As String:    Dim S2 As String:    Dim S3 As String:    Dim S4 As String
    Dim S5 As String:    Dim S6 As String:    Dim S7 As String:    Dim S8 As String
    Dim S9 As String:    Dim S10 As String:    Dim S11 As String:    Dim S12 As String
    Dim S13 As String:    Dim S14 As String:    Dim S15 As String:    Dim S16 As String
    Dim HEADER As String
    
    C1 = "EMP NO.": c2 = ", EMPLOYEE NAME"
    If Check3.Value = 1 Then c3 = ", ADDRESS": If Not Check3.Value = 1 Then c3 = ""
    If Check4.Value = 1 Then c4 = ", CONTACT NO.": If Not Check4.Value = 1 Then c4 = ""
    If Check5.Value = 1 Then c5 = ", BIRTH DATE": If Not Check5.Value = 1 Then c5 = ""
    If Check6.Value = 1 Then c6 = ", RELIGION": If Not Check6.Value = 1 Then c6 = ""
    If Check7.Value = 1 Then C7 = ", CITIZENSHIP": If Not Check7.Value = 1 Then C7 = ""
    If Check8.Value = 1 Then C8 = ", SSS NO.": If Not Check8.Value = 1 Then C8 = ""
    If Check9.Value = 1 Then C9 = ", TIN NO.": If Not Check9.Value = 1 Then C9 = ""
    If Check10.Value = 1 Then C10 = ", PHIL. NO.": If Not Check10.Value = 1 Then C10 = ""
    If Check11.Value = 1 Then C11 = ", PAGIBIG NO.": If Not Check11.Value = 1 Then C11 = ""
    If Check12.Value = 1 Then C12 = ", POSITION": If Not Check12.Value = 1 Then C12 = ""
    If Check13.Value = 1 Then C13 = ", DATE HIRED": If Not Check13.Value = 1 Then C13 = ""
    If Check14.Value = 1 Then C14 = ", ACCT NO.": If Not Check14.Value = 1 Then C14 = ""
    If Check15.Value = 1 Then C15 = ", ALLOWANCE": If Not Check15.Value = 1 Then C15 = ""
    If Check16.Value = 1 Then C16 = ", BASIC PAY": If Not Check16.Value = 1 Then C16 = ""
    
    If Check3.Value = 1 Then S3 = ", ADDRESS": If Not Check3.Value = 1 Then S3 = ""
    If Check4.Value = 1 Then S4 = ", TELEPHONE": If Not Check4.Value = 1 Then S4 = ""
    If Check5.Value = 1 Then S5 = ", BIRTHDATE": If Not Check5.Value = 1 Then S5 = ""
    If Check6.Value = 1 Then S6 = ", RELIGION": If Not Check6.Value = 1 Then S6 = ""
    If Check7.Value = 1 Then S7 = ", CITIZEN": If Not Check7.Value = 1 Then S7 = ""
    If Check8.Value = 1 Then S8 = ", SSSNO": If Not Check8.Value = 1 Then S8 = ""
    If Check9.Value = 1 Then S9 = ", TINNO": If Not Check9.Value = 1 Then S9 = ""
    If Check10.Value = 1 Then S10 = ", PHNO": If Not Check10.Value = 1 Then S10 = ""
    If Check11.Value = 1 Then S11 = ", PAGIBIGNO": If Not Check11.Value = 1 Then S11 = ""
    If Check12.Value = 1 Then S12 = ", POSITION": If Not Check12.Value = 1 Then S12 = ""
    If Check13.Value = 1 Then S13 = ", DATEHIRED": If Not Check13.Value = 1 Then S13 = ""
    If Check14.Value = 1 Then S14 = ", ACCOUNTNO": If Not Check14.Value = 1 Then S14 = ""
    If Check15.Value = 1 Then S15 = ", ALLOWANCE": If Not Check15.Value = 1 Then S15 = ""
    If Check16.Value = 1 Then S16 = ", BASICSALARY": If Not Check16.Value = 1 Then S16 = ""
    
    HEADER = C1 & c2 & c3 & c4 & c5 & c6 & C7 & C8 & C9 & C10 & C11 & C12 & C13 & C14 & C15 & C16
    
    Dim XACTINC As String:    Dim XLEVEL As String:    Dim XSEX As String:    Dim XDEPT As String
    Dim XTYPE As String:    Dim XSTAT As String:    Dim XEXM As String
    
    If cboInc.Text = "Active" Then XACTINC = " ACTIVEINACTIVE = 'A'"
    If Not cboInc.Text = "Active" Then XACTINC = " ACTIVEINACTIVE = 'I'"
    
    If cboLevel.Text = "All" Then XLEVEL = ""
    If cboLevel.Text = "Rank End" Then XLEVEL = " AND EMPLEVEL = 'E'"
    If cboLevel.Text = "Managers" Then XLEVEL = " AND EMPLEVEL = 'M'"
    
    If cboSEX.Text = "All" Then XSEX = ""
    If cboSEX.Text = "Male" Then XSEX = " AND SEX = 'M'"
    If cboSEX.Text = "Female" Then XSEX = " AND SEX = 'F'"
    
    If cboDept.Text = "All" Then XDEPT = ""
    If Not cboDept.Text = "All" Then XDEPT = " AND DEPTCODE = '" & FindDepartmentCode(cboDept) & "'"
    
    If cboTYPE = "All" Then XTYPE = ""
    If cboTYPE = "Semi-Monthly" Then XTYPE = " AND PAYROLLTYPE = 'Semi-Monthly Base'"
    If cboTYPE = "Monthly" Then XTYPE = " AND PAYROLLTYPE = 'Monthly Base'"
    
    If cboStat = "All" Then XSTAT = ""
    If cboStat = "Single" Then XSTAT = " AND STATUS = 'Single'"
    If cboStat = "Married" Then XSTAT = " AND STATUS = 'Married'"
    If cboEXM.Text = "All" Then XEXM = ""
    If Not cboEXM.Text = "All" Then XEXM = " AND EXSTATUS = '" & cboEXM & "'"
    
    QUERY = "SELECT EMPNO, LASTNAME + ', ' + FIRSTNAME + ' ' + MIDDLENAME + '.' AS FULLNAME         " & S3 & S4 & S5 & S6 & S7 & S8 & S9 & S10 & S11 & S12 & S13 & S14 & S15 & S16
    SQL = QUERY & " FROM HRMS_EMPINFO WHERE " & XACTINC & XLEVEL & XSEX & XDEPT & XTYPE & XSTAT & XEXM & " ORDER BY FULLNAME"
'    Call DisplaySearch(SQL, HEADER)
    SQL_QUERY = SQL
    
    Call FillGrid(SQL)
End Sub

Sub DisplaySearch(XXX As String, XHEADER As String)
    Screen.MousePointer = 11
    Call ReportControlAddColumnHeader(rptLIST, XHEADER)
    Call ReportControlPaintManager(rptLIST)
    'rptLIST.Columns(0).Alignment = xtpAlignmentCenter
    'rptLIST.Columns(1).Alignment = xtpAlignmentCenter
    'rptLIST.Columns(2).Alignment = xtpAlignmentCenter
    'rptLIST.Columns(3).Alignment = xtpAlignmentRight
    'rptLIST.Columns(4).Alignment = xtpAlignmentRight
    'rptLIST.Columns(5).Alignment = xtpAlignmentRight
    'rptLIST.Columns(9).Alignment = xtpAlignmentCenter
    'rptLIST.Columns(10).Alignment = xtpAlignmentCenter
    'rptLIST.Columns(11).Alignment = xtpAlignmentCenter
    Call ResizeColumnHeader(rptLIST, " 5, 15, 15, 9, 8, 11, 8, 6, 8,5, 5")
    Call flex_FillReportView(gconDMIS.Execute(XXX), rptLIST)
    Screen.MousePointer = 0
End Sub

Sub FillExmStatus()
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DISTINCT EXSTATUS FROM HRMS_EMPINFO ORDER BY EXSTATUS")
    cboEXM.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboEXM.AddItem Null2String(RSTMP!EXSTATUS)
            RSTMP.MoveNext
        Loop
        cboEXM.AddItem "All"
        cboEXM.Text = "All"
    End If
    Set RSTMP = Nothing
End Sub

Function FindDepartmentCode(XXX As String) As String
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DEPTCODE FROM HRMS_DEPARTMENT WHERE DEPTNAME = '" & XXX & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindDepartmentCode = Null2String(RSTMP!DEPTCODE)
    End If
    Set RSTMP = Nothing
End Function

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    If Check1.Value = 0 Then
        If rptLIST.Records.count <= 0 Then Exit Sub
        rptLIST.PrintOptions.HEADER.TextCenter = "User Define 201 Reports"
        rptLIST.PrintOptions.MarginLeft = 0.5
        rptLIST.PrintOptions.MarginRight = 0.5
        'rptLIST.PrintOptions.
        rptLIST.PrintPreview True
    Else
        Call PRINTEXCEL
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    cboLevel = "All"
    cboStat = "All"
    cboSEX.Text = "All"
    cboTYPE.Text = "All"
    cboInc.ListIndex = 0
    Call fillDepartment
    Call FillExmStatus
    
    Call InitializeRC
End Sub

Sub fillDepartment()
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DEPTNAME FROM HRMS_DEPARTMENT ORDER BY DEPTNAME")
    cboDept.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboDept.AddItem Null2String(RSTMP!DEPTNAME)
            RSTMP.MoveNext
        Loop
        cboDept.AddItem "All"
        cboDept.Text = "All"
    End If
    Set RSTMP = Nothing
End Sub


Sub ReportControlAddColumnHeader(lst As ReportControl, StringHeaders As String)
    Dim ar()                                        As String
    Dim I                                           As Integer

    ar = Split(StringHeaders, ",")
    lst.Columns.DeleteAll
    For I = LBound(ar) To UBound(ar)
        lst.Columns.Add I, ar(I), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(lst As ReportControl)
    With lst
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                        As String
    Dim cWidth                                      As Long
    Dim I                                           As Integer
    Dim scwidth                                     As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                         As ADODB.Field
    Dim j                                           As Long
    Dim REC                                         As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.FIELDS
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Private Sub ShortcutCaption1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 2 Then
        If Picture1.Top = 7110 Then
            Picture1.Top = 1830
            Picture1.ZOrder 0
        Else
            Picture1.Top = 7110
            Picture1.ZOrder 1
        End If
    End If
End Sub

Private Sub txtsearch_Change()
    rptLIST.FilterText = txtSearch
    rptLIST.Populate
End Sub

Sub InitializeRC()
    Dim X As Integer
    X = 0
    With rptLIST
        .Columns.DeleteAll
        If Check1.Value = 1 Then .Columns.Add X, "Emp no.", 50, True:           .Columns(X).Alignment = xtpAlignmentCenter:     .Columns(X).AllowRemove = False: X = X + 1
        If Check2.Value = 1 Then .Columns.Add X, "Employee Name", 200, True:    .Columns(X).Alignment = xtpAlignmentLeft:       .Columns(X).AllowRemove = False: X = X + 1
        If Check3.Value = 1 Then .Columns.Add X, "Address", 200, True:          .Columns(X).Alignment = xtpAlignmentLeft:       .Columns(X).AllowRemove = False: X = X + 1
        If Check4.Value = 1 Then .Columns.Add X, "Contact no", 80, True:        .Columns(X).Alignment = xtpAlignmentLeft:       .Columns(X).AllowRemove = False: X = X + 1
        If Check5.Value = 1 Then .Columns.Add X, "Birth Date", 80, True:        .Columns(X).Alignment = xtpAlignmentCenter:     .Columns(X).AllowRemove = False: X = X + 1
        If Check6.Value = 1 Then .Columns.Add X, "Religion", 100, True:         .Columns(X).Alignment = xtpAlignmentLeft:       .Columns(X).AllowRemove = False: X = X + 1
        If Check7.Value = 1 Then .Columns.Add X, "Citizenship", 100, True:      .Columns(X).Alignment = xtpAlignmentLeft:       .Columns(X).AllowRemove = False: X = X + 1
        If Check8.Value = 1 Then .Columns.Add X, "SSS no", 80, True:            .Columns(X).Alignment = xtpAlignmentCenter:     .Columns(X).AllowRemove = False: X = X + 1
        If Check9.Value = 1 Then .Columns.Add X, "TIN no", 80, True:            .Columns(X).Alignment = xtpAlignmentCenter:     .Columns(X).AllowRemove = False: X = X + 1
        If Check10.Value = 1 Then .Columns.Add X, "Philhealth no", 80, True:    .Columns(X).Alignment = xtpAlignmentCenter:     .Columns(X).AllowRemove = False: X = X + 1
        If Check11.Value = 1 Then .Columns.Add X, "Pagibig no", 90, True:       .Columns(X).Alignment = xtpAlignmentCenter:     .Columns(X).AllowRemove = False: X = X + 1
        If Check12.Value = 1 Then .Columns.Add X, "Position", 150, True:        .Columns(X).Alignment = xtpAlignmentLeft:       .Columns(X).AllowRemove = False: X = X + 1
        If Check13.Value = 1 Then .Columns.Add X, "Date Hired", 80, True:       .Columns(X).Alignment = xtpAlignmentCenter:     .Columns(X).AllowRemove = False: X = X + 1
        If Check14.Value = 1 Then .Columns.Add X, "Acct no", 80, True:          .Columns(X).Alignment = xtpAlignmentCenter:     .Columns(X).AllowRemove = False: X = X + 1
        If Check15.Value = 1 Then .Columns.Add X, "Allowance", 100, True:       .Columns(X).Alignment = xtpAlignmentRight:      .Columns(X).AllowRemove = False: X = X + 1
        If Check16.Value = 1 Then .Columns.Add X, "Basic Pay", 100, True:       .Columns(X).Alignment = xtpAlignmentRight:      .Columns(X).AllowRemove = False: X = X + 1
       
        
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
    End With
End Sub

Sub FillGrid(XXX As String)
    Dim RSUPLOAD                                        As New ADODB.Recordset
    Dim REC                                             As XtremeReportControl.ReportRecord
    Dim X                                               As Integer
    
    X = 0
    Call InitializeRC
    Set RSUPLOAD = gconDMIS.Execute(XXX)
    rptLIST.Records.DeleteAll
    While Not RSUPLOAD.EOF
        Set REC = rptLIST.Records.Add
        X = 0
        If Check1.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check2.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check3.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check4.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check5.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check6.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check7.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check8.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check9.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check10.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check11.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check12.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check13.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check14.Value = 1 Then REC.AddItem (Trim(RSUPLOAD.FIELDS(X))):    X = X + 1
        If Check15.Value = 1 Then REC.AddItem ToDoubleNumber((Trim(RSUPLOAD.FIELDS(X)))):     X = X + 1
        If Check16.Value = 1 Then REC.AddItem ToDoubleNumber((Trim(RSUPLOAD.FIELDS(X)))):     X = X + 1
        
        RSUPLOAD.MoveNext
        Set REC = Nothing
    Wend
    rptLIST.Populate
End Sub

Private Sub TXTSEARCH_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

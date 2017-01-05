VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_VehicleSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicles Sales Report"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_VehicleSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   4125
   Begin VB.ComboBox cboReportType 
      Height          =   360
      ItemData        =   "Report_VehicleSales.frx":0E42
      Left            =   450
      List            =   "Report_VehicleSales.frx":0E64
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   390
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2100
      MouseIcon       =   "Report_VehicleSales.frx":0F9B
      MousePointer    =   99  'Custom
      Picture         =   "Report_VehicleSales.frx":10ED
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   3360
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1230
      MouseIcon       =   "Report_VehicleSales.frx":1538
      MousePointer    =   99  'Custom
      Picture         =   "Report_VehicleSales.frx":168A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   3360
      Width           =   885
   End
   Begin Crystal.CrystalReport rptSales 
      Left            =   3390
      Top             =   3420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Yearly Ending Inventory"
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
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   390
      TabIndex        =   3
      Top             =   750
      Width           =   3495
      Begin VB.OptionButton optYear 
         Caption         =   "Yearly"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2370
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "Monthly"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1005
      End
      Begin VB.OptionButton optRange 
         Caption         =   "Ranged"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.PictureBox picYearly 
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   360
      ScaleHeight     =   1515
      ScaleWidth      =   4365
      TabIndex        =   8
      Top             =   1680
      Width           =   4365
      Begin VB.TextBox txtYearly_Year 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         MaxLength       =   4
         TabIndex        =   16
         Top             =   420
         Width           =   2805
      End
      Begin VB.Label Label4 
         Caption         =   "For The Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   1515
      End
   End
   Begin VB.PictureBox picMonthly 
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   390
      ScaleHeight     =   1635
      ScaleWidth      =   4365
      TabIndex        =   9
      Top             =   1650
      Width           =   4365
      Begin VB.TextBox txtMonthly_Year 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   60
         MaxLength       =   4
         TabIndex        =   19
         Top             =   1020
         Width           =   3165
      End
      Begin VB.ComboBox cboMonthly_Month 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   3195
      End
      Begin VB.Label Label6 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   20
         Top             =   750
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "For the Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   15
         Top             =   90
         Width           =   1815
      End
   End
   Begin VB.PictureBox picRange 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   360
      ScaleHeight     =   1485
      ScaleWidth      =   4365
      TabIndex        =   7
      Top             =   1650
      Width           =   4365
      Begin MSComCtl2.DTPicker dt_Range_From 
         Height          =   390
         Left            =   60
         TabIndex        =   10
         Top             =   315
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   52363265
         CurrentDate     =   39203
      End
      Begin MSComCtl2.DTPicker dt_Range_To 
         Height          =   390
         Left            =   60
         TabIndex        =   11
         Top             =   1050
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   52363265
         CurrentDate     =   39203
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   13
         Top             =   765
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   12
         Top             =   30
         Width           =   885
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Select Your Report Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   390
      TabIndex        =   18
      Top             =   60
      Width           =   2835
   End
End
Attribute VB_Name = "frmSMIS_Report_VehicleSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboReportType_Click()
    If UCase(cboReportType) = "VEHICLES SALES BY PROJECTION" Then
        picMonthly.Visible = True
        cboMonthly_Month = MonthName(Month(LOGDATE))
        Frame1.Enabled = False

        picMonthly.Enabled = False
        picYearly.Enabled = False
        picRange.Enabled = False
    Else
        Frame1.Enabled = True
        picMonthly.Enabled = True
        picYearly.Enabled = True
        picRange.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim CRYS_FILTER                                                   As String
    Dim MonthPrint                                                    As String
    Dim Filtered                                                      As String
    rptSales.Reset

    If picYearly.Visible = True Then
        If IsNumeric(txtYearly_Year) = False Then
            MsgSpeechBox " Wrong Input, Invalid Year!"
            txtYearly_Year.SetFocus
            Exit Sub
        End If

    ElseIf picMonthly.Visible = True Then

        If IsNumeric(txtMonthly_Year) = False Then
            MsgSpeechBox " Wrong Input, Invalid Year!"
            txtMonthly_Year.SetFocus
            Exit Sub
        End If

        If What_month(cboMonthly_Month) > 12 Or What_month(cboMonthly_Month) <= 0 Then
            MsgSpeechBox " Select Appropriate Month From The List!"
            cboMonthly_Month.SetFocus
            Exit Sub
        End If
    ElseIf picRange.Visible = True Then

    Else
        Exit Sub
    End If


    If picYearly.Visible = True Then
        CRYS_FILTER = "year({PURCHAGREE.DateReleased})=" & txtYearly_Year

        MonthPrint = "FOR THE YEAR OF " & txtYearly_Year

    ElseIf picMonthly.Visible = True Then

        CRYS_FILTER = "year({PURCHAGREE.DateReleased})=" & txtMonthly_Year & " AND month({PURCHAGREE.DateReleased}) = " & What_month(cboMonthly_Month)

        MonthPrint = "FROM THE MONTH OF " & cboMonthly_Month & " " & txtYearly_Year

    ElseIf picRange.Visible = True Then

        '        CRYS_FILTER = "((((year({PURCHAGREE.DateReleased}) >=" & Year(dt_Range_From) & " AND month({PURCHAGREE.DateReleased}) >= " & Month(dt_Range_From) & " AND Day({PURCHAGREE.DateReleased}) >= " & Day(dt_Range_From) & ")))"
        '        CRYS_FILTER = CRYS_FILTER & " AND " & " (((year({PURCHAGREE.DateReleased}) <= " & Year(dt_Range_To) & " AND month({PURCHAGREE.DateReleased}) <= " & Month(dt_Range_To) & " AND Day({PURCHAGREE.DateReleased}) <= " & Day(dt_Range_To) & " ))))"

        'THIS IS THE NEW FILTERING
        CRYS_FILTER = "{PURCHAGREE.DATERELEASED} >= DATE(" & Year(dt_Range_From.Value) & "," & Month(dt_Range_From.Value) & ", " & Day(dt_Range_From.Value) & ") AND {PURCHAGREE.DATERELEASED} <=DATE(" & Year(dt_Range_To.Value) & "," & Month(dt_Range_To.Value) & "," & Day(dt_Range_To.Value) & ")"

        MonthPrint = "FROM : " & Format(dt_Range_From, "MM/DD/YYYY") & " TO: " & Format(dt_Range_To, "MM/DD/YYYY")

    Else
        Exit Sub

    End If


    rptSales.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptSales.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptSales.Formulas(2) = "MonthPrint = '" & UCase(MonthPrint) & "'"

    Select Case UCase(LTrim(RTrim(cboReportType)))
        Case "VEHICLES SALES SEASONAL REPORT"
            If Year(dt_Range_From) <> Year(dt_Range_To) Then
                MsgBox "Please Select Only Range With In Year", vbInformation
                On Error Resume Next
                dt_Range_From.SetFocus
                Exit Sub
            End If
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\SESONAL REPORT.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY MODEL"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\MODEL.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY COLOR"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\COLOR.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY SAE"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\SAE.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY GENDER"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\GENDER.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY CUSTOMER TYPE"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\CUSTOMER TYPE.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY ACCOUNT TYPE"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\ACCOUNT TYPE.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY MODE OF PAYMENT"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\MODE OF PAYMENT.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY FINANCING COMPANY"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\FINANCING COMPANY.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY INSURANCE COMPANY"
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\INSURANCE COMPANY.RPT", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
        Case "VEHICLES SALES BY PROJECTION"    '*******RYAN CULAWAY MAY 8 11:46PM
            'rptSales.Formulas(3) = "curr_month=cdate('1/" & What_month(cboMonthly_Month) & "/" & txtYearly_Year & "')"
            CRYS_FILTER = ""
            PrintSQLReport rptSales, SMIS_REPORT_PATH & "VS\VehiclesSalesProjection.rpt", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
            '******************RYAN CULAWAY MAY 8 11:46PM
    End Select
    'UPDATED BY: JUN
    'DATE UPDATED: 09032008 5:00
    
    If optMonth.Value = True Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES REPORTS", "", "", "", cboReportType & " " & cboMonthly_Month & " " & txtMonthly_Year, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf optRange.Value = True Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES REPORTS", "", "", "", cboReportType & " " & dt_Range_From & " " & dt_Range_To, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf optYear.Value = True Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES REPORTS", "", "", "", cboReportType & " " & txtYearly_Year, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Else
    End If
    
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES REPORTS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SALES REPORTS", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dt_Range_To.Value = firstDay(Date)
    dt_Range_To.Value = Date

    fillcbomonth cboMonthly_Month
    '
    cboReportType.ListIndex = 0
    dt_Range_To.Value = Date
    dt_Range_From = firstDay(Date)
    txtYearly_Year = Year(LOGDATE)
    txtMonthly_Year = Year(LOGDATE)
    cboMonthly_Month = MonthName(Month(LOGDATE))
    optMonth.Value = True
    'Vehicles Sales Sesonal Report
    'Vehicles Sales by Model
    'Vehicles Sales by Color
    'Vehicles Sales by SAE
    'Vehicles Sales by Gender
    'Vehicles Sales by Customer Type
    'Vehicles Sales by Mode of Payment


End Sub


Private Sub optMonth_Click()
    If optMonth.Value = True Then: picMonthly.Visible = True: picRange.Visible = False: picYearly.Visible = False: picMonthly.ZOrder 0
End Sub

Private Sub optRange_Click()
    If optRange.Value = True Then: picMonthly.Visible = False: picRange.Visible = True: picYearly.Visible = False:
End Sub

Private Sub optYear_Click()
    If optYear.Value = True Then: picMonthly.Visible = False: picRange.Visible = False: picYearly.Visible = True:
End Sub


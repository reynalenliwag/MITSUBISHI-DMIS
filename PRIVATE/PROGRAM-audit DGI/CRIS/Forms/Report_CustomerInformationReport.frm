VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Report_CustomerInformationReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Information Report"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Report_CustomerInformationReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "Daily Customer Summary Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   34
      ToolTipText     =   "Displays Total Customer/Prospect Information For Given Range With Vehicle Information"
      Top             =   1800
      Width           =   4215
   End
   Begin VB.OptionButton optCustomerInfo 
      Caption         =   "Customer Information Report (Group By)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   9
      ToolTipText     =   "Detail Customer Information Registered In Company Data Base Filtered By Various Group"
      Top             =   30
      Width           =   4125
   End
   Begin VB.OptionButton optRecommendation 
      Caption         =   "Summary of Customer Suggestion /Recommendation-Service"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   60
      TabIndex        =   8
      ToolTipText     =   "Service Customer Suggestion and Recommendation Summary Report"
      Top             =   1020
      Width           =   4125
   End
   Begin VB.OptionButton optFeedBack 
      Caption         =   "Summary Customer Followups- Service"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   7
      ToolTipText     =   "Service Customer Feed Back Summary Report"
      Top             =   690
      Width           =   4125
   End
   Begin VB.OptionButton optReturningCustomer 
      Caption         =   "Returning Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "Customer with Trasnsaction History and Non Returing In (n) no of Days"
      Top             =   390
      Width           =   4125
   End
   Begin Crystal.CrystalReport rptCustomer_Information 
      Left            =   4020
      Top             =   870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Customer Information Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "KSFJLSDJFLJDLFJLKSDJF"
      BoundReportFooter=   -1  'True
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2145
      MouseIcon       =   "Report_CustomerInformationReport.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Report_CustomerInformationReport.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   3570
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1275
      MouseIcon       =   "Report_CustomerInformationReport.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "Report_CustomerInformationReport.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   3570
      Width           =   885
   End
   Begin VB.OptionButton optCustProsp 
      Caption         =   "Customer/Prospect Vehicle Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Displays Total Customer/Prospect Information For Given Range With Vehicle Information"
      Top             =   1500
      Width           =   4215
   End
   Begin VB.PictureBox picCV 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   4095
      TabIndex        =   24
      Top             =   2100
      Visible         =   0   'False
      Width           =   4125
      Begin MSComCtl2.DTPicker CV_1 
         Height          =   345
         Left            =   60
         TabIndex        =   25
         Top             =   270
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         Format          =   52756481
         CurrentDate     =   39484
      End
      Begin MSComCtl2.DTPicker CV_2 
         Height          =   345
         Left            =   1920
         TabIndex        =   26
         Top             =   270
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         Format          =   52756481
         CurrentDate     =   39484
      End
      Begin VB.Label Label1 
         Caption         =   "To Date"
         Height          =   285
         Left            =   1890
         TabIndex        =   27
         Top             =   30
         Width           =   2265
      End
      Begin VB.Label Label8 
         Caption         =   "From"
         Height          =   285
         Left            =   60
         TabIndex        =   28
         Top             =   30
         Width           =   2265
      End
   End
   Begin VB.PictureBox picSummaryFollow 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   4095
      TabIndex        =   19
      Top             =   2100
      Visible         =   0   'False
      Width           =   4125
      Begin MSComCtl2.DTPicker FDate 
         Height          =   345
         Left            =   60
         TabIndex        =   21
         Top             =   270
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         Format          =   52756481
         CurrentDate     =   39484
      End
      Begin MSComCtl2.DTPicker TDate 
         Height          =   345
         Left            =   1920
         TabIndex        =   22
         Top             =   270
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         Format          =   52756481
         CurrentDate     =   39484
      End
      Begin VB.Label Label7 
         Caption         =   "To Date(Released)"
         Height          =   285
         Left            =   1890
         TabIndex        =   23
         Top             =   30
         Width           =   2265
      End
      Begin VB.Label Label6 
         Caption         =   "From Date(Release)"
         Height          =   285
         Left            =   60
         TabIndex        =   20
         Top             =   30
         Width           =   2265
      End
   End
   Begin VB.PictureBox picCustInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   4095
      TabIndex        =   10
      Top             =   2100
      Visible         =   0   'False
      Width           =   4125
      Begin VB.ComboBox cboGroupBy 
         Height          =   345
         ItemData        =   "Report_CustomerInformationReport.frx":19D0
         Left            =   90
         List            =   "Report_CustomerInformationReport.frx":19D2
         TabIndex        =   11
         Top             =   300
         Width           =   3765
      End
      Begin VB.Label Label2 
         Caption         =   "Group By"
         Height          =   285
         Left            =   60
         TabIndex        =   12
         Top             =   30
         Width           =   2265
      End
   End
   Begin VB.PictureBox picRetCust 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   4095
      TabIndex        =   4
      Top             =   2100
      Visible         =   0   'False
      Width           =   4125
      Begin VB.TextBox txtNDaysVisit 
         Height          =   360
         Left            =   1590
         TabIndex        =   6
         Text            =   "0"
         Top             =   300
         Width           =   2205
      End
      Begin VB.ComboBox cboDepartment 
         Height          =   345
         ItemData        =   "Report_CustomerInformationReport.frx":19D4
         Left            =   60
         List            =   "Report_CustomerInformationReport.frx":19D6
         TabIndex        =   5
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "(n) Days Last Visited"
         Height          =   255
         Left            =   1620
         TabIndex        =   18
         Top             =   30
         Width           =   1875
      End
      Begin VB.Label Label4 
         Caption         =   "Department"
         Height          =   255
         Left            =   90
         TabIndex        =   17
         Top             =   60
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   4095
      TabIndex        =   29
      Top             =   2100
      Visible         =   0   'False
      Width           =   4125
      Begin MSComCtl2.DTPicker dtFortheDay 
         Height          =   375
         Left            =   1920
         TabIndex        =   33
         Top             =   270
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   52756481
         CurrentDate     =   39489
      End
      Begin VB.ComboBox cboDepartment2 
         Height          =   345
         ItemData        =   "Report_CustomerInformationReport.frx":19D8
         Left            =   90
         List            =   "Report_CustomerInformationReport.frx":19DA
         TabIndex        =   30
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label10 
         Caption         =   "For the Day"
         Height          =   285
         Left            =   1860
         TabIndex        =   32
         Top             =   30
         Width           =   2265
      End
      Begin VB.Label Label9 
         Caption         =   "Department"
         Height          =   285
         Left            =   60
         TabIndex        =   31
         Top             =   30
         Width           =   2265
      End
   End
   Begin VB.PictureBox picSuggestion 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   4095
      TabIndex        =   14
      Top             =   2100
      Visible         =   0   'False
      Width           =   4125
      Begin VB.ComboBox cboSuggestions 
         Height          =   345
         ItemData        =   "Report_CustomerInformationReport.frx":19DC
         Left            =   90
         List            =   "Report_CustomerInformationReport.frx":19DE
         TabIndex        =   15
         Top             =   300
         Width           =   3765
      End
      Begin VB.Label Label3 
         Caption         =   "Suggestion"
         Height          =   285
         Left            =   60
         TabIndex        =   16
         Top             =   30
         Width           =   2265
      End
   End
   Begin VB.Label labDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   60
      TabIndex        =   13
      Top             =   2880
      Width           =   4125
   End
End
Attribute VB_Name = "frmCRIS_Report_CustomerInformationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub InitCbo()
    cboDepartment.AddItem ("SERVICE")
    cboDepartment.AddItem ("SALES")
    With cboGroupBy
        .AddItem ("CustomerType")
        .AddItem ("Gender")
        .AddItem ("LeadSource")
        .AddItem ("Location")
        .AddItem ("Position")
        .AddItem ("BirthDate")
    End With

    With cboDepartment2
        .AddItem ("SERVICE")
        .AddItem ("SALES")
        .AddItem ("PARTS")
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    Dim RecordSelection                                               As String
    Dim lng                                                           As Integer
    Dim rsCusInfo                                                     As ADODB.Recordset
    Set rsCusInfo = New ADODB.Recordset

    rptCustomer_Information.Reset
    rptCustomer_Information.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptCustomer_Information.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptCustomer_Information.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"



    If optCustomerInfo.Value = True Then
        Set rsCusInfo = gconDMIS.Execute("Select COUNT(*)  from All_Customer")
        If rsCusInfo.Fields(0).Value = 0 Then:    ShowNoRecord: Exit Sub
        Select Case UCase(cboGroupBy.Text)
            Case "CUSTOMERTYPE"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CUSTOMERINFOREPORT_BY_CUSTOMERTYPE.RPT", "", CRIS_REPORT_PATH, 1
                'LogAudit "V", "CUSTOMER INFORMATION-REPORT", "CUSTOMER TYPE"
            Case "GENDER"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CUSTOMERINFOREPORT_BY_GENDER.RPT", "", CRIS_REPORT_PATH, 1
                'LogAudit "V", "CUSTOMER INFORMATION-REPORT", "GENDER"
            Case "LEADSOURCE"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CUSTOMERINFOREPORT_BY_LEADSOURCE.RPT", "", CRIS_REPORT_PATH, 1
                'LogAudit "V", "CUSTOMER INFORMATION-REPORT", "LEADSOURCE"
            Case "LOCATION"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CUSTOMERINFOREPORT_BY_LOCATION.RPT", "", CRIS_REPORT_PATH, 1
                'LogAudit "V", "CUSTOMER INFORMATION-REPORT", "LOCATION"
            Case "POSITION"
                'PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CUSTOMERINFOREPORT_BY_POSITION.RPT", "", CRIS_REPORT_PATH, 1
                LogAudit "V", "CUSTOMER INFORMATION-REPORT", "POSITION"
            Case "BIRTHDATE"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CUSTOMERINFOREPORT_BY_BIRTHDATE.RPT", "", CRIS_REPORT_PATH, 1
                'LogAudit "V", "CUSTOMER INFORMATION-REPORT", "DATE OF BIRTH"
            Case Else
                MsgBox "Please Select Proper Selection From The List", vbInformation
                cboGroupBy.SetFocus
        End Select
        
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "CUSTOMER INFORMATION", "", "", "", "CUSTOMER INFORMATION REPORT GROUP BY -" & cboGroupBy, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

        
    ElseIf optReturningCustomer.Value = True Then
        If txtNDaysVisit = "0" Then
            MsgBox "Please Select Indicate Value Greater than Zero(0) ", vbInformation
            txtNDaysVisit.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        Select Case UCase(cboDepartment)
            Case "SERVICE"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "ReturningCustomer_Service.RPT", "{@DaysVisit}<=" & txtNDaysVisit, CRIS_REPORT_PATH, 1
                LogAudit "V", "SERVICE-NON RETURING CUSTOMER", txtNDaysVisit
            Case "SALES"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "ReturningCustomer_Sales.RPT", "{@DaysVisit}<=" & txtNDaysVisit, CRIS_REPORT_PATH, 1
                LogAudit "V", "SALES-NON RETURING CUSTOMER", txtNDaysVisit
            Case Else
                MsgBox "Please Select Proper Selection From The List", vbInformation
                cboDepartment.SetFocus
                Screen.MousePointer = 0
        End Select
        
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "CUSTOMER INFORMATION", "", "", "", "RETURNING CUSTOMER " & "DEPARTMENT: " & " " & "(n)DAYS LAST VISITED" & " " & txtNDaysVisit, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
    ElseIf optFeedBack.Value = True Then
        RecordSelection = "{CSMS_REPOR.DTE_REL} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_REPOR.DTE_REL} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")"
        PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "Summary Customer Followups.RPT", RecordSelection, CRIS_REPORT_PATH, 1
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "CUSTOMER INFORMATION", "", "", "", "Summary of Customer Follow-ups" & "FROM" & FDate & " " & "TO " & TDate, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'LogAudit "V", "Summary of Customer Followups", "FOR THE RANGE:" & DateValue(FDate) & " To " & DateValue(TDate)
    ElseIf optRecommendation.Value = True Then
        PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "Summary Suggestion.RPT", RecordSelection, CRIS_REPORT_PATH, 1
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "CUSTOMER INFORMATION", "", "", "", "SUMMARY OF CUSTOMER RECOMMENDATION/SUGGESTION - " & cboSuggestions, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'LogAudit "V", "Summary Suggestion", "FOR :" & cboSuggestions
    ElseIf optCustProsp.Value = True Then
        RecordSelection = "{SP.INVOICEDDATE} >= date(" & Year(CV_1) & "," & Month(CV_1) & "," & Day(CV_1) & ") AND {SP.INVOICEDDATE} <= date(" & Year(CV_2) & "," & Month(CV_2) & "," & Day(CV_2) & ")"
        PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CustomerWVehicle.RPT", RecordSelection, CRIS_REPORT_PATH, 1
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "CUSTOMER INFORMATION", "", "", "", "CUSTOMER/PROSPECT VEHICLE REPORT -" & "FROM " & CV_1 & " " & "TO " & CV_2, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'LogAudit "V", "Customer/Prospect Report W/Vehicle Information", "FOR THE RANGE:" & DateValue(CV_1) & " To " & DateValue(CV_2)
    
    ElseIf Option1.Value = True Then
        'RecordSelection = "{SP.INVOICEDDATE} >= date(" & Year(CV_1) & "," & Month(CV_1) & "," & Day(CV_1) & ") AND {SP.INVOICEDDATE} <= date(" & Year(CV_2) & "," & Month(CV_2) & "," & Day(CV_2) & ")"
        rptCustomer_Information.WindowAllowDrillDown = True
        Select Case cboDepartment2
            Case ("SERVICE")
                lng = gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_REPAIRORDER WHERE APPOINTMENTDATE='" & DateValue(dtFortheDay) & "'").Fields(0).Value
                If lng = 0 Then
                    MsgBox "No Service Transaction Record For the " & FormatDateTime(dtFortheDay, vbLongDate), vbInformation
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                rptCustomer_Information.ReportTitle = "As of " & DateValue(Now)
                RecordSelection = "{RO.APPOINTMENTDATE}=CDATE('" & DateValue(dtFortheDay) & "')"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CustomerForTheDay_SERVICE.RPT", RecordSelection, CRIS_REPORT_PATH, 1
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "CUSTOMER INFORMATION", "", "", "", "CUSTOMER FOR THE DAY SERVICE -" & dtFortheDay, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                'LogAudit "V", "Customer For The Day Service", "FOR THE DAY:" & DateValue(dtFortheDay)
            Case ("SALES")
                lng = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE DEYT='" & DateValue(dtFortheDay) & "'").Fields(0).Value
                If lng = 0 Then
                    MsgBox "No Sales Transaction Record For the " & FormatDateTime(dtFortheDay, vbLongDate), vbInformation
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                rptCustomer_Information.ReportTitle = "As of " & DateValue(Now)
                RecordSelection = "{SO.DEYT}=CDATE('" & DateValue(dtFortheDay) & "')"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CustomerForTheDay_Sales.RPT", RecordSelection, CRIS_REPORT_PATH, 1
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "CUSTOMER INFORMATION", "", "", "", "CUSTOMER FOR THE DAY SALES -" & dtFortheDay, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                'LogAudit "V", "Customer For The Day Sales", "FOR THE DAY:" & DateValue(dtFortheDay)
            Case ("PARTS")
                lng = gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_VW_ISS_HISTORY WHERE TRANDATE='" & DateValue(dtFortheDay) & "'").Fields(0).Value
                If lng = 0 Then
                    MsgBox "No Parts Transaction Record For the " & FormatDateTime(dtFortheDay, vbLongDate), vbInformation
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                rptCustomer_Information.ReportTitle = "As of " & DateValue(Now)
                RecordSelection = "{ISS.TRANDATE}=CDATE('" & DateValue(dtFortheDay) & "')"
                PrintSQLReport rptCustomer_Information, CRIS_REPORT_PATH & "CustomerForTheDay_parts.RPT", RecordSelection, CRIS_REPORT_PATH, 1
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "CUSTOMER INFORMATION", "", "", "", "CUSTOMER FOR THE DAY PARTS -" & dtFortheDay, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                'LogAudit "V", "Customer For The Day Parts", "FOR THE DAY:" & DateValue(dtFortheDay)
            Case Else
                MsgBox "Please Select Proper Selection From the List", vbInformation
                cboDepartment2.SetFocus

        End Select
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CUSTOMER INFORMATION)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CUSTOMER INFORMATION", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitCbo
    FDate = firstDay(LOGDATE)
    TDate = Now
    CV_1 = firstDay(LOGDATE)
    CV_2 = LOGDATE
    dtFortheDay = LOGDATE
    optCustomerInfo.Value = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labDescription = ""
End Sub

Private Sub optCustomerInfo_Click()
    ShowHidePictureBox2 picCustInfo, True
End Sub

Private Sub optCustomerInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labDescription = optCustomerInfo.ToolTipText
End Sub

Private Sub optFeedBack_Click()
    ShowHidePictureBox2 picSummaryFollow, True
End Sub

Private Sub optFeedBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labDescription = optFeedBack.ToolTipText
End Sub

Private Sub optCustProsp_Click()
    ShowHidePictureBox2 picCV, True
End Sub

Private Sub optCustProsp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labDescription = optCustProsp.ToolTipText
End Sub

Private Sub Option1_Click()
    ShowHidePictureBox2 Picture1, True
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labDescription = Option1.ToolTipText
End Sub

Private Sub optRecommendation_Click()
    ShowHidePictureBox2 picSuggestion, True
End Sub

Private Sub optRecommendation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labDescription = optRecommendation.ToolTipText
End Sub

Private Sub optReturningCustomer_Click()
    ShowHidePictureBox2 picRetCust, True
End Sub

Private Sub optReturningCustomer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labDescription = optReturningCustomer.ToolTipText
End Sub

Private Sub txtNDaysVisit_GotFocus()
    If txtNDaysVisit = "0" Then txtNDaysVisit = ""
End Sub

Private Sub txtNDaysVisit_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtNDaysVisit_LostFocus()
    If IsNumeric(txtNDaysVisit) = False Then txtNDaysVisit = 0
End Sub


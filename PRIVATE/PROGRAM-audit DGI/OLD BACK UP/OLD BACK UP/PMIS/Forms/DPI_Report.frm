VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMISReports_DealerPartInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dealer Part Inquiry Report"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   Icon            =   "DPI_Report.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Filter Option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   990
      Width           =   3675
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   345
         Left            =   1890
         TabIndex        =   5
         Top             =   540
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20709377
         CurrentDate     =   39562
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   345
         Left            =   150
         TabIndex        =   6
         Top             =   540
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20709377
         CurrentDate     =   39562
      End
      Begin VB.Label Label3 
         Caption         =   "From Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "To Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1890
         TabIndex        =   7
         Top             =   270
         Width           =   765
      End
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
      Height          =   795
      Left            =   1950
      MouseIcon       =   "DPI_Report.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "DPI_Report.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   2010
      Width           =   735
   End
   Begin VB.OptionButton optTechnicalInquiryReport 
      Caption         =   "Technical Inquiry Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   90
      Value           =   -1  'True
      Width           =   3285
   End
   Begin VB.OptionButton optPriceInquiryReport 
      Caption         =   "Price Inquiry Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.OptionButton optETAInquiryReport 
      Caption         =   "ETA Inquiry Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   630
      Width           =   3975
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
      Height          =   795
      Left            =   1230
      MouseIcon       =   "DPI_Report.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "DPI_Report.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   2010
      Width           =   735
   End
   Begin Crystal.CrystalReport rptDPI_Report 
      Left            =   2970
      Top             =   2130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Daily Sales Report (As per Issuance)"
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
End
Attribute VB_Name = "frmPMISReports_DealerPartInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If optPriceInquiryReport.Value = True Then
        Screen.MousePointer = 11
        Dim rsPrice                                    As ADODB.Recordset
        Set rsPrice = New ADODB.Recordset
        Set rsPrice = gconDMIS.Execute("SELECT DPI_DATE FROM PMIS_DPIHEADER WHERE DPI_DATE BETWEEN '" & DateValue(dtpFrom.Value) & "' AND '" & DateValue(dtpTo.Value) & "'")
        If Not rsPrice.EOF And Not rsPrice.BOF Then
            rptDPI_Report.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
            rptDPI_Report.Formulas(2) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptDPI_Report.Formulas(3) = "PRINTEDBY='" & LOGCODE & "'"
            rptDPI_Report.Formulas(4) = "DATEVAR = 'FROM " & DateValue(dtpFrom.Value) & " TO " & DateValue(dtpTo.Value) & "'"
            rptDPI_Report.WindowTitle = "Price Inquiry Report"
            rptDPI_Report.ReportTitle = "Price Inquiry Report"
            Screen.MousePointer = 11
            PrintSQLReport rptDPI_Report, PMIS_REPORT_PATH & "Price Inquiry Report.rpt", "{DPI_HD.DPI_DATE} >= CDATE('" & DateValue(dtpFrom.Value) & "') AND {DPI_HD.DPI_DATE} <= CDATE('" & DateValue(dtpTo.Value) & "')", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            Screen.MousePointer = 0
            Exit Sub
        End If
    ElseIf optETAInquiryReport.Value = True Then
        Screen.MousePointer = 11
        Dim rsETA                                      As ADODB.Recordset
        Set rsETA = New ADODB.Recordset
        Set rsETA = gconDMIS.Execute("SELECT DPI_DATE FROM PMIS_DPIHEADER WHERE DPI_DATE BETWEEN '" & DateValue(dtpFrom.Value) & "' AND '" & DateValue(dtpTo.Value) & "'")
        If Not rsETA.EOF And Not rsETA.BOF Then
            rptDPI_Report.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
            rptDPI_Report.Formulas(2) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptDPI_Report.Formulas(3) = "PRINTEDBY='" & LOGCODE & "'"
            rptDPI_Report.Formulas(4) = "DATEVAR = 'FROM " & DateValue(dtpFrom.Value) & " TO " & DateValue(dtpTo.Value) & "'"
            rptDPI_Report.WindowTitle = "Estimate Time of Arrival Inquiry Report"
            rptDPI_Report.ReportTitle = "Estimate Time of Arrival Inquiry Report"
            Screen.MousePointer = 11
            PrintSQLReport rptDPI_Report, PMIS_REPORT_PATH & "ETAInquiryReport.rpt", "{DPI_HD.DPI_DATE} >= CDATE('" & DateValue(dtpFrom.Value) & "') AND {DPI_HD.DPI_DATE} <= CDATE('" & DateValue(dtpTo.Value) & "')", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            Screen.MousePointer = 0
            Exit Sub
        End If
    Else
        Screen.MousePointer = 11
        Dim rsTECH                                     As ADODB.Recordset
        Set rsTECH = New ADODB.Recordset
        Set rsTECH = gconDMIS.Execute("SELECT DPI_DATE FROM PMIS_DPIHEADER WHERE DPI_DATE BETWEEN '" & DateValue(dtpFrom.Value) & "' AND '" & DateValue(dtpTo.Value) & "'")
        
        If Not rsTECH.EOF And Not rsTECH.BOF Then
            rptDPI_Report.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
            rptDPI_Report.Formulas(2) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptDPI_Report.Formulas(3) = "PRINTEDBY='" & LOGCODE & "'"
            rptDPI_Report.Formulas(4) = "DATEVAR = 'FROM " & DateValue(dtpFrom.Value) & " TO " & DateValue(dtpTo.Value) & "'"
            rptDPI_Report.WindowTitle = "Technical Inquiry Report"
            rptDPI_Report.ReportTitle = "Technical Inquiry Report"
            Screen.MousePointer = 11
            PrintSQLReport rptDPI_Report, PMIS_REPORT_PATH & "Technical Inquiry Report.rpt", "{DPI_HD.DPI_DATE} >= CDATE('" & DateValue(dtpFrom.Value) & "') AND {DPI_HD.DPI_DATE} <= CDATE('" & DateValue(dtpTo.Value) & "')", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    dtpFrom = Format(firstDay(LOGDATE), "DD-MMM-YY")
    dtpTo = Format(LOGDATE, "DD-MMM-YY")
    Screen.MousePointer = 0
End Sub





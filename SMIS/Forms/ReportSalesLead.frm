VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_SalesLead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lost Sales Monitoring"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportSalesLead.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   5145
   Begin VB.CheckBox chkSummary 
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2790
      TabIndex        =   8
      Top             =   1530
      Width           =   2295
   End
   Begin VB.ComboBox cboSAE 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   360
      ItemData        =   "ReportSalesLead.frx":0E42
      Left            =   60
      List            =   "ReportSalesLead.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   990
      Width           =   5025
   End
   Begin VB.ComboBox cboLeadBy 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   405
      ItemData        =   "ReportSalesLead.frx":0E46
      Left            =   90
      List            =   "ReportSalesLead.frx":0E48
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   5025
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   45
      ScaleHeight     =   405
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   1440
      Width           =   2715
      Begin VB.OptionButton optRange 
         Caption         =   "Ranged"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton optMonthly 
         Caption         =   "Monthly"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optYearly 
         Caption         =   "Yearly"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.TextBox txtLeadDays 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4110
      TabIndex        =   22
      Text            =   "0"
      Top             =   2340
      Width           =   945
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
      Left            =   2430
      MouseIcon       =   "ReportSalesLead.frx":0E4A
      MousePointer    =   99  'Custom
      Picture         =   "ReportSalesLead.frx":0F9C
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Close Window"
      Top             =   2730
      Width           =   885
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   570
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "List of Registrations"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
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
      Left            =   1560
      MouseIcon       =   "ReportSalesLead.frx":13E7
      MousePointer    =   99  'Custom
      Picture         =   "ReportSalesLead.frx":1539
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Print Report"
      Top             =   2730
      Width           =   885
   End
   Begin VB.PictureBox picYearly 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      ScaleHeight     =   495
      ScaleWidth      =   5025
      TabIndex        =   9
      Top             =   1830
      Width           =   5025
      Begin VB.ComboBox cboByYear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   405
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   30
         Width           =   3315
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   990
         TabIndex        =   10
         Top             =   60
         Width           =   510
      End
   End
   Begin VB.PictureBox picRange 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      ScaleHeight     =   495
      ScaleWidth      =   5025
      TabIndex        =   17
      Top             =   1830
      Width           =   5025
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Top             =   45
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
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
         Format          =   56950785
         CurrentDate     =   39427
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   375
         Left            =   3060
         TabIndex        =   21
         Top             =   45
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
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
         Format          =   56950785
         CurrentDate     =   39427
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   2370
         TabIndex        =   20
         Top             =   90
         Width           =   600
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   60
         TabIndex        =   18
         Top             =   90
         Width           =   600
      End
   End
   Begin VB.PictureBox picMonthly 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      ScaleHeight     =   495
      ScaleWidth      =   5025
      TabIndex        =   12
      Top             =   1830
      Width           =   5025
      Begin VB.ComboBox cboMonth 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   405
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   30
         Width           =   2295
      End
      Begin VB.TextBox txtYear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   435
         Left            =   4050
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "9999"
         Top             =   30
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   60
         Width           =   600
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   3420
         TabIndex        =   15
         Top             =   105
         Width           =   510
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Account Executive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   3030
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Lost Sales By"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   1590
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Lead Days Limit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2460
      TabIndex        =   23
      Top             =   2400
      Width           =   2610
   End
End
Attribute VB_Name = "frmSMIS_Report_SalesLead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLostSales                                                       As ADODB.Recordset

Sub LostSalesByModel()
    On Error GoTo Errorcode:

    Dim FILTER                                                        As String
    On Error GoTo Errorcode:
    Set rsLostSales = New ADODB.Recordset
    If optMonthly.Value = True Then
        rsLostSales.Open "SELECT * FROM CRIS_PROSPECTS WHERE MONTH(LOGINITIALINQUIRY) = " & What_month(cboMonth) & " AND YEAR(LOGINITIALINQUIRY) = " & txtYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf optYearly.Value = True Then
        rsLostSales.Open "SELECT * FROM CRIS_PROSPECTS WHERE YEAR(LOGINITIALINQUIRY) = " & txtYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf optRange.Value = True Then
        rsLostSales.Open "SELECT * FROM CRIS_PROSPECTS WHERE LOGINITIALINQUIRY BETWEEN  '" & dtFrom & "' AND '" & dtTo & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If

    If Not rsLostSales.BOF And Not rsLostSales.EOF Then
        Screen.MousePointer = 11
        rptGenREP.WindowTitle = "Lost Sales By " & cboLeadBy
        rptGenREP.ReportTitle = UCase("Lost Sales By " & cboLeadBy)

        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        If optMonthly.Value = True Then
            rptGenREP.Formulas(2) = "DateRange = '" & cboMonth & " " & txtYear & "'"

            'FILTER = "month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & " and year({CP.LOGINITIALINQUIRY}) = " & txtYear.Text
            FILTER = "(YEAR({CP.LOGCLOSINGDATE})=" & txtYear.Text & " AND MONTH({CP.LOGCLOSINGDATE})=" & What_month(cboMonth) & " Or Year({CP.LOGINITIALINQUIRY})= " & txtYear & " And Month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & ")"

        ElseIf optRange.Value = True Then
            'FILTER = "({CP.LOGINITIALINQUIRY} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {CP.LOGINITIALINQUIRY} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & "))"
            FILTER = "({CP.LOGCLOSINGDATE}>=Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ")  Or {CP.LOGINITIALINQUIRY}>=Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & "))"
            rptGenREP.Formulas(2) = "DateRange = 'From " & dtFrom & " To  " & dtTo & "'"
        ElseIf optYearly.Value = True Then
            'FILTER = " YEAR({CP.LOGINITIALINQUIRY}) = " & cboByYear

            FILTER = "(YEAR({CP.LOGCLOSINGDATE})>=" & txtYear.Text & " AND MONTH({CP.LOGCLOSINGDATE})>=" & What_month(dtFrom) & " AND DAY({CP.LOGCLOSINGDATE})>=" & Day(dtFrom) & "  Or Year({CP.LOGINITIALINQUIRY})>= " & txtYear & " And Month({CP.LOGINITIALINQUIRY}) >=" & What_month(dtTo) & " AND DAY({CP.LOGINITIALINQUIRY}) >= " & Day(dtTo) & ")"
            rptGenREP.Formulas(2) = "DateRange = ' For The Year " & cboByYear & "'"
        End If
        rptGenREP.Formulas(3) = "LOSTSALESPARAM=" & txtLeadDays
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "Monitoring\LostSalesByModel.rpt", FILTER, DMIS_REPORT_Connection, 1
        If chkSummary.Value = 1 Then

            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "Monitoring\LostSalesByModel-Summary.rpt", FILTER, DMIS_REPORT_Connection, 1
        End If
        Call SaveSetting("SMIS", "REPORTS", "LEADDAYS", txtLeadDays)

        Screen.MousePointer = 0

        'PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "saleslead.rpt", "month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & " and year({CP.LOGINITIALINQUIRY}) = " & txtYear.Text, DMIS_REPORT_Connection, 1
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonth.Text
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub LostSalesByLeadSource()
    On Error GoTo Errorcode:
    Dim FILTER                                                        As String
    On Error GoTo Errorcode:
    Set rsLostSales = New ADODB.Recordset
    If optMonthly.Value = True Then
        rsLostSales.Open "SELECT * FROM CRIS_PROSPECTS WHERE MONTH(LOGINITIALINQUIRY) = " & What_month(cboMonth) & " AND YEAR(LOGINITIALINQUIRY) = " & txtYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf optYearly.Value = True Then
        rsLostSales.Open "SELECT * FROM CRIS_PROSPECTS WHERE YEAR(LOGINITIALINQUIRY) = " & txtYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf optRange.Value = True Then
        rsLostSales.Open "SELECT * FROM CRIS_PROSPECTS WHERE LOGINITIALINQUIRY BETWEEN  '" & dtFrom & "' AND '" & dtTo & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If

    If Not rsLostSales.BOF And Not rsLostSales.EOF Then
        Screen.MousePointer = 11
        rptGenREP.WindowTitle = "Lost Sales By " & cboLeadBy
        rptGenREP.ReportTitle = UCase("Lost Sales By " & cboLeadBy)
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        If optMonthly.Value = True Then
            rptGenREP.Formulas(2) = "DateRange = '" & cboMonth & " " & txtYear & "'"
            'FILTER = "month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & " and year({CP.LOGINITIALINQUIRY}) = " & txtYear.Text
            FILTER = "(YEAR({CP.LOGCLOSINGDATE})=" & txtYear.Text & " AND MONTH({CP.LOGCLOSINGDATE})=" & What_month(cboMonth) & " Or Year({CP.LOGINITIALINQUIRY})= " & txtYear & " And Month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & ")"
        ElseIf optRange.Value = True Then
            'FILTER = " ({CP.LOGINITIALINQUIRY} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {CP.LOGINITIALINQUIRY} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) "
            FILTER = "({CP.LOGCLOSINGDATE}>=Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ")  Or {CP.LOGINITIALINQUIRY}>=Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & "))"
            rptGenREP.Formulas(2) = "DateRange = 'From " & dtFrom & " To  " & dtTo & "'"
        ElseIf optYearly.Value = True Then
            'FILTER = " YEAR({CP.LOGINITIALINQUIRY}) = " & cboByYear
            FILTER = "(YEAR({CP.LOGCLOSINGDATE})>=" & cboByYear.Text & " AND MONTH({CP.LOGCLOSINGDATE})>=" & What_month(dtFrom) & " AND DAY({CP.LOGCLOSINGDATE})>=" & Day(dtFrom) & "  Or Year({CP.LOGINITIALINQUIRY})>= " & txtYear & " And Month({CP.LOGINITIALINQUIRY}) >=" & What_month(dtTo) & " AND DAY({CP.LOGINITIALINQUIRY}) >= " & Day(dtTo) & ")"
            rptGenREP.Formulas(2) = "DateRange = ' For The Year " & cboByYear & "'"
        End If
        rptGenREP.Formulas(3) = "LOSTSALESPARAM=" & txtLeadDays
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "Monitoring\LostSalesByLeadSource.rpt", FILTER, DMIS_REPORT_Connection, 1
        If chkSummary.Value = 1 Then
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "Monitoring\LostSalesByLeadSource-Summary.rpt", FILTER, DMIS_REPORT_Connection, 1
        End If
        Call SaveSetting("SMIS", "REPORTS", "LEADDAYS", txtLeadDays)

        Screen.MousePointer = 0

        'PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "saleslead.rpt", "month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & " and year({CP.LOGINITIALINQUIRY}) = " & txtYear.Text, DMIS_REPORT_Connection, 1
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonth.Text
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub LostSalesForSC()
    Dim FILTER                                                        As String
    On Error GoTo Errorcode:
    Set rsLostSales = New ADODB.Recordset
    If optMonthly.Value = True Then
        rsLostSales.Open "select * from CRIS_PROSPECTS WHERE sae='" & cboSAE.Text & "' and  month(LOGINITIALINQUIRY) = " & What_month(cboMonth) & " and year(LOGINITIALINQUIRY) = " & txtYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf optYearly.Value = True Then
        rsLostSales.Open "select * from CRIS_PROSPECTS WHERE sae='" & cboSAE.Text & "' and year(LOGINITIALINQUIRY) = " & txtYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf optRange.Value = True Then
        rsLostSales.Open "select * from CRIS_PROSPECTS WHERE sae='" & cboSAE.Text & "' and LOGINITIALINQUIRY BETWEEN  '" & dtFrom & "' and '" & dtTo & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If

    If Not rsLostSales.BOF And Not rsLostSales.EOF Then
        Screen.MousePointer = 11
        rptGenREP.WindowTitle = "Lost Sales By " & cboLeadBy
        rptGenREP.ReportTitle = UCase("Lost Sales By " & cboLeadBy)
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        If optMonthly.Value = True Then
            rptGenREP.Formulas(2) = "DateRange = '" & cboMonth & " " & txtYear & "'"

            'FILTER = "{CP.SAE}='" & cboSAE & "' AND month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & " and year({CP.LOGINITIALINQUIRY}) = " & txtYear.Text

            FILTER = "{CP.SAE}='" & cboSAE & "' AND (YEAR({CP.LOGCLOSINGDATE})=" & txtYear & " AND MONTH({CP.LOGCLOSINGDATE})=" & What_month(cboMonth) & " OR YEAR({CP.LOGINITIALINQUIRY})=" & txtYear & " AND MONTH({CP.LOGINITIALINQUIRY})=" & What_month(cboMonth) & ")"

        ElseIf optRange.Value = True Then
            'FILTER = "{CP.SAE}='" & cboSAE & "' AND ({CP.LOGINITIALINQUIRY} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {CP.LOGINITIALINQUIRY} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) "
            FILTER = "{CP.SAE}='" & cboSAE & "' AND ({CP.LOGCLOSINGDATE}>=Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ")  Or {CP.LOGINITIALINQUIRY}>=Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & "))"
            rptGenREP.Formulas(2) = "DateRange = 'From " & dtFrom & " To  " & dtTo & "'"
        ElseIf optYearly.Value = True Then
            'FILTER = "{CP.SAE}='" & cboSAE & "' AND YEAR({CP.LOGINITIALINQUIRY}) = " & cboByYear
            FILTER = "{CP.SAE}='" & cboSAE & "' AND YEAR({CP.LOGCLOSINGDATE}) = " & cboByYear
            rptGenREP.Formulas(2) = "DateRange = ' For The Year " & cboByYear & "'"
        End If
        rptGenREP.Formulas(3) = "LOSTSALESPARAM=" & txtLeadDays
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "Monitoring\LostSalesByAgent.rpt", FILTER, DMIS_REPORT_Connection, 1
        If chkSummary.Value = 1 Then
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "Monitoring\LostSalesByAgent-Summary.rpt", FILTER, DMIS_REPORT_Connection, 1
        End If
        Call SaveSetting("SMIS", "REPORTS", "LEADDAYS", txtLeadDays)

        Screen.MousePointer = 0

        'PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "saleslead.rpt", "month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & " and year({CP.LOGINITIALINQUIRY}) = " & txtYear.Text, DMIS_REPORT_Connection, 1
    Else
        MsgSpeechBox "No Record for the given selection "
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub LostSalesForAllSC()
    'Programmed Modified By:Ryan Culaway
    'Date Modified: JULY 9 2008

    Dim FILTER                                                        As String
    Set rsLostSales = New ADODB.Recordset
    If optMonthly.Value = True Then
        rsLostSales.Open "select * from CRIS_PROSPECTS WHERE MONTH(LOGINITIALINQUIRY) = " & What_month(cboMonth) & " and year(LOGINITIALINQUIRY) = " & txtYear.Text & " AND STATUS='L'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf optYearly.Value = True Then
        rsLostSales.Open "select * from CRIS_PROSPECTS WHERE year(LOGINITIALINQUIRY) = " & txtYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf optRange.Value = True Then
        rsLostSales.Open "select * from CRIS_PROSPECTS WHERE LOGINITIALINQUIRY BETWEEN  '" & dtFrom & "' and '" & dtTo & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsLostSales.BOF And Not rsLostSales.EOF Then
        Screen.MousePointer = 11
        rptGenREP.WindowTitle = "Lost Sales By " & cboLeadBy
        rptGenREP.ReportTitle = UCase("Lost Sales By " & cboLeadBy)
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        If optMonthly.Value = True Then
            rptGenREP.Formulas(2) = "DateRange = '" & cboMonth & " " & txtYear & "'"
            'FILTER = "Month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & " and Year({CP.LOGINITIALINQUIRY}) = " & txtYear.Text

            'JULY 9 2008
            FILTER = "(YEAR({CP.LOGCLOSINGDATE})=" & txtYear.Text & " AND MONTH({CP.LOGCLOSINGDATE})=" & What_month(cboMonth) & " Or Year({CP.LOGINITIALINQUIRY})= " & txtYear & " And Month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & ")"


        ElseIf optRange.Value = True Then

            'FILTER = "YEAR({CP.LOGCLOSINGDATE})>=" & Year(dtFrom) & " AND MONTH({CP.LOGCLOSINGDATE})>=" & Month(dtFrom) & " AND DAY({CP.LOGCLOSINGDATE})>=" & Day(dtFrom) & " AND YEAR({CP.LOGCLOSINGDATE})<=" & Year(dtTo) & " AND MONTH({CP.LOGCLOSINGDATE})<=" & Month(dtTo) & " AND DAY({CP.LOGCLOSINGDATE})<=" & Day(dtTo) & ")) OR ((Year({CP.LOGINITIALINQUIRY})>= " & Year(dtFrom) & " And Month({CP.LOGINITIALINQUIRY}) >= " & Month(dtFrom) & " AND DAY({CP.LOGINITIALINQUIRY}) >= " & Day(dtFrom) & " AND Year({CP.LOGINITIALINQUIRY})<= " & Year(dtTo) & " And Month({CP.LOGINITIALINQUIRY}) <= " & Month(dtTo) & " AND DAY({CP.LOGINITIALINQUIRY}) <= " & Day(dtTo) & "))"


            'FILTER = "(YEAR({CP.LOGCLOSINGDATE}))>=" & Year(dtFrom) & " AND Month({CP.LOGCLOSINGDATE})>=" & Month(dtFrom) & " AND DAY({CP.LOGCLOSINGDATE})>=" & Day(dtFrom) & " AND (YEAR({CP.LOGCLOSINGDATE}))<=" & Year(dtTo) & " AND Month({CP.LOGCLOSINGDATE})<=" & Month(dtTo) & " AND DAY({CP.LOGCLOSINGDATE})<=" & Day(dtTo) & " OR (YEAR({CP.LOGINITIALINQUIRY}))>=" & Year(dtFrom) & " AND Month({CP.LOGINITIALINQUIRY})>=" & Month(dtFrom) & " AND DAY({CP.LOGINITIALINQUIRY})>=" & Day(dtFrom) & " AND (YEAR({CP.LOGINITIALINQUIRY}))<=" & Year(dtTo) & " AND Month({CP.LOGINITIALINQUIRY})<=" & Month(dtTo) & " AND DAY({CP.LOGINITIALINQUIRY})<=" & Day(dtTo) & ""
            ''''DFD
            'FILTER = "(YEAR({CP.LOGCLOSINGDATE})>=" & txtYear.Text & " AND MONTH({CP.LOGCLOSINGDATE})>=" & What_month(dtFrom) & " AND DAY({CP.LOGCLOSINGDATE})>=" & Day(dtFrom) & "  Or Year({CP.LOGINITIALINQUIRY})>= " & txtYear & " And Month({CP.LOGINITIALINQUIRY}) >=" & What_month(dtTo) & " AND DAY({CP.LOGINITIALINQUIRY}) >= " & Day(dtTo) & ")"

            'FILTER = "{CP.LOGCLOSINGDATE}>=Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ")  Or {CP.LOGINITIALINQUIRY}<=Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")"

            FILTER = "({CP.LOGCLOSINGDATE}>=Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ")  Or {CP.LOGINITIALINQUIRY}>=Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & "))"
            rptGenREP.Formulas(2) = "DateRange = 'From " & dtFrom & " To  " & dtTo & "'"

        ElseIf optYearly.Value = True Then
            'FILTER = "YEAR({CP.LOGINITIALINQUIRY}) = " & cboByYear

            'JULY 9 2008
            FILTER = "(YEAR({CP.LOGCLOSINGDATE})=" & txtYear & " OR YEAR({CP.LOGINITIALINQUIRY})=" & txtYear & ")"

            rptGenREP.Formulas(2) = "DateRange = ' For The Year " & cboByYear & "'"
        End If
        rptGenREP.Formulas(3) = "LOSTSALESPARAM=" & txtLeadDays
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "Monitoring\LostSalesByAgent.rpt", FILTER, DMIS_REPORT_Connection, 1
        If chkSummary.Value = 1 Then
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "Monitoring\LostSalesByAgent-Summary.rpt", FILTER, DMIS_REPORT_Connection, 1
        End If
        Call SaveSetting("SMIS", "REPORTS", "LEADDAYS", txtLeadDays)
        Screen.MousePointer = 0
        'PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "saleslead.rpt", "month({CP.LOGINITIALINQUIRY}) = " & What_month(cboMonth) & " and year({CP.LOGINITIALINQUIRY}) = " & txtYear.Text, DMIS_REPORT_Connection, 1
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonth.Text
    End If
End Sub

Private Sub cboLeadBy_Change()
    If cboLeadBy.ListIndex = 0 Then
        cboSAE.Enabled = True
    Else
        cboSAE.Enabled = False
    End If

End Sub

Private Sub cboLeadBy_Click()
    cboLeadBy_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode
    If cboLeadBy = "Monitoring All SC Performance" Then
        LostSalesForAllSC
        
    ElseIf cboLeadBy = "Monitoring Per SC Performance" Then
        If cboSAE.ListIndex = -1 Then
            MsgBox "Please Select Your Sales Account Executive Name", vbInformation
            cboSAE.SetFocus
            Exit Sub
        End If
        LostSalesForSC
    ElseIf cboLeadBy = "Model" Then
        LostSalesByModel
    ElseIf cboLeadBy = "Lead Source" Then
        LostSalesByLeadSource
    End If
    
    If optMonthly.Value = True And cboLeadBy.Text = "Monitoring Per SC Performance" Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "LOST SALES MONITORING", "", "", "", cboLeadBy & " " & cboSAE & " " & cboMonth & " " & txtYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf optRange.Value = True Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "LOST SALES MONITORING", "", "", "", cboLeadBy & " " & " " & "FROM " & dtFrom & " " & "TO " & dtFrom, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Else
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "LOST SALES MONITORING", "", "", "", cboLeadBy & " " & cboByYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End If
    Exit Sub

Errorcode:
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LOST SALES MONITORING)"
            Call frmALL_AuditInquiry.DisplayHistory("", "LOST SALES MONITORING", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    With cboLeadBy
        .AddItem "Monitoring Per SC Performance"
        .AddItem "Monitoring All SC Performance"
        .AddItem "Model"
        .AddItem "Lead Source"
        .ListIndex = 0
    End With
    fillcbomonth cboMonth
    fillcbomoreyear cboByYear
    Combo_Loadval cboSAE, gconDMIS.Execute("Select DISTINCT sae from cris_prospects")
    dtFrom = DateValue(firstDay(LOGDATE))
    dtTo = Date
    txtLeadDays = GetSetting("SMIS", "REPORTS", "LEADDAYS", 0)
    cboMonth.Text = The_month(Month(LOGDATE))
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

Private Sub optMonthly_Click()
    picYearly.Visible = False
    picMonthly.Visible = True
    picRange.Visible = False

End Sub

Private Sub optRange_Click()
    picYearly.Visible = False
    picMonthly.Visible = False
    picRange.Visible = True

End Sub

Private Sub optYearly_Click()
    picYearly.Visible = True
    picMonthly.Visible = False
    picRange.Visible = False
End Sub

Private Sub txtLeadDays_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub


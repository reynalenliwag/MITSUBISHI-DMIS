VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_GrossProfit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Gross Profit Report"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
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
   Icon            =   "frmSMIS_Report_GrossProfit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   4665
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   4245
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
      Height          =   765
      Left            =   2220
      MouseIcon       =   "frmSMIS_Report_GrossProfit.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_GrossProfit.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   1950
      Width           =   915
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
      Height          =   765
      Left            =   1320
      MouseIcon       =   "frmSMIS_Report_GrossProfit.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_GrossProfit.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   1950
      Width           =   915
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   4050
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Gross Profit Rate Report"
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
   Begin VB.PictureBox picMonthly 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   210
      ScaleHeight     =   1215
      ScaleWidth      =   4245
      TabIndex        =   9
      Top             =   720
      Width           =   4245
      Begin VB.ComboBox cboMonthlyMonth2 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   1965
      End
      Begin VB.ComboBox cboMonthlyMonth1 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   75
         Width           =   1965
      End
      Begin VB.ComboBox cboMonthlyYear 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   870
         Width           =   1965
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Height          =   255
         Left            =   30
         TabIndex        =   28
         Top             =   870
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   510
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.PictureBox picDaily 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   210
      ScaleHeight     =   1215
      ScaleWidth      =   4245
      TabIndex        =   3
      Top             =   720
      Width           =   4245
      Begin MSComCtl2.DTPicker dtDailyFrom 
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Top             =   75
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   56557569
         CurrentDate     =   39427
      End
      Begin MSComCtl2.DTPicker dtDailyTo 
         Height          =   315
         Left            =   900
         TabIndex        =   7
         Top             =   480
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "MMMM"
         Format          =   56557569
         CurrentDate     =   39427
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.PictureBox picYearly 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   210
      ScaleHeight     =   1215
      ScaleWidth      =   4245
      TabIndex        =   22
      Top             =   720
      Width           =   4245
      Begin VB.ComboBox cboYearly2 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   480
         Width           =   1965
      End
      Begin VB.ComboBox cboYearly1 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   75
         Width           =   1965
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.PictureBox picWeekly 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   210
      ScaleHeight     =   1215
      ScaleWidth      =   4245
      TabIndex        =   8
      Top             =   780
      Width           =   4245
      Begin MSComCtl2.DTPicker dtWeeklyFrom 
         Height          =   315
         Left            =   900
         TabIndex        =   15
         Top             =   75
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Format          =   56557569
         CurrentDate     =   39427
      End
      Begin MSComCtl2.DTPicker dtWeeklyTo 
         Height          =   315
         Left            =   900
         TabIndex        =   16
         Top             =   480
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMMM"
         Format          =   56557569
         CurrentDate     =   39427
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   90
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   480
         Width           =   825
      End
   End
   Begin VB.PictureBox picQuarterly 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   210
      ScaleHeight     =   1215
      ScaleWidth      =   4245
      TabIndex        =   19
      Top             =   720
      Width           =   4245
      Begin VB.ComboBox cboQuarter 
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
         ForeColor       =   &H00800000&
         Height          =   345
         ItemData        =   "frmSMIS_Report_GrossProfit.frx":19D0
         Left            =   900
         List            =   "frmSMIS_Report_GrossProfit.frx":19D2
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   75
         Width           =   1965
      End
      Begin VB.ComboBox cboQuarterlyYear 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   480
         Width           =   1965
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   510
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quarter"
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
         Height          =   255
         Left            =   -30
         TabIndex        =   29
         Top             =   90
         Width           =   825
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Report"
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
      Height          =   255
      Left            =   210
      TabIndex        =   27
      Top             =   60
      Width           =   6135
   End
End
Attribute VB_Name = "frmSMIS_Report_GrossProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset

Sub PrintDailyGross()
    '   On Error GoTo Errorcode:
    Dim FILTER
    Dim DateFrom                                                      As Date
    Dim DateTo                                                        As Date

    Set rsPurchAgree = New ADODB.Recordset

    'rsPurchAgree.Open "SELECT * FROM SMIS_PURCHAGREE WHERE DATERELEASED BETWEEN " & DateValue(dtDailyFrom.Value) & " AND " & DateValue(dtDailyTo.Value) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly

    rsPurchAgree.Open "SELECT * FROM SMIS_PURCHAGREE WHERE Year(DATERELEASED)>=" & Year(dtDailyFrom.Value) & " AND MOnth(DATERELEASED)>=" & Month(dtDailyFrom.Value) & " AND DAY(DATERELEASED)>=" & Day(dtDailyFrom.Value) & " AND Year(DATERELEASED)<=" & Year(dtDailyTo) & " AND MOnth(DATERELEASED)<=" & Month(dtDailyTo) & " AND DAY(DATERELEASED)<=" & Day(dtDailyTo) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
        Screen.MousePointer = 11
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"


        If DateValue(dtDailyFrom.Value) = DateValue(dtDailyTo.Value) Then
            'FILTER = "Year({purchagree.datereleased})= " & Year(dtDailyFrom.Value) & " AND Month({purchagree.datereleased})= " & Month(dtDailyFrom.Value) & " AND Day({purchagree.datereleased})= " & Day(dtDailyFrom.Value) & ")"
            'Year(DATERELEASED)>=2008  AND MOnth(DATERELEASED)>=6 AND DAY(DATERELEASED)>=28 AND  Year(DATERELEASED)<=2008  AND MOnth(DATERELEASED)<=6 AND DAY(DATERELEASED)<=28
            '
            FILTER = "Year({purchagree.datereleased})>= " & Year(dtDailyFrom.Value) & " AND Month({purchagree.datereleased})>= " & Month(dtDailyFrom.Value) & " AND Day({purchagree.datereleased})>= " & Day(dtDailyFrom.Value) & " AND Year({purchagree.datereleased})<= " & Year(dtDailyTo.Value) & "  AND Month({purchagree.datereleased})<= " & Month(dtDailyTo.Value) & " AND DAY({purchagree.datereleased})<= " & Day(dtDailyTo.Value) & ""

        Else
            DateFrom = DateSerial(Year(dtDailyFrom), Month(dtDailyFrom), Day(dtDailyFrom))
            DateTo = DateSerial(Year(dtDailyTo), Month(dtDailyTo), Day(dtDailyTo))
            '"{REPOR.DTE_COMP} >= DATE(" & Year(txtfrom.Value) & "," & Month(txtfrom.Value) & "," & Day(txtfrom.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtto.Value) & "," & Month(txtto.Value) & "," & Day(txtto.Value) & ")", CSMS_REPORT_CONNECTION, 1
            FILTER = "(Date(Year({purchagree.datereleased}),month({purchagree.datereleased}),day({purchagree.datereleased})) >= Date(" & Year(dtDailyFrom) & "," & Month(dtDailyFrom) & "," & Day(dtDailyFrom) & ") and Date(Year({purchagree.datereleased}),month({purchagree.datereleased}),day({purchagree.datereleased})) <= Date(" & Year(dtDailyTo) & "," & Month(dtDailyTo) & "," & Day(dtDailyTo) & "))"
        End If

        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossDaily.rpt", FILTER, DMIS_REPORT_Connection, 1
        'FILTER = "MONTH({purchagree.datereleased}) >= " & What_month(cboMonthlyMonth1) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonthlyMonth2) & " AND Day({purchagree.datereleased}) <= " & Day(cboMonthlyMonth2) & " AND year({purchagree.datereleased}) = " & cboMonthlyYear.Text
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record in the Range"
    End If
    Exit Sub
    Exit Sub
    'Errorcode:
    '    ShowVBError
End Sub

Sub PrintMonthlyGross()
    On Error GoTo ErrorCode:
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE month(datereleased) >= " & What_month(cboMonthlyMonth1) & " AND month(datereleased) <=" & What_month(cboMonthlyMonth2), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossMonthly.rpt", "month({purchagree.datereleased}) >= " & What_month(cboMonthlyMonth1) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonthlyMonth2) & " AND year({purchagree.datereleased}) = " & cboMonthlyYear.Text, DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonthlyMonth1.Text
    End If
    Exit Sub
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub PrintQuarterlyGross()
    'PROGRAMMED BY: RYAN CULAWAY
    'DATE MODIFIED: JULY 7 2008

    On Error GoTo ErrorCode:

    If cboQuarter.Text = "ALL" Then
        Set rsPurchAgree = New ADODB.Recordset
        rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE Year(datereleased) = " & cboQuarterlyYear & "", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then

            Screen.MousePointer = 11
            rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossQuarter.rpt", "YEAR({PURCHAGREE.DATERELEASED})=" & cboQuarterlyYear & "", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            Exit Sub
        Else
            MsgSpeechBox "No Record for the Year of " & cboQuarterlyYear.Text
        End If

    Else
        If cboQuarter.Text = "QUARTER I" Then
            Set rsPurchAgree = New ADODB.Recordset
            rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE month(datereleased) >=1 AND month(datereleased) <=3 AND YEAR(datereleased) = " & cboQuarterlyYear.Text & "", gconDMIS

            If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
                Screen.MousePointer = 11
                rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

                PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossQuarter.rpt", "MONTH({purchagree.datereleased}) >=1 AND MONTH({purchagree.datereleased}) <=3  AND YEAR({purchagree.datereleased}) = " & cboQuarterlyYear.Text, DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for the Quanter of " & cboQuarter.Text
                Exit Sub
            End If
        ElseIf cboQuarter.Text = "QUARTER II" Then
            Set rsPurchAgree = New ADODB.Recordset
            rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE month(datereleased) >=4 AND month(datereleased) <=6 AND YEAR(datereleased) = " & cboQuarterlyYear.Text & "", gconDMIS

            If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
                Screen.MousePointer = 11
                rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossQuarter.rpt", "MONTH({purchagree.datereleased}) >=4 AND MONTH({purchagree.datereleased}) <=6  AND YEAR({purchagree.datereleased}) = " & cboQuarterlyYear.Text, DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for the Quanter of " & cboQuarter.Text
                Exit Sub
            End If
        ElseIf cboQuarter.Text = "QUARTER III" Then
            Set rsPurchAgree = New ADODB.Recordset
            rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE month(datereleased) >=7 AND month(datereleased) <=9 AND YEAR(datereleased) = " & cboQuarterlyYear.Text & "", gconDMIS

            If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
                Screen.MousePointer = 11
                rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossQuarter.rpt", "MONTH({purchagree.datereleased}) >=7 AND MONTH({purchagree.datereleased}) <=9  AND YEAR({purchagree.datereleased}) = " & cboQuarterlyYear.Text, DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for the Quanter of " & cboQuarter.Text
                Exit Sub
            End If
        ElseIf cboQuarter.Text = "QUARTER IV" Then
            Set rsPurchAgree = New ADODB.Recordset
            rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE month(datereleased) >=10 AND month(datereleased) <=12 AND YEAR(datereleased) = " & cboQuarterlyYear.Text & "", gconDMIS

            If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
                Screen.MousePointer = 11
                rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossQuarter.rpt", "MONTH({purchagree.datereleased}) >=10 AND MONTH({purchagree.datereleased}) <=12  AND YEAR({purchagree.datereleased}) = " & cboQuarterlyYear.Text, DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for the Quanter of " & cboQuarter.Text
                Exit Sub
            End If

        End If
    End If
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Sub PrintWeeklyGross()
    On Error GoTo ErrorCode:
    Dim FILTER
    Set rsPurchAgree = New ADODB.Recordset

    rsPurchAgree.Open "SELECT * FROM SMIS_PURCHAGREE WHERE  DATERELEASED BETWEEN  '" & DateValue(dtWeeklyFrom) & "' AND '" & DateValue(dtWeeklyTo) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        'FILTER = "({purchagree.datereleased} >= date(" & Year(dtWeeklyFrom) & "," & Month(dtWeeklyFrom) & "," & Day(dtWeeklyFrom) & ") AND {purchagree.datereleased} <= date(" & Year(dtWeeklyTo) & "," & Month(dtWeeklyTo) & "," & Day(dtWeeklyTo) & ")) "
        'UPDATE BY: JUN
        'DATE UPDATE: 08/02/2008
        FILTER = "(Date(Year({purchagree.datereleased}),month({purchagree.datereleased}),day({purchagree.datereleased})) >= date(" & Year(dtWeeklyFrom) & "," & Month(dtWeeklyFrom) & "," & Day(dtWeeklyFrom) & ") AND Date(Year({purchagree.datereleased}),month({purchagree.datereleased}),day({purchagree.datereleased})) <= date(" & Year(dtWeeklyTo) & "," & Month(dtWeeklyTo) & "," & Day(dtWeeklyTo) & ")) "
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossWeekly.rpt", FILTER, DMIS_REPORT_Connection, 1
        '"MONTH({purchagree.datereleased}) >= " & What_month(cboMonthlyMonth1) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonthlyMonth2) & " AND Day({purchagree.datereleased}) <= " & Day(cboMonthlyMonth2) & " AND year({purchagree.datereleased}) = " & cboMonthlyYear.Text
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record in the Range"
    End If
    Exit Sub
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub PrintYearlyGross()
    On Error GoTo ErrorCode:
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE YEAR(Datereleased) >= " & cboYearly1 & " AND YEAR(datereleased) <=" & cboYearly2, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "VS/GrossYearly.rpt", "YEAR({purchagree.datereleased}) >= " & cboYearly1 & " AND YEAR({purchagree.datereleased}) <= " & cboYearly2, DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the in the range of " & cboYearly1.Text & " " & cboYearly2.Text
    End If
    Exit Sub
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    If Combo1.Text = "DAILY VEHICLES GROSS PROFIT REPORT" Then
        PrintDailyGross
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call NEW_LogAudit("V", "VEHICLE SALES", "", "", "", Combo1 & " " & "FROM " & dtDailyFrom & " " & "TO " & dtDailyTo, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'LogAudit "V", "DAILY VEHICLES GROSS PROFIT REPORT", "DATE " & dtDailyFrom & " " & dtDailyTo
    ElseIf Combo1.Text = "WEEKLY VEHICLES GROSS PROFIT REPORT" Then
        PrintWeeklyGross
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call NEW_LogAudit("V", "VEHICLE SALES", "", "", "", Combo1 & " " & "FROM " & dtWeeklyFrom & " " & "TO " & dtWeeklyTo, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'LogAudit "V", "WEEKLY VEHICLES GROSS PROFIT REPORT", "DATE " & dtDailyFrom & " " & dtDailyTo
    ElseIf Combo1.Text = "MONTHLY VEHICLES GROSS PROFIT REPORT" Then
        PrintMonthlyGross
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call NEW_LogAudit("V", "VEHICLE SALES", "", "", "", Combo1 & " " & "FROM " & cboMonthlyMonth1 & " " & "TO " & cboMonthlyMonth2 & " " & cboMonthlyYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'LogAudit "V", "WEEKLY VEHICLES GROSS PROFIT REPORT", "DATE " & cboMonthlyMonth1 & " " & cboMonthlyMonth2 & " " & cboMonthlyYear
    ElseIf Combo1.Text = "QUARTERLY VEHICLES GROSS PROFIT REPORT" Then
        PrintQuarterlyGross
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call NEW_LogAudit("V", "VEHICLE SALES", "", "", "", Combo1 & " " & cboQuarter & " " & cboQuarterlyYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'LogAudit "V", "QUARTERLY VEHICLES GROSS PROFIT REPORT", "DATE " & cboQuarter & " " & cboQuarterlyYear
    ElseIf Combo1.Text = "YEARLY VEHICLES GROSS PROFIT REPORT" Then
        PrintYearlyGross
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call NEW_LogAudit("V", "VEHICLE SALES", "", "", "", Combo1 & " " & "FROM " & cboYearly1 & " " & "TO " & cboYearly2, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'LogAudit "V", "QUARTERLY VEHICLES GROSS PROFIT REPORT", "DATE " & cboYearly1 & " " & cboYearly2
    End If
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Combo1_Change()
    picDaily.Visible = False: picWeekly.Visible = False: picMonthly.Visible = False: picQuarterly.Visible = False: picYearly.Visible = False
    If Combo1.Text = "DAILY VEHICLES GROSS PROFIT REPORT" Then
        picDaily.Visible = True
    ElseIf Combo1.Text = "WEEKLY VEHICLES GROSS PROFIT REPORT" Then
        picWeekly.Visible = True
    ElseIf Combo1.Text = "MONTHLY VEHICLES GROSS PROFIT REPORT" Then
        picMonthly.Visible = True
    ElseIf Combo1.Text = "QUARTERLY VEHICLES GROSS PROFIT REPORT" Then
        picQuarterly.Visible = True
    ElseIf Combo1.Text = "YEARLY VEHICLES GROSS PROFIT REPORT" Then
        picYearly.Visible = True
    End If
End Sub

Private Sub Combo1_Click()
    Combo1_Change
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE SALES)"
            Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLE SALES", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SetComboWidth Combo1, "320"
    With Combo1
        .AddItem "DAILY VEHICLES GROSS PROFIT REPORT"
        .AddItem "WEEKLY VEHICLES GROSS PROFIT REPORT"
        .AddItem "MONTHLY VEHICLES GROSS PROFIT REPORT"
        .AddItem "QUARTERLY VEHICLES GROSS PROFIT REPORT"
        .AddItem "YEARLY VEHICLES GROSS PROFIT REPORT"
        .ListIndex = 0
    End With
    'DAILY
    dtDailyFrom.Value = firstDay(Date)
    dtDailyTo.Value = Date
    'WEEKLY
    dtWeeklyFrom.Value = firstDay(Date)
    dtWeeklyTo.Value = Date
    'MONTHLY
    fillcbomonth cboMonthlyMonth1
    fillcbomonth cboMonthlyMonth2
    FillCboMoreYear cboMonthlyYear
    'QUARTERLY
    cboQuarter.AddItem "ALL"
    cboQuarter.AddItem "QUARTER I"
    cboQuarter.AddItem "QUARTER II"
    cboQuarter.AddItem "QUARTER III"
    cboQuarter.AddItem "QUARTER IV"
    cboQuarter.ListIndex = 0
    FillCboMoreYear cboQuarterlyYear
    cboQuarterlyYear.Text = Year(LOGDATE)
    'YEARLY
    FillCboMoreYear cboYearly1
    FillCboMoreYear cboYearly2
    cboYearly1.Text = Year(LOGDATE)
    cboYearly2.Text = Year(LOGDATE)
    cboMonthlyMonth1.Text = The_month(Month(LOGDATE))
    cboMonthlyMonth2.Text = The_month(Month(LOGDATE))
    cboMonthlyYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

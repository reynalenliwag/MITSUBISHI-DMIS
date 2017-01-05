VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMISReports_PQIRReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PQIR Reports"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3660
   Icon            =   "PQIRReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3660
   Begin VB.Frame Frame6 
      Height          =   885
      Left            =   120
      TabIndex        =   28
      Top             =   510
      Width           =   3465
      Begin MSComCtl2.DTPicker txtInDate1 
         Height          =   315
         Left            =   1050
         TabIndex        =   29
         Top             =   120
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39562
      End
      Begin MSComCtl2.DTPicker txtInDate2 
         Height          =   315
         Left            =   1050
         TabIndex        =   30
         Top             =   480
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39562
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   32
         Top             =   150
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   2
         Left            =   300
         TabIndex        =   31
         Top             =   480
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   120
      TabIndex        =   11
      Top             =   510
      Visible         =   0   'False
      Width           =   3465
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   0
         Left            =   120
         TabIndex        =   27
         Top             =   2700
         Width           =   2985
      End
      Begin VB.TextBox txtCJRNumber 
         Height          =   375
         Left            =   1260
         TabIndex        =   8
         Top             =   660
         Width           =   2655
      End
      Begin VB.TextBox txtPQIRNumber 
         Height          =   315
         Left            =   1230
         TabIndex        =   7
         Top             =   150
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PQIR Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CJR Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   12
         Top             =   720
         Width           =   1065
      End
   End
   Begin VB.ComboBox cboPQIRReportType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   150
      Width           =   3465
   End
   Begin Crystal.CrystalReport rptPQIR 
      Left            =   120
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   1830
      MouseIcon       =   "PQIRReport.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "PQIRReport.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Close Window"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   1110
      MouseIcon       =   "PQIRReport.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "PQIRReport.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   1440
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Height          =   0
      Left            =   4650
      TabIndex        =   14
      Top             =   4050
      Visible         =   0   'False
      Width           =   345
      Begin VB.ComboBox cboDealers 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   240
         Width           =   2865
      End
      Begin VB.ComboBox CboClaimType 
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   600
         Width           =   2865
      End
      Begin VB.ComboBox CboJudgement 
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Top             =   990
         Width           =   2865
      End
      Begin VB.ComboBox CboRecommendation 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Top             =   1770
         Width           =   2865
      End
      Begin VB.ComboBox CboClassification 
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         Top             =   2130
         Width           =   2865
      End
      Begin MSComCtl2.DTPicker txtTo 
         Height          =   345
         Left            =   3090
         TabIndex        =   4
         Top             =   1380
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39562
      End
      Begin MSComCtl2.DTPicker txtFrom 
         Height          =   345
         Left            =   1650
         TabIndex        =   3
         Top             =   1380
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39562
      End
      Begin VB.Line Line1 
         X1              =   3810
         X2              =   4590
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dealer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   450
         TabIndex        =   20
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "In/Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   960
         TabIndex        =   19
         Top             =   1425
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Claim Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   630
         TabIndex        =   18
         Top             =   630
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Judgement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   600
         TabIndex        =   17
         Top             =   1035
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Recommendation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   60
         TabIndex        =   16
         Top             =   1830
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Classification"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   420
         TabIndex        =   15
         Top             =   2160
         Width           =   1155
      End
   End
   Begin VB.Frame Frame4 
      Height          =   75
      Left            =   4290
      TabIndex        =   22
      Top             =   4050
      Visible         =   0   'False
      Width           =   195
      Begin VB.TextBox txtCS_CJRNumber 
         Height          =   375
         Left            =   1260
         TabIndex        =   24
         Top             =   210
         Width           =   2655
      End
      Begin VB.TextBox txtCS_PQIRNumber 
         Height          =   375
         Left            =   1260
         TabIndex        =   23
         Top             =   660
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PQIR Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   90
         TabIndex        =   26
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CJR Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   150
         TabIndex        =   25
         Top             =   270
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmPMISReports_PQIRReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPQIR                                             As ADODB.Recordset


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    rptPQIR.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPQIR.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    Select Case UCase(cboPQIRReportType.Text)
        Case "PARTS QUALITY INFORMATION REPORT"
            Dim FromDate                               As Date
            Dim ToDate                                 As Date
            Dim SubjectFilter, ClaimNo, STATUS, PartNoFilter As String

            FromDate = CDate(txtInDate1.Value)
            ToDate = CDate(txtInDate2.Value)

            rptPQIR.ReportTitle = "Parts Quality Information Report (" & cboDealerCode & ")"
            rptPQIR.WindowTitle = "PARTS QUALITY INFORMATION REPORT"
            rptPQIR.Formulas(11) = "mindate = '" & FromDate & "'"
            rptPQIR.Formulas(12) = "maxdate = '" & ToDate & "'"
            Set rsPQIR = New ADODB.Recordset
            Set rsPQIR = gconDMIS.Execute("Select DATEPQI from PMIS_PQIR where DATEPQI >= '" & FromDate & "' and DATEPQI  <= '" & ToDate & "'")
            If Not rsPQIR.EOF And Not rsPQIR.BOF Then
                Screen.MousePointer = vbHourglass
                PrintSQLReport rptPQIR, PMIS_REPORT_PATH & "Parts Quality Info report.rpt", "{PMIS_PQIR.DATEPQI} >= Date(" & Year(FromDate) & "," & Month(FromDate) & "," & Day(FromDate) & ") AND {PMIS_PQIR.DATEPQI} <= Date(" & Year(ToDate) & "," & Month(ToDate) & "," & Day(ToDate) & ")", DMIS_REPORT_Connection, 1
                Screen.MousePointer = vbDefault
            Else
                ShowNoRecord
                Screen.MousePointer = 0
                Exit Sub
            End If
        Case "PQIR DETAIL REPORT"
            rptPQIR.ReportTitle = "PQIR Detail"
            rptPQIR.WindowTitle = "PQIR Detail"
            rptPQIR.Formulas(11) = "PQIRNo = '" & txtPQIRNumber.Text & "'"
            If RTrim(LTrim(txtPQIRNumber.Text)) = "" Then
                MsgSpeechBox "PQIR Number must not be empty!"
                txtPQIRNumber.SetFocus
                Exit Sub
            End If
            Set rsPQIR = New ADODB.Recordset
            Set rsPQIR = gconDMIS.Execute("Select PQI_CODE from PMIS_PQIR where PQI_CODE = " & N2Str2Null(RTrim(LTrim(txtPQIRNumber.Text))))
            If Not rsPQIR.EOF And Not rsPQIR.BOF Then
                Screen.MousePointer = vbHourglass
                PrintSQLReport rptPQIR, PMIS_REPORT_PATH & "PQIR Detail.rpt", "{PMIS_PQIR.PQI_CODE} = " & N2Str2Null(RTrim(LTrim(txtPQIRNumber.Text))), DMIS_REPORT_Connection, 1
                Screen.MousePointer = vbDefault
            Else
                ShowNoRecord
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Case "CLAIM JUDGEMENT REPORT"
            rptPQIR.ReportTitle = "Claim Judgement Report"
            rptPQIR.WindowTitle = "CLAIM JUDGEMENT REPORT"
            rptPQIR.Formulas(11) = "CJRNo = '" & txtPQIRNumber.Text & "'"

            If Trim(txtPQIRNumber.Text) = "" Then
                MsgSpeechBox "CJR Number must not be empty!"
                txtPQIRNumber.SetFocus
                Exit Sub
            End If
            Dim rsCJR                                  As ADODB.Recordset
            Set rsCJR = New ADODB.Recordset
            Set rsCJR = gconDMIS.Execute("Select PQINO from PMIS_PQIR where CRJ_NO = " & N2Str2Null(RTrim(LTrim(txtPQIRNumber.Text))))
            If Not rsCJR.EOF And Not rsCJR.BOF Then
                Screen.MousePointer = vbHourglass
                PrintSQLReport rptPQIR, PMIS_REPORT_PATH & "Claim Judgement Report.rpt", "{PMIS_PQIR.CRJ_NO} = " & N2Str2Null(LTrim(RTrim(txtPQIRNumber.Text))), DMIS_REPORT_Connection, 1
                Screen.MousePointer = vbDefault
            Else
                ShowNoRecord
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    txtInDate1 = Format(firstDay(LOGDATE), "DD-MMM-YY")
    txtInDate2 = Format(LOGDATE, "DD-MMM-YY")
    Screen.MousePointer = 0
    'FillCbo1
    With cboPQIRReportType
        .AddItem "Parts Quality Information Report"
        .AddItem "PQIR Detail Report"
        '.AddItem "PQIR Summary Report"
        .AddItem "Claim Judgement Report"
        '.AddItem "Claim Status Detail"
        .ListIndex = 0
    End With
End Sub
Private Sub cboPQIRReportType_Click()
    Select Case UCase(cboPQIRReportType.Text)
        Case "PARTS QUALITY INFORMATION REPORT"
            Frame6.Visible = True
            Frame2.Visible = False
            '            Frame3.Visible = False
            '            Frame4.Visible = False
        Case "PQIR DETAIL REPORT"
            Label1(7).Caption = "PQIR Number"
            Frame6.Visible = False
            Frame2.Visible = True
            txtPQIRNumber.SetFocus
            'Frame3.Visible = False
            'Frame4.Visible = False
        Case "PQIR SUMMARY REPORT"
            'Frame1.Visible = False
            'Frame2.Visible = False
            'Frame3.Visible = True
            'Frame4.Visible = False
            'txtFrom.Value = Format(firstDay(LOGDATE), "DD-MMM-YY")
            'txtTo.Value = Format(LOGDATE, "DD-MMM-YY")
        Case "CLAIM JUDGEMENT REPORT"
            Label1(7).Caption = "CJR Number"
            Frame6.Visible = False
            Frame2.Visible = True
            txtPQIRNumber.SetFocus
            'Frame1.Visible = False
            'Frame2.Visible = False
            'Frame3.Visible = False
            'Frame4.Visible = True
        Case "CLAIM STATUS REPORT"
            'Frame1.Visible = False
            'Frame2.Visible = False
            'Frame3.Visible = False
            'Frame4.Visible = False
            'Frame5.Visible = True
    End Select
End Sub
'Sub FillCbo1()
'    Combo_Loadval cboDealerName, gconDMIS.Execute("select DEALER_NAME from all_dealers order by dealer_name asc")
'    cboDealerName.AddItem "ALL", 0
'    cboDealerName.ListIndex = 0
'End Sub

Private Sub txtCS_CJRNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPrint_Click
    End If
End Sub

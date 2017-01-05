VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_LogSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Summary Report"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   Icon            =   "Report_LogSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Report_LogSummary.frx":058A
      Left            =   1020
      List            =   "Report_LogSummary.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   180
      Width           =   2445
   End
   Begin VB.OptionButton optRanged 
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
      Height          =   360
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.OptionButton optRangedMonthly 
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
      Height          =   360
      Left            =   1020
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin Crystal.CrystalReport rptLogs 
      Left            =   0
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   30
      TabIndex        =   15
      Top             =   1830
      Width           =   4275
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Cancelled Sales Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   2100
         TabIndex        =   28
         Tag             =   "{CRIS_ViewLog.LogName}= 'SALES ORDER CANCELLED'"
         Top             =   1620
         Width           =   2115
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Cancelled Sales Invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   2100
         TabIndex        =   27
         Tag             =   "{CRIS_ViewLog.LogName}= 'SALES INVOICE CANCELLED'"
         Top             =   1320
         Width           =   2085
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Sales Invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   2100
         TabIndex        =   26
         Tag             =   "{CRIS_ViewLog.LogName}= 'SALES INVOICE'"
         Top             =   1020
         Width           =   1425
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Sales Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   2100
         TabIndex        =   25
         Tag             =   "{CRIS_ViewLog.LogName}= 'SALES ORDER'"
         Top             =   720
         Width           =   1425
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Test Drive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   2100
         TabIndex        =   24
         Tag             =   "{CRIS_ViewLog.LogName}='TEST DRIVE'"
         Top             =   135
         Width           =   1425
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Loan Application"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   90
         TabIndex        =   23
         Tag             =   "{CRIS_ViewLog.LogName}='LOAN APPLICATION'"
         Top             =   1710
         Width           =   1695
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Initial Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Tag             =   "{CRIS_ViewLog.LogName}='INITIAL INQUIRY'"
         Top             =   180
         Width           =   1515
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Letters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   21
         Tag             =   "{CRIS_ViewLog.LogName}='LETTERS'"
         Top             =   1455
         Width           =   1425
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Tag             =   "{CRIS_ViewLog.LogName}='EMAIL'"
         Top             =   930
         Width           =   1425
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Visits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   19
         Tag             =   "{CRIS_ViewLog.LogName}='VISITS'"
         Top             =   1200
         Width           =   1425
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Quotation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   2100
         TabIndex        =   18
         Tag             =   "{CRIS_ViewLog.LogName}= 'QUOTATION' "
         Top             =   435
         Width           =   1425
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Calls"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Tag             =   "{CRIS_ViewLog.LogName}='CALLS'"
         Top             =   435
         Width           =   975
      End
      Begin VB.CheckBox chk_Opt 
         Caption         =   "Sales Appointments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Tag             =   "{CRIS_ViewLog.LogName}='APPOINTMENT'"
         Top             =   690
         Width           =   1755
      End
      Begin VB.Label Label6 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   1830
         TabIndex        =   33
         Top             =   690
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   1890
         TabIndex        =   32
         Top             =   1740
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   3690
         TabIndex        =   31
         Top             =   480
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   3690
         TabIndex        =   30
         Top             =   150
         Width           =   165
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
      Height          =   825
      Left            =   2250
      MouseIcon       =   "Report_LogSummary.frx":05AC
      MousePointer    =   99  'Custom
      Picture         =   "Report_LogSummary.frx":06FE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   4140
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
      Left            =   1380
      MouseIcon       =   "Report_LogSummary.frx":0B49
      MousePointer    =   99  'Custom
      Picture         =   "Report_LogSummary.frx":0C9B
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   4140
      Width           =   885
   End
   Begin VB.PictureBox picRange 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   3735
      TabIndex        =   5
      Top             =   960
      Width           =   3735
      Begin MSComCtl2.DTPicker dtpToDateLog 
         Height          =   435
         Left            =   930
         TabIndex        =   6
         Top             =   480
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   767
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
         CalendarForeColor=   0
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   52887553
         CurrentDate     =   39232
      End
      Begin MSComCtl2.DTPicker dtpFromDateLog 
         Height          =   375
         Left            =   915
         TabIndex        =   7
         Top             =   60
         Width           =   2430
         _ExtentX        =   4286
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
         CalendarForeColor=   0
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   52887553
         CurrentDate     =   39203
      End
      Begin VB.Label Label2 
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
         Left            =   180
         TabIndex        =   9
         Top             =   570
         Width           =   885
      End
      Begin VB.Label Label1 
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
         Left            =   0
         TabIndex        =   8
         Top             =   90
         Width           =   1185
      End
   End
   Begin VB.PictureBox picMonthly 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   -90
      ScaleHeight     =   915
      ScaleWidth      =   3735
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1140
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   510
         Width           =   2385
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   1140
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   90
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   570
         TabIndex        =   13
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   420
         TabIndex        =   12
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Applicable only to Prospect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Index           =   5
      Left            =   240
      TabIndex        =   35
      Top             =   3870
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   34
      Top             =   3900
      Width           =   165
   End
   Begin VB.Label Label5 
      Caption         =   "FOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   29
      Top             =   210
      Width           =   645
   End
End
Attribute VB_Name = "frmSMIS_Report_LogSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim i                                                             As Integer
    Dim FilterStringValue                                             As String

    On Error GoTo ErrorCode

    For i = 0 To chk_Opt.Count - 1
        If chk_Opt(i).Value = 1 Then
            FilterStringValue = FilterStringValue & chk_Opt(i).Tag & " OR "
        End If
    Next
    If FilterStringValue = "" Then
        MsgBox " Please Select Option From The List", vbInformation
        Exit Sub
    End If
    FilterStringValue = "(" & Left(FilterStringValue, Len(FilterStringValue) - 4) & ")"



    Screen.MousePointer = 11

    '
    '

    Dim FDate                                                         As Date
    Dim TDate                                                         As Date
    Dim rsLogs                                                        As ADODB.Recordset
    Dim RecordSelection                                               As String
    Set rsLogs = New ADODB.Recordset

    FDate = CDate(dtpFromDateLog.Value)
    TDate = CDate(dtpToDateLog.Value)

    If Combo2.ListIndex = 0 Then
        Dim rsC_Log                                                   As ADODB.Recordset
        Set rsC_Log = New ADODB.Recordset
        Set rsC_Log = gconDMIS.Execute("Select * from CRIS_ViewLog where cscde is not NULL")
        If Not rsC_Log.EOF And Not rsC_Log.BOF Then
            If picMonthly.Visible = False Then
                rptLogs.Formulas(0) = "datefrom = '" & " FROM " & FDate & " TO " & TDate & "'"
                rptLogs.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
                rptLogs.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptLogs.Formulas(3) = "PRINTEDBY = '" & LOGNAME & "'"
                RecordSelection = "({CRIS_ViewLog.Deyt} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CRIS_ViewLog.Deyt} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")) "
            Else
                rptLogs.Formulas(0) = "datefrom = '" & "FOR THE MONTH OF " & Combo1 & " " & Text1 & "'"
                rptLogs.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
                rptLogs.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptLogs.Formulas(3) = "PRINTEDBY = '" & LOGNAME & "'"
                RecordSelection = "(YEAR({CRIS_ViewLog.Deyt}) =" & Text1 & " AND MONTH({CRIS_ViewLog.Deyt}) =" & What_month(Combo1) & ")"
            End If

            PrintSQLReport rptLogs, CRIS_REPORT_PATH & "CustomerLogReport.rpt", RecordSelection & " AND " & FilterStringValue, CRIS_REPORT_PATH, 1
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
             If picMonthly.Visible = True Then
                 Call NEW_LogAudit("V", "LOG SUMMARY REPORT", "", "", "", "CUSTOMER LOG SUMMARY REPORT -" & Combo1 & " " & Text1, "", "")
             Else
                 Call NEW_LogAudit("V", "LOG SUMMARY REPORT", "", "", "", "CUSTOMER LOG SUMMARY REPORT -" & "FROM " & dtpFromDateLog & " " & "TO " & dtpToDateLog, "", "")
             End If
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

            'LogAudit "V", "CUSTOMER LOG SUMMARY REPORT "
        Else
            ShowNoRecord
        End If
    Else
        Dim rsP_Log                                                   As ADODB.Recordset
        Set rsP_Log = New ADODB.Recordset
        Set rsP_Log = gconDMIS.Execute("Select * from CRIS_ViewLog where cscde is NULL")
        If Not rsP_Log.EOF And Not rsP_Log.BOF Then
            If picMonthly.Visible = False Then
                rptLogs.Formulas(0) = "datefrom = '" & " FROM " & FDate & " TO " & TDate & "'"
                rptLogs.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
                rptLogs.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptLogs.Formulas(3) = "PRINTEDBY = '" & LOGNAME & "'"
                RecordSelection = "({CRIS_ViewLog.Deyt} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CRIS_ViewLog.Deyt} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")) "
            Else
                rptLogs.Formulas(0) = "datefrom = '" & "FOR THE MONTH OF " & Combo1 & " " & Text1 & "'"
                rptLogs.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
                rptLogs.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptLogs.Formulas(3) = "PRINTEDBY = '" & LOGNAME & "'"
                RecordSelection = "(YEAR({CRIS_ViewLog.Deyt}) =" & Text1 & " AND MONTH({CRIS_ViewLog.Deyt}) =" & What_month(Combo1) & ")"
            End If
            PrintSQLReport rptLogs, CRIS_REPORT_PATH & "ProspectLogReport.rpt", RecordSelection & " AND " & FilterStringValue, CRIS_REPORT_PATH, 1
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
             If picMonthly.Visible = True Then
                 Call NEW_LogAudit("V", "LOG SUMMARY REPORT", "", "", "", "PROSPECT LOG SUMMARY REPORT -" & Combo1 & " " & Text1, "", "")
             Else
                 Call NEW_LogAudit("V", "LOG SUMMARY REPORT", "", "", "", "PROSPECT LOG SUMMARY REPORT -" & "FROM " & dtpFromDateLog & " " & "TO " & dtpToDateLog, "", "")
             End If
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
            'LogAudit "V", "PROSPECT LOG SUMMARY REPORT "
        Else
            ShowNoRecord
        End If
    End If

    'End of update

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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LOG SUMMARY REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "LOG SUMMARY REPORT", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    Combo2.ListIndex = 0
    dtpFromDateLog.Value = firstDay(LOGDATE)
    dtpToDateLog.Value = LOGDATE
    fillcbomonth Combo1
    Combo1.Text = MonthName(Month(LOGDATE), False)
    Text1 = Year(LOGDATE)
End Sub

Private Sub optCustomerLog_Click()
    Label1.Enabled = True
    Label2.Enabled = True
    dtpFromDateLog.Enabled = True
    dtpToDateLog.Enabled = True
    cmdPrint.Enabled = True
End Sub

Private Sub optProspectLog_Click()
    Label1.Enabled = True
    Label2.Enabled = True
    dtpFromDateLog.Enabled = True
    dtpToDateLog.Enabled = True
    cmdPrint.Enabled = True
End Sub

Private Sub optRanged_Click()
    optRangedMonthly_Click

End Sub

Private Sub optRangedMonthly_Click()
    If optRangedMonthly.Value = True Then
        picMonthly.Visible = True
        picRange.Visible = False
    Else
        picMonthly.Visible = False
        picRange.Visible = True
    End If
End Sub


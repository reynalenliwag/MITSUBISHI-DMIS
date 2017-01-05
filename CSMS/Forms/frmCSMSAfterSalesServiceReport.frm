VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSAfterSalesServiceReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "After Sales Service Report"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3225
   Icon            =   "frmCSMSAfterSalesServiceReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   3225
   Begin VB.OptionButton optByDateRange 
      Caption         =   "By Date Range"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1725
      TabIndex        =   1
      Top             =   135
      Width           =   1605
   End
   Begin VB.OptionButton optByMonthYear 
      Caption         =   "By Month/Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   1575
   End
   Begin Crystal.CrystalReport rptAfterSalesServiceReport 
      Left            =   150
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "After Sales Service Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.PictureBox pixMonthYear 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   90
      ScaleHeight     =   930
      ScaleWidth      =   3225
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   3225
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select month from the list"
         Top             =   75
         Width           =   1815
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select year from the list"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   465
         TabIndex        =   10
         Top             =   105
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   585
         TabIndex        =   9
         Top             =   495
         Width           =   390
      End
   End
   Begin VB.PictureBox pixByDateRange 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   180
      ScaleHeight     =   930
      ScaleWidth      =   3030
      TabIndex        =   11
      Top             =   420
      Width           =   3030
      Begin MSComCtl2.DTPicker dtpFromDateSalesService 
         Height          =   330
         Left            =   1005
         TabIndex        =   4
         Top             =   105
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20185089
         CurrentDate     =   39203
      End
      Begin MSComCtl2.DTPicker dtpToDateSalesService 
         Height          =   330
         Left            =   1005
         TabIndex        =   5
         Top             =   495
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20185089
         CurrentDate     =   39203
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   285
         TabIndex        =   13
         Top             =   525
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   165
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      MouseIcon       =   "frmCSMSAfterSalesServiceReport.frx":27A2
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSAfterSalesServiceReport.frx":28F4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Close Window"
      Top             =   1350
      Width           =   915
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
      Height          =   855
      Left            =   1140
      MouseIcon       =   "frmCSMSAfterSalesServiceReport.frx":2D3F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSAfterSalesServiceReport.frx":2E91
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Report"
      Top             =   1350
      Width           =   915
   End
End
Attribute VB_Name = "frmCSMSAfterSalesServiceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboYear_KeyPress(KeyAscii As Integer)
    cmdPrint_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "AFTER SALES REPORT") = False Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo Errorcode

    Dim FDate                                          As Date
    Dim TDate                                          As Date

    FDate = CDate(dtpFromDateSalesService.Value)
    TDate = CDate(dtpToDateSalesService.Value)

    Dim rsAfterSalesService                            As ADODB.Recordset
    Set rsAfterSalesService = New ADODB.Recordset
    Set rsAfterSalesService = gconDMIS.Execute("SELECT * from CSMS_Repor where TRANSTYPE = '" & "R" & "'")
    If Not rsAfterSalesService.BOF And Not rsAfterSalesService.EOF Then
        If optByMonthYear.Value = True Then
            Dim rsMonthYear                            As ADODB.Recordset
            Set rsMonthYear = New ADODB.Recordset
            Set rsMonthYear = gconDMIS.Execute("SELECT * from CSMS_Repor where Month(DTE_COMP) = '" & What_month(cboMonth.Text) & "' AND Year(DTE_COMP) = " & cboYear.Text)
            If Not rsMonthYear.BOF And Not rsMonthYear.EOF Then

                'JUN 02/05/2005
                rptAfterSalesServiceReport.Formulas(0) = "Company Name = '" & COMPANY_NAME & "'"
                rptAfterSalesServiceReport.Formulas(1) = "Company Address = '" & COMPANY_ADDRESS & "'"
                rptAfterSalesServiceReport.Formulas(2) = "Printedby = '" & LOGNAME & "'"

                rptAfterSalesServiceReport.ReportTitle = "After Sales Service For The Month Of " & cboMonth.Text & "  " & cboYear.Text
                rptAfterSalesServiceReport.WindowTitle = "After Sales Service For The Month Of " & cboMonth.Text & "  " & cboYear.Text
                PrintSQLReport rptAfterSalesServiceReport, CSMS_REPORT_PATH & "AfterSalesServiceReport.rpt", "Month({CSMS_Repor.DTE_COMP}) = " & What_month(cboMonth.Text) & " AND Year({CSMS_Repor.DTE_COMP}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1

                LogAudit "V", "AFTER SALE SERVICE - REPORT", cboMonth & cboYear

            Else
                ShowNoRecord
                Exit Sub
            End If
        Else
            Dim rsDateRange                            As ADODB.Recordset
            Set rsDateRange = New ADODB.Recordset
            Set rsDateRange = gconDMIS.Execute("SELECT * from CSMS_Repor where DTE_COMP >= '" & FDate & "' AND DTE_COMP <= '" & TDate & "'")
            If Not rsDateRange.BOF And Not rsDateRange.EOF Then

                'JUN 02/05/2005
                rptAfterSalesServiceReport.Formulas(0) = "Company Name = '" & COMPANY_NAME & "'"
                rptAfterSalesServiceReport.Formulas(1) = "Company Address = '" & COMPANY_ADDRESS & "'"
                rptAfterSalesServiceReport.Formulas(2) = "Printedby = '" & LOGNAME & "'"

                rptAfterSalesServiceReport.ReportTitle = "After Sales Service From " & FDate & "  To " & TDate
                rptAfterSalesServiceReport.WindowTitle = "After Sales Service From " & FDate & "  To " & TDate
                PrintSQLReport rptAfterSalesServiceReport, CSMS_REPORT_PATH & "AfterSalesServiceReport.rpt", "{CSMS_Repor.DTE_COMP} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_Repor.DTE_COMP} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", CSMS_REPORT_CONNECTION, 1

                LogAudit "V", "AFTER SALES SERVICE REPORT", dtpFromDateSalesService & "-" & dtpToDateSalesService
            Else
                ShowNoRecord
                Exit Sub
            End If
        End If
    Else
        ShowNoRecord
    End If
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
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0

    dtpFromDateSalesService.Value = firstDay(LOGDATE)
    dtpToDateSalesService.Value = LOGDATE
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    optByMonthYear.Value = True
End Sub

Private Sub optByDateRange_Click()
    pixByDateRange.Visible = True
    pixMonthYear.Visible = False
End Sub

Private Sub optByMonthYear_Click()
    pixByDateRange.Visible = False
    pixMonthYear.Visible = True
End Sub


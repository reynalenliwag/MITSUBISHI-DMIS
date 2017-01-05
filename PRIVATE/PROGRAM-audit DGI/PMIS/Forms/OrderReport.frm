VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_OrderReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Report"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   ForeColor       =   &H00DEDFDE&
   Icon            =   "OrderReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3090
   Begin VB.CheckBox chkHistIssuance 
      Caption         =   "Look in History File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   900
      Width           =   2415
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
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   510
      Width           =   1965
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
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptOrderReport 
      Left            =   30
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Issuances"
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
      Height          =   795
      Left            =   1680
      MouseIcon       =   "OrderReport.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "OrderReport.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   1230
      Width           =   735
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
      Height          =   795
      Left            =   960
      MouseIcon       =   "OrderReport.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "OrderReport.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Report"
      Top             =   1230
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
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   3
      Top             =   3000
      Width           =   495
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
      Left            =   60
      TabIndex        =   2
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmPMISReports_OrderReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo Errorcode:

    If ORDER_REPORT = "HARI" Then
        If chkHistIssuance.Value = False Then
            If What_month(cboMonth) >= Month(Now) Then
                Dim RSPO_HD                            As ADODB.Recordset
                Set RSPO_HD = New ADODB.Recordset
                RSPO_HD.Open "select podate from PMIS_PO_Hd where TYPE = 'P' AND SUPCODE = 'H00001' AND month(podate) = " & What_month(cboMonth) & " AND year(podate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
                    Screen.MousePointer = 11
                    rptOrderReport.WindowTitle = "Order Report from HARI (Per Type & Classification)"
                    rptOrderReport.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptOrderReport.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    rptOrderReport.Formulas(11) = "ForTheMonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
                    PrintSQLReport rptOrderReport, PMIS_REPORT_PATH & "OrderReport_HARI.rpt", "{PMIS_Po_Hd.TYPE} = 'P' AND month({PMIS_Po_Hd.PODATE}) = " & What_month(cboMonth.Text) & " AND year({PMIS_Po_Hd.PODATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                    Screen.MousePointer = 0
                    LogAudit "V", "ORDER REPORT FROM HARI (Per Type & Classification)", cboMonth & "-" & cboYear
                Else
                    MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
                End If
            Else
                MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
            End If
        Else
            Dim RSPO_HIST                              As ADODB.Recordset
            Set RSPO_HIST = New ADODB.Recordset
            RSPO_HIST.Open "select podate from PMIS_Po_Hist where TYPE = 'P' AND SUPCODE = 'H00001' AND month(podate) = " & What_month(cboMonth) & " AND year(podate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSPO_HIST.EOF And Not RSPO_HIST.BOF Then
                Screen.MousePointer = 11
                rptOrderReport.WindowTitle = "Order Report from HARI (Per Type & Classification)"
                rptOrderReport.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptOrderReport.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptOrderReport.Formulas(11) = "ForTheMonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
                PrintSQLReport rptOrderReport, PMIS_REPORT_PATH & "OrderReport_HARI_Hist.rpt", "{PMIS_Po_Hd.TYPE} = 'P' AND month({PMIS_Po_Hd.PODATE}) = " & What_month(cboMonth.Text) & " AND year({PMIS_Po_Hd.PODATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
                LogAudit "V", "ORDER REPORT FROM HARI (Per Type & Classification)", cboMonth & "-" & cboYear
            Else
                MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
            End If
        End If
    Else
        If chkHistIssuance.Value = False Then
            If What_month(cboMonth) >= Month(Now) Then
                Dim rsPO_HD_NonHARI                    As ADODB.Recordset
                Set rsPO_HD_NonHARI = New ADODB.Recordset
                rsPO_HD_NonHARI.Open "select podate from PMIS_PO_Hd where TYPE = 'P' AND SUPCODE <> 'H00001' AND month(podate) = " & What_month(cboMonth) & " AND year(podate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsPO_HD_NonHARI.EOF And Not rsPO_HD_NonHARI.BOF Then
                    Screen.MousePointer = 11
                    rptOrderReport.WindowTitle = "Order Report from Other Supplier (Per Type & Classification)"
                    rptOrderReport.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptOrderReport.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    rptOrderReport.Formulas(11) = "ForTheMonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
                    PrintSQLReport rptOrderReport, PMIS_REPORT_PATH & "OrderReport_NonHARI.rpt", "{PMIS_Po_Hd.TYPE} = 'P' AND month({PMIS_Po_Hd.PODATE}) = " & What_month(cboMonth.Text) & " AND year({PMIS_Po_Hd.PODATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                    Screen.MousePointer = 0
                    LogAudit "V", "ORDER REPORT FROM OTHER SUPPLIER (Per Type & Classification)", cboMonth & "-" & cboYear
                Else
                    MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
                End If
            Else
                MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
            End If
        Else
            Dim rsPO_Hist_NonHARI                      As ADODB.Recordset
            Set rsPO_Hist_NonHARI = New ADODB.Recordset
            rsPO_Hist_NonHARI.Open "select podate from PMIS_Po_Hist where TYPE = 'P' AND SUPCODE <> 'H00001' AND month(podate) = " & What_month(cboMonth) & " AND year(podate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPO_Hist_NonHARI.EOF And Not rsPO_Hist_NonHARI.BOF Then
                Screen.MousePointer = 11
                rptOrderReport.WindowTitle = "Order Report from Other Supplier (Per Type & Classification)"
                rptOrderReport.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptOrderReport.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptOrderReport.Formulas(11) = "ForTheMonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
                PrintSQLReport rptOrderReport, PMIS_REPORT_PATH & "OrderReport_NonHARI_Hist.rpt", "{PMIS_Po_Hd.TYPE} = 'P' AND month({PMIS_Po_Hd.PODATE}) = " & What_month(cboMonth.Text) & " AND year({PMIS_Po_Hd.PODATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
                LogAudit "V", "ORDER REPORT FROM OTHER SUPPLIER (Per Type & Classification)", cboMonth & "-" & cboYear
            Else
                MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
            End If
        End If

    End If
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_OrderReport = Nothing
    UnloadForm Me
End Sub


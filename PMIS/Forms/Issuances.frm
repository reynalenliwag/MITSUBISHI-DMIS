VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMISReports_Issuances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Issuance"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Issuances.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   3315
   Begin VB.ComboBox cbotype 
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
      ItemData        =   "Issuances.frx":0E42
      Left            =   1080
      List            =   "Issuances.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   2145
   End
   Begin Crystal.CrystalReport rptIssuances 
      Left            =   4680
      Top             =   4560
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
      Left            =   2610
      MouseIcon       =   "Issuances.frx":0E46
      MousePointer    =   99  'Custom
      Picture         =   "Issuances.frx":0F98
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   1680
      Width           =   675
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
      Left            =   1950
      MouseIcon       =   "Issuances.frx":13E3
      MousePointer    =   99  'Custom
      Picture         =   "Issuances.frx":1535
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   1680
      Width           =   675
   End
   Begin VB.Frame frm1 
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   3015
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
         ItemData        =   "Issuances.frx":19D4
         Left            =   840
         List            =   "Issuances.frx":19D6
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select month from the list"
         Top             =   240
         Width           =   2145
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
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select year from the list"
         Top             =   630
         Width           =   2145
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
         Left            =   0
         TabIndex        =   9
         Top             =   240
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
         Left            =   0
         TabIndex        =   8
         Top             =   660
         Width           =   735
      End
   End
   Begin VB.Frame frm2 
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   3015
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Format          =   122945537
         CurrentDate     =   41034
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   0
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7935
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12015
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "BY"
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
      Left            =   120
      TabIndex        =   4
      Top             =   160
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   0
      Top             =   3000
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISReports_Issuances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSORD_HIST                                         As ADODB.Recordset
Dim LOCALACESS                                         As String
Public XREPtype                                        As String

Private Sub cbotype_Change()
    If cbotype = "As Of" Then
        frm1.Visible = False
        frm2.Visible = True
    Else
        frm2.Visible = False
        frm1.Visible = True
    End If
End Sub

Private Sub CBOtype_Click()
    If cbotype = "As Of" Then
        frm1.Visible = False
        frm2.Visible = True
    Else
        frm2.Visible = False
        frm1.Visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim XXX As String
    
    Dim rptApp                                              As CRAXDRT.Application
    Dim rptRep                                              As CRAXDRT.Report
    
    XXX = MonthName(Month(DTPicker1.Value))
    If ISSREPTYPE = "RIV_INPROCESS" Then
        ' If Function_Access(LOGID, "Acess_Print", "REPORTS RIV FOR WORKINPROGRESS") = False Then Exit Sub
    Else
        'If Function_Access(LOGID, "Acess_Print", "PARTS MONTHLY REPORT") = False Then Exit Sub
    End If
    On Error GoTo ErrorCode:
    Set RSORD_HIST = New ADODB.Recordset
    If cbotype.Text = "As Of" Then
        RSORD_HIST.Open "select trandate from PMIS_Ord_Hist where TYPE = '" & XREPtype & "' AND TranDate <= '" & CDate(DTPicker1.Value) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        RSORD_HIST.Open "select trandate from PMIS_Ord_Hist where TYPE = 'P' AND month(TranDate) = " & What_month(cboMonth) & " AND year(TranDate) = " & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    '===================================
    'updating code:     JAA - 02112008
    If cbotype.Text = "As Of" Then
        If ISSREPTYPE = "RIV_INPROCESS" Then
            Screen.MousePointer = 11
            rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                If XREPtype = "M" Then
                     If COMPANY_CODE = "DJM" Then
                        Me.WindowState = vbMaximized
                        CRViewer1.Height = Me.Height - 800
                        CRViewer1.Width = Me.Width
                        CRViewer1.ZOrder 0
                        Set rptApp = New CRAXDRT.Application
                        Set rptRep = rptApp.OpenReport(PMIS_REPORT_PATH & "RIV_InProcess_ASOF_MAT.Rpt", 1)
                        Call rptRep.ParameterFields(1).AddCurrentValue(CDate(DTPicker1.Value))

                                With CRViewer1
                                    .ReportSource = rptRep
                                    .DisplayGroupTree = False
                                    .DisplayTabs = False
                                    .DisplayToolbar = True
                                    .ViewReport
                                End With
        
                                Set rptApp = Nothing
                                Set rptRep = Nothing
                Else
                    PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RIV_InProcess.rpt", "{ORD_HD.TYPE} = 'P' AND {Ord_hd.TranDate} <= date(" & Year(DTPicker1.Value) & "," & Month(DTPicker1.Value) & "," & Day(DTPicker1.Value) & ") ", DMIS_REPORT_Connection, 1
                End If
            Else
                     If COMPANY_CODE = "DJM" Then
                        Me.WindowState = vbMaximized
                        CRViewer1.Height = Me.Height - 800
                        CRViewer1.Width = Me.Width
                        CRViewer1.ZOrder 0
                        Set rptApp = New CRAXDRT.Application
                        Set rptRep = rptApp.OpenReport(PMIS_REPORT_PATH & "RIV_InProcess_ASOF.Rpt", 1)
                        Call rptRep.ParameterFields(1).AddCurrentValue(CDate(DTPicker1.Value))
                        
                                With CRViewer1
                                    .ReportSource = rptRep
                                    .DisplayGroupTree = False
                                    .DisplayTabs = False
                                    .DisplayToolbar = True
                                    .ViewReport
                                End With
        
                                Set rptApp = Nothing
                                Set rptRep = Nothing
                    Else
                        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RIV_InProcess.rpt", "{ORD_HD.TYPE} = 'P' AND {Ord_hd.TranDate} <= date(" & Year(DTPicker1.Value) & "," & Month(DTPicker1.Value) & "," & Day(DTPicker1.Value) & ") ", DMIS_REPORT_Connection, 1
                    End If
                End If
            Call NEW_LogAudit("V", "RIV FOR WORKINPROGRESS", "", "", "", MonthName(Month(DTPicker1.Value)) & " " & Year(DTPicker1.Value), "RIV IN PROCESS", "")
            Screen.MousePointer = 0
        Else
            If Not RSORD_HIST.EOF And Not RSORD_HIST.EOF Then
                Screen.MousePointer = 11
                rptIssuances.WindowTitle = "MONTHLY ISSUANCE"
                rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "Issuances.rpt", "{ORD_HD.TYPE} = 'P' AND  {Ord_hd.TranDate} <= date('" & Year(DTPicker1.Value) & "," & Month(DTPicker1.Value) & "," & Day(DTPicker1.Value) & "')", DMIS_REPORT_Connection, 1
                PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "Issuancesum.rpt", "{ORD_HD.TYPE} = 'P' AND {Ord_hd.TranDate} <= date('" & Year(DTPicker1.Value) & "," & Month(DTPicker1.Value) & "," & Day(DTPicker1.Value) & "')", DMIS_REPORT_Connection, 1
                Call NEW_LogAudit("V", "PARTS MONTHLY REPORT", "", "", "", MonthName(Month(DTPicker1.Value)) & " " & Year(DTPicker1.Value), "ISSUANCES", "")
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
            End If
        End If
    Else
        If ISSREPTYPE = "RIV_INPROCESS" Then
            Screen.MousePointer = 11
            rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            If XREPtype = "M" Then
                If COMPANY_CODE = "DJM" Then
                    Me.WindowState = vbMaximized
                    CRViewer1.Height = Me.Height - 800
                    CRViewer1.Width = Me.Width
                    CRViewer1.ZOrder 0
                    Set rptApp = New CRAXDRT.Application
                    Set rptRep = rptApp.OpenReport(PMIS_REPORT_PATH & "RIV_InProcess_MAT.Rpt", 1)
                    
                    Call rptRep.ParameterFields(1).AddCurrentValue(What_month(cboMonth.Text))
                    Call rptRep.ParameterFields(2).AddCurrentValue(Year(cboYear.Text))
    
                            With CRViewer1
                                .ReportSource = rptRep
                                .DisplayGroupTree = False
                                .DisplayTabs = False
                                .DisplayToolbar = True
                                .ViewReport
                            End With
    
                            Set rptApp = Nothing
                            Set rptRep = Nothing
                Else
                    PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RIV_InProcess_mat.rpt", "{ORD_HD.TYPE} = '" & XREPtype & "' AND month({Ord_hd.TranDate}) = " & What_month(cboMonth.Text) & " AND year({Ord_hd.TranDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                End If
            Else
                If COMPANY_CODE = "DJM" Then
                    Me.WindowState = vbMaximized
                    CRViewer1.Height = Me.Height - 800
                    CRViewer1.Width = Me.Width
                    CRViewer1.ZOrder 0
                    Set rptApp = New CRAXDRT.Application
                    Set rptRep = rptApp.OpenReport(PMIS_REPORT_PATH & "RIV_InProcess.Rpt", 1)
                    
                    Call rptRep.ParameterFields(1).AddCurrentValue(What_month(cboMonth.Text))
                    Call rptRep.ParameterFields(2).AddCurrentValue(Year(cboYear.Text))
    
                            With CRViewer1
                                .ReportSource = rptRep
                                .DisplayGroupTree = False
                                .DisplayTabs = False
                                .DisplayToolbar = True
                                .ViewReport
                            End With
    
                            Set rptApp = Nothing
                            Set rptRep = Nothing
                Else
                PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RIV_InProcess.rpt", "{ORD_HD.TYPE} = 'P' AND month({Ord_hd.TranDate}) = " & What_month(cboMonth.Text) & " AND year({Ord_hd.TranDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                End If
        End If
           
            Call NEW_LogAudit("V", "RIV FOR WORKINPROGRESS", "", "", "", cboMonth & " " & cboYear, "RIV IN PROCESS", "")
            Screen.MousePointer = 0
        Else
            If Not RSORD_HIST.EOF And Not RSORD_HIST.EOF Then
                Screen.MousePointer = 11
                rptIssuances.WindowTitle = "MONTHLY ISSUANCE"
                rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "Issuances.rpt", "{ORD_HD.TYPE} = 'P' AND month({Ord_hd.TranDate}) = " & What_month(cboMonth.Text) & " AND year({Ord_hd.TranDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "Issuancesum.rpt", "{ORD_HD.TYPE} = 'P' AND month({Ord_hd.TranDate}) = " & What_month(cboMonth.Text) & " AND year({Ord_hd.TranDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                Call NEW_LogAudit("V", "PARTS MONTHLY REPORT", "", "", "", cboMonth & " " & cboYear, "ISSUANCES", "")
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
            End If
        End If
    End If
    '===================================
    Exit Sub
ErrorCode:
    MsgBox err.Description
    ShowVBError

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
            If ISSREPTYPE = "" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (PARTS MONTHLY REPORT)"
                Call frmALL_AuditInquiry.DisplayHistory("", "PARTS MONTHLY REPORT", "PRINTING")
            Else
                frmALL_AuditInquiry.Caption = "Audit Inquiry (RIV FOR WORKINPROGRESS)"
                Call frmALL_AuditInquiry.DisplayHistory("", "RIV FOR WORKINPROGRESS", "PRINTING")
            End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    
    cbotype.AddItem "Monthly"
    cbotype.AddItem "As Of"
    cbotype.ListIndex = 0
'    If ISSREPTYPE = "RIV_INPROCESS" Then
'        Me.Caption = "RIV In-Process"
'        cboMonth.Enabled = False
'        cboYear.Enabled = False
'    Else
'        Me.Caption = "Issuance Report"
'        cboMonth.Enabled = True
'        cboYear.Enabled = True
'    End If

    If ISSREPTYPE = "RIV_INPROCESS" Then
        LOCALACESS = "RIV FOR WORKINPROGRESS"

    Else
        LOCALACESS = "PARTS MONTHLY REPORT"
    End If
    Screen.MousePointer = 0
End Sub
Function getmonthcode(XXX As String)
    Dim Indx                                                As Integer
    Select Case cboMonth.Text
        Case "January": Indx = 1
        Case "February": Indx = 2
        Case "March": Indx = 3
        Case "April": Indx = 4
        Case "May": Indx = 5
        Case "June": Indx = 6
        Case "July": Indx = 7
        Case "August": Indx = 8
        Case "September": Indx = 9
        Case "October": Indx = 10
        Case "November": Indx = 11
        Case "December": Indx = 12
    End Select
    getmonthcode = Indx
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_Issuances = Nothing
    UnloadForm Me
End Sub


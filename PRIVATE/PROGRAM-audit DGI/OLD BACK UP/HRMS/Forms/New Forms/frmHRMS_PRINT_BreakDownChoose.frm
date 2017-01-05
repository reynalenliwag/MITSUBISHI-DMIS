VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPRINT_BreakDown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print BreakDown"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_PRINT_BreakDownChoose.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1920
   ScaleWidth      =   3120
   Begin VB.Timer tme_Load 
      Interval        =   200
      Left            =   300
      Top             =   2220
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   540
      Width           =   885
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Width           =   1845
   End
   Begin VB.ComboBox cboQuensina 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2775
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
      Left            =   2130
      MouseIcon       =   "frmHRMS_PRINT_BreakDownChoose.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_PRINT_BreakDownChoose.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1020
      Width           =   855
   End
   Begin Crystal.CrystalReport rptBreak 
      Left            =   180
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
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
      Left            =   1290
      MouseIcon       =   "frmHRMS_PRINT_BreakDownChoose.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_PRINT_BreakDownChoose.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label lblLoad 
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   1590
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "frmHRMSPRINT_BreakDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim matt                                                          As Integer
    If cboQuensina.Text = "1st Cut-Off" Then
        matt = 1
    ElseIf cboQuensina.Text = "2nd Cut-Off" Then
        matt = 2
    End If
    rptBreak.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptBreak.Formulas(1) = "PrintedBy = '" & LOGNAME & "'"
    rptBreak.Formulas(2) = "companyAddress = '" & COMPANY_ADDRESS & "'"
    Select Case Me.Caption
        Case "Print OverTime BreakDown":
            PrintSQLReport rptBreak, HRMS_REPORT_PATH & "OverTime BreakDown.rpt", "{hrms_overtime.cut_off} = '" & matt & "' AND {hrms_overtime.pay_month} = " & What_month(cboMOnth) & " and {hrms_overtime.pay_year} = " & cboyear & "", DMIS_REPORT_Connection, 1
            LogAudit "V", "PRINT OVERTIME BREAKDOWN", cboQuensina & "-" & cboMOnth & ", " & cboyear
        Case "Print Commission BreakDown":
            PrintSQLReport rptBreak, HRMS_REPORT_PATH & "Commission BreakDown.rpt", "{hrms_commission.cut_off} = '" & matt & "' AND {hrms_commission.pay_month} = " & What_month(cboMOnth) & " and {hrms_commission.pay_year} = " & cboyear & "", DMIS_REPORT_Connection, 1
            LogAudit "V", "PRINT COMMISSION BREAKDOWN", cboQuensina & "-" & cboMOnth & ", " & cboyear
        Case "Print Deduction BreakDown":
            PrintSQLReport rptBreak, HRMS_REPORT_PATH & "Deduction BreakDown.rpt", "{hrms_deductions.cut_off} = '" & matt & "' AND {hrms_deductions.pay_month} = " & What_month(cboMOnth) & " and {hrms_deductions.pay_year} = " & cboyear & "", DMIS_REPORT_Connection, 1
            LogAudit "V", "PRINT DEDUCTION BREAKDOWN", cboQuensina & "-" & cboMOnth & ", " & cboyear
        Case "Print Adjustment BreakDown":
            PrintSQLReport rptBreak, HRMS_REPORT_PATH & "Adjustment BreakDown.rpt", "{hrms_adjustment.cut_off} = '" & matt & "' AND {hrms_adjustment.pay_month} = " & What_month(cboMOnth) & " and {hrms_adjustment.pay_year} = " & cboyear & "", DMIS_REPORT_Connection, 1
            LogAudit "V", "PRINT ADJUSTMENT BREAKDOWN", cboQuensina & "-" & cboMOnth & ", " & cboyear
        Case "Allowance Computation Report"
            PrintSQLReport rptBreak, HRMS_REPORT_PATH & "Allowance Computation.rpt", "{hrms_overtime.cut_off} = '" & matt & "' AND {hrms_overtime.pay_month} = " & What_month(cboMOnth) & " and {hrms_overtime.pay_year} = " & cboyear, DMIS_REPORT_Connection, 1
            LogAudit "V", "Allowance Computation Report", cboQuensina & "-" & cboMOnth & ", " & cboyear
        Case "Loans Balances"
            'PrintSQLReport rptBreak, HRMS_REPORT_PATH & "Allowance Computation.rpt", "{hrms_overtime.cut_off} = '" & matt & "' AND {hrms_overtime.pay_month} = " & What_month(cboMonth) & " and {hrms_overtime.pay_year} = " & cboYear, DMIS_REPORT_Connection, 1
            LogAudit "V", "Allowance Computation Report", cboQuensina & "-" & cboMOnth & ", " & cboyear
         Case "ATM SUMMARY LIST"
            Call atm_summary
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cboQuensina.AddItem "1st Cut-Off"
    cboQuensina.AddItem "2nd Cut-Off"
    cboQuensina.ListIndex = 0
    fillcbomonth cboMOnth
    'FillcboYear cboyear
    fillcombo_up cboyear
    cboMOnth.Text = MonthName(MONTH(Now))
    cboyear.Text = YEAR(Now)
End Sub

Sub atm_summary()
    Screen.MousePointer = 11

    Dim RSTMP                                           As New ADODB.Recordset
    Dim XXX                                             As String
    Dim xlApp                                           As Excel.Application
    Dim xlbook                                          As Excel.Workbook
    Dim xlsheet                                         As Excel.Worksheet
    Dim cmd                                             As ADODB.Command
    
    Dim X As String
    Dim Y As String
    Dim z As String
    
    Dim vCUTOFF As Integer
    

    Y = What_month(cboMOnth.Text)

    If cboQuensina.Text = "1st Cut-Off" Then
        vCUTOFF = 1
    End If
    If cboQuensina.Text = "2nd Cut-Off" Then
        vCUTOFF = 2
    End If

    
    Set cmd = New ADODB.Command
    cmd.NamedParameters = True
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SP_HRMS_ATM_SUMMARY"
    cmd.ActiveConnection = gconDMIS
    cmd.Parameters.Append cmd.CreateParameter("@CUT_OFF", adVarChar, adParamInput, 15, vCUTOFF)
    cmd.Parameters.Append cmd.CreateParameter("@MONTH", adVarChar, adParamInput, 15, Y)
    cmd.Parameters.Append cmd.CreateParameter("@YEAR", adVarChar, adParamInput, 15, cboyear.Text)
    
    Set RSTMP = cmd.Execute
    
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        If Len(Dir(HRMS_REPORT_PATH & "atmsummary.xlt")) = 0 Then
            MessagePop InfoStop, "Error", "Atm Summary.xlt cannot be found in server Report Path." & vbCrLf & "Please contact I.T Department", vbInformation
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        Set xlApp = New Excel.Application
        Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "atmsummary.xlt")
        Set xlsheet = xlbook.Worksheets(1)
        
        xlsheet.Range("A8").CopyFromRecordset RSTMP
        xlApp.Visible = True
        If Not xlbook Is Nothing Then
            Set xlbook = Nothing
            Set xlApp = Nothing
        End If
        Set xlApp = Nothing
    Else
        Call ShowNoRecord
    End If
    Set RSTMP = Nothing
    Screen.MousePointer = 0
End Sub



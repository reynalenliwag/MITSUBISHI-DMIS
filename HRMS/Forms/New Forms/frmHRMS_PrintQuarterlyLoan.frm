VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMS_PrintQuarterlyLoan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Quarterly Remiittance"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_PrintQuarterlyLoan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   4185
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
      Left            =   3330
      MouseIcon       =   "frmHRMS_PrintQuarterlyLoan.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_PrintQuarterlyLoan.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   900
      Width           =   735
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2790
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1305
   End
   Begin VB.ComboBox cboQuarter 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmHRMS_PrintQuarterlyLoan.frx":0B27
      Left            =   180
      List            =   "frmHRMS_PrintQuarterlyLoan.frx":0B37
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   2565
   End
   Begin Crystal.CrystalReport rptLOAN 
      Left            =   60
      Top             =   1200
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
      Left            =   2610
      MouseIcon       =   "frmHRMS_PrintQuarterlyLoan.frx":0B4F
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_PrintQuarterlyLoan.frx":0CA1
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   900
      Width           =   735
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "YEAR"
      Height          =   225
      Index           =   1
      Left            =   2790
      TabIndex        =   5
      Top             =   180
      Width           =   465
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "QUARTER"
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   150
      Width           =   870
   End
End
Attribute VB_Name = "frmHRMS_PrintQuarterlyLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub GetLaonRecord(FM As Integer, SM As Integer, TM As Integer)
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim RSLOAN                                                        As New ADODB.Recordset
    Dim XTOTAL                                                        As Currency

    gconDMIS.Execute ("Delete From HRMS_Loan_Quarterly")

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_LoanMasDet Where Month(DEYT) = " & FM & " And YEAR(deyt) = " & cboyear & " Order By Empno ASC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set RSLOAN = gconDMIS.Execute("Select * From HRMS_LOAN_QUARTERLY Where Empno = '" & RSTMP!EMPNO & "'")
            If Not (RSLOAN.BOF And RSLOAN.EOF) Then
                XTOTAL = XTOTAL + RSLOAN!Month1 + RSLOAN!MOnth2 + RSLOAN!Month3
                gconDMIS.Execute ("Update HRMS_LOAN_QUARTERLY Set MONTH1 = " & RSLOAN!Month1 + RSTMP!AMOUNT & _
                                  ",Xtotal = " & XTOTAL & _
                                " Where Empno = '" & RSTMP!EMPNO & "'")
            Else
                XTOTAL = RSTMP!AMOUNT
                gconDMIS.Execute ("Insert Into HRMS_LOAN_QUARTERLY (EMPNO,MONTH1,MONTH2,MONTH3,XTOTAL) VALUES('" & RSTMP!EMPNO & _
                                  "'," & RSTMP!AMOUNT & "," & 0 & "," & 0 & "," & RSTMP!AMOUNT & ")")
            End If

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_LoanMasDet Where Month(DEYT) = " & SM & " And YEAR(deyt) = " & cboyear & " Order By Empno ASC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set RSLOAN = gconDMIS.Execute("Select * From HRMS_LOAN_QUARTERLY Where Empno = '" & RSTMP!EMPNO & "'")
            If Not (RSLOAN.BOF And RSLOAN.EOF) Then
                XTOTAL = RSLOAN!Month1 + RSLOAN!MOnth2 + RSLOAN!Month3
                gconDMIS.Execute ("Update HRMS_LOAN_QUARTERLY Set MONTH2 = " & RSTMP!AMOUNT & _
                                  ",Xtotal = " & XTOTAL & _
                                " Where Empno = '" & RSTMP!EMPNO & "'")
            Else
                gconDMIS.Execute ("Insert Into HRMS_LOAN_QUARTERLY (EMPNO,MONTH1,MONTH2,MONTH3,XTOTAL) VALUES('" & RSTMP!EMPNO & _
                                  "'," & 0 & "," & RSTMP!AMOUNT & "," & 0 & "," & RSTMP!AMOUNT & ")")
            End If

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_LoanMasDet Where Month(DEYT) = " & TM & " And YEAR(deyt) = " & cboyear & " Order By Empno ASC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set RSLOAN = gconDMIS.Execute("Select * From HRMS_LOAN_QUARTERLY Where Empno = '" & RSTMP!EMPNO & "'")
            If Not (RSLOAN.BOF And RSLOAN.EOF) Then
                XTOTAL = RSLOAN!Month1 + RSLOAN!MOnth2 + RSLOAN!Month3
                gconDMIS.Execute ("Update HRMS_LOAN_QUARTERLY Set MONTH3 = " & RSTMP!AMOUNT & _
                                  ",Xtotal = " & XTOTAL & _
                                " Where Empno = '" & RSTMP!EMPNO & "'")
            Else
                gconDMIS.Execute ("Insert Into HRMS_LOAN_QUARTERLY (EMPNO,MONTH1,MONTH2,MONTH3,XTOTAL) VALUES('" & RSTMP!EMPNO & _
                                  "'," & 0 & "," & 0 & "," & RSTMP!AMOUNT & "," & RSTMP!AMOUNT & ")")
            End If

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim X                                                             As Integer

    'If Function_Access(LOGID, "ACESS PRINT", "REPORT QUARTERLY LOANS") = False Then Exit Sub
    rptLOAN.WindowTitle = "Quarterly Loans"
    If cboQuarter.Text = "1ST" Then
        rptLOAN.Formulas(1) = "Quarter = '" & "1ST" & "'"
        rptLOAN.Formulas(2) = "YER = '" & cboyear & "'"
        rptLOAN.Formulas(3) = "FMonth = '" & "JANUARY" & "'"
        rptLOAN.Formulas(4) = "SMonth = '" & "FEBUARY" & "'"
        rptLOAN.Formulas(5) = "TMonth = '" & "MARCH" & "'"
        rptLOAN.Formulas(6) = "CompanyName = '" & COMPANY_NAME & "'"
        rptLOAN.Formulas(7) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptLOAN.Formulas(8) = "PrintedBy = '" & LOGNAME & "'"

        GetLaonRecord 1, 2, 3
        PrintSQLReport rptLOAN, HRMS_REPORT_PATH & "LoanQuarterlyRemit.rpt", "", DMIS_REPORT_Connection, 1

    ElseIf cboQuarter.Text = "2ND" Then
        rptLOAN.Formulas(1) = "Quarter = '" & "2ND" & "'"
        rptLOAN.Formulas(2) = "YER = '" & cboyear & "'"
        rptLOAN.Formulas(3) = "FMonth = '" & "APRIL" & "'"
        rptLOAN.Formulas(4) = "SMonth = '" & "MAY" & "'"
        rptLOAN.Formulas(5) = "TMonth = '" & "JUNE" & "'"
        rptLOAN.Formulas(6) = "CompanyName = '" & COMPANY_NAME & "'"
        rptLOAN.Formulas(7) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptLOAN.Formulas(8) = "PrintedBy = '" & LOGNAME & "'"

        GetLaonRecord 4, 5, 6
        PrintSQLReport rptLOAN, HRMS_REPORT_PATH & "LoanQuarterlyRemit.rpt", "", DMIS_REPORT_Connection, 1

    ElseIf cboQuarter.Text = "3RD" Then
        rptLOAN.Formulas(1) = "Quarter = '" & "3RD" & "'"
        rptLOAN.Formulas(2) = "YER = '" & cboyear & "'"
        rptLOAN.Formulas(3) = "FMonth = '" & "JULY" & "'"
        rptLOAN.Formulas(4) = "SMonth = '" & "AUGUST" & "'"
        rptLOAN.Formulas(5) = "TMonth = '" & "SEPTEMBER" & "'"
        rptLOAN.Formulas(6) = "CompanyName = '" & COMPANY_NAME & "'"
        rptLOAN.Formulas(7) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptLOAN.Formulas(8) = "PrintedBy = '" & LOGNAME & "'"

        GetLaonRecord 7, 8, 9
        PrintSQLReport rptLOAN, HRMS_REPORT_PATH & "LoanQuarterlyRemit.rpt", "", DMIS_REPORT_Connection, 1

    Else
        rptLOAN.Formulas(1) = "Quarter = '" & "4TH" & "'"
        rptLOAN.Formulas(2) = "YER = '" & cboyear & "'"
        rptLOAN.Formulas(3) = "FMonth = '" & "OCTOBER" & "'"
        rptLOAN.Formulas(4) = "SMonth = '" & "NOVEMBER" & "'"
        rptLOAN.Formulas(5) = "TMonth = '" & "DECEMBER" & "'"
        rptLOAN.Formulas(6) = "CompanyName = '" & COMPANY_NAME & "'"
        rptLOAN.Formulas(7) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptLOAN.Formulas(8) = "PrintedBy = '" & LOGNAME & "'"

        GetLaonRecord 10, 11, 12
        PrintSQLReport rptLOAN, HRMS_REPORT_PATH & "LoanQuarterlyRemit.rpt", "", DMIS_REPORT_Connection, 1
    End If

    LogAudit "G", "QUARTERLY LOAN REMMITANCE", cboQuarter & " " & cboyear
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    'FillcboYear cboyear
    fillcombo_up cboyear
    cboyear.ListIndex = 0
    cboyear.Text = YEAR(Now)
    cboQuarter.ListIndex = 0
End Sub


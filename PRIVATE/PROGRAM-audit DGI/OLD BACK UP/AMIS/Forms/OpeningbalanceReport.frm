VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmOpeningBalanceReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening Balance Report"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5040
   ForeColor       =   &H8000000F&
   Icon            =   "OpeningbalanceReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   5040
   Begin VB.OptionButton chkCancelled 
      Caption         =   "Cancelled"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   180
      Width           =   1425
   End
   Begin VB.OptionButton chkUnposted 
      Caption         =   "Un-Posted"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   180
      Width           =   1425
   End
   Begin VB.OptionButton chkPosted 
      Caption         =   "Posted"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   180
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.CheckBox chkBank 
      Caption         =   "Bank Opening Balance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1050
      TabIndex        =   4
      Top             =   1380
      Width           =   2985
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
      Height          =   825
      Left            =   2490
      MouseIcon       =   "OpeningbalanceReport.frx":014A
      MousePointer    =   99  'Custom
      Picture         =   "OpeningbalanceReport.frx":029C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1710
      Width           =   885
   End
   Begin VB.CheckBox ChkVendor 
      Caption         =   "Vendor Opening Balance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1050
      TabIndex        =   3
      Top             =   1080
      Width           =   2985
   End
   Begin VB.CheckBox chkcustomer 
      Caption         =   "Customer Opening Balance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1050
      TabIndex        =   2
      Top             =   780
      Value           =   1  'Checked
      Width           =   2925
   End
   Begin Crystal.CrystalReport OpeningRPT 
      Left            =   4320
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Opening Balance Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
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
      Height          =   825
      Left            =   1620
      MouseIcon       =   "OpeningbalanceReport.frx":06E7
      MousePointer    =   99  'Custom
      Picture         =   "OpeningbalanceReport.frx":0839
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1710
      Width           =   885
   End
End
Attribute VB_Name = "frmOpeningBalanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim J_STATUS As String

Sub printCustomerOpening()
'Update By BTT : 06/20/2008
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    'If Function_Access(LOGID, "Acess_Print", "Cancelled Report") = False Then Exit Sub
 
    SQL = "SELECT * from AMIS_Journal_HD where jtype='COB'order by voucherno asc"
    'SQL = "SELECT * from AMIS_Journal_HD where jtype='COB'order by voucherno asc"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        Screen.MousePointer = 11
        OpeningRPT.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        OpeningRPT.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        OpeningRPT.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
        OpeningRPT.Formulas(3) = "tojdate ='" & dtpTo & "'"
        OpeningRPT.WindowTitle = "Customer Opening Balance Report"
        'PrintSQLReport OpeningRPT, AMIS_REPORT_PATH & "OpeningBalance\CustomerOpening.rpt", "{Amis_journal_HD.JDATE} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Amis_journal_HD.JDATE} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") ", DMIS_REPORT_Connection, 1
        PrintSQLReport OpeningRPT, AMIS_REPORT_PATH & "OpeningBalance\CustomerOpening.rpt", "{AMIS_JOURNAL_HD.STATUS} = '" & J_STATUS & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    LogAudit "V", "Customer Opening Balance Report", dtpFrom & "-" & dtpTo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub printVendorOpening()
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    'If Function_Access(LOGID, "Acess_Print", "Cancelled Report") = False Then Exit Sub
 
    SQL = "SELECT * from AMIS_Journal_HD where jtype='VPJ'order by voucherno asc"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        Screen.MousePointer = 11
        OpeningRPT.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        OpeningRPT.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        OpeningRPT.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
        OpeningRPT.Formulas(3) = "tojdate ='" & dtpTo & "'"
        OpeningRPT.WindowTitle = "Vendor Opening Balance Report"
        PrintSQLReport OpeningRPT, AMIS_REPORT_PATH & "OpeningBalance\VendorOpening.rpt", "{AMIS_JOURNAL_HD.STATUS} = '" & J_STATUS & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    LogAudit "V", "Vendor Opening Balance Report", dtpFrom & "-" & dtpTo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub printBankOpening()
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    'If Function_Access(LOGID, "Acess_Print", "Cancelled Report") = False Then Exit Sub

    SQL = "SELECT * from AMIS_Journal_HD where jtype='BOB' order by voucherno asc"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        Screen.MousePointer = 11
        OpeningRPT.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        OpeningRPT.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        OpeningRPT.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
        OpeningRPT.Formulas(3) = "tojdate ='" & dtpTo & "'"
        OpeningRPT.WindowTitle = "Bank Opening Balance Report"
        PrintSQLReport OpeningRPT, AMIS_REPORT_PATH & "OpeningBalance\BankOpening.rpt", "{AMIS_JOURNAL_HD.STATUS} = '" & J_STATUS & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    LogAudit "V", "Bank Opening Balance Report", dtpFrom & "-" & dtpTo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub chkCancelled_Click()
    J_STATUS = "C"
End Sub

Private Sub chkPosted_Click()
    J_STATUS = "P"
End Sub

Private Sub chkUnposted_Click()
    J_STATUS = "N"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If chkcustomer.Value = 1 Then
        printCustomerOpening
    End If
    If ChkVendor.Value = 1 Then
        printVendorOpening
    End If
    If chkBank.Value = 1 Then
        printBankOpening
    End If
    If chkcustomer.Value = 0 And ChkVendor.Value = 0 And chkBank.Value = 0 Then
        MsgBox "Please select from the option box.", vbInformation, "Warning"
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    Screen.MousePointer = 0
End Sub


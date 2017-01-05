VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPrintATM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print ATM Advice"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3630
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PrintATM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   3630
   Begin SHDocVwCtl.WebBrowser browAtmAdvice 
      Height          =   11640
      Left            =   6480
      TabIndex        =   12
      Top             =   7440
      Width           =   15480
      ExtentX         =   27305
      ExtentY         =   20532
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Copy For President and Corp.Sec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   2565
      Width           =   3405
   End
   Begin VB.OptionButton OptManager 
      Caption         =   "Print Managers ATM Advice"
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
      Left            =   150
      TabIndex        =   6
      Top             =   1725
      Width           =   3315
   End
   Begin VB.OptionButton OptAllowanceBase 
      Caption         =   "Print for Allowance Base Employees"
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
      Left            =   150
      TabIndex        =   5
      Top             =   1485
      Width           =   3315
   End
   Begin VB.OptionButton OptContractual 
      Caption         =   "Print for Contractual Employees"
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
      Left            =   150
      TabIndex        =   4
      Top             =   1245
      Width           =   3315
   End
   Begin VB.OptionButton OptProbReg 
      Caption         =   "Print for Probationary/Regular Employees"
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
      Left            =   150
      TabIndex        =   3
      Top             =   1005
      Value           =   -1  'True
      Width           =   3315
   End
   Begin VB.CheckBox chkCopyCorpSec 
      Caption         =   "Copy For Corporate Secretary Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   2025
      Width           =   3405
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Copy For Treasurer Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   2295
      Width           =   3405
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
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   2775
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
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   555
      Width           =   1845
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
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   555
      Width           =   885
   End
   Begin Crystal.CrystalReport rptPrintATM 
      Left            =   2910
      Top             =   3135
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
      Height          =   735
      Left            =   1770
      MouseIcon       =   "PrintATM.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "PrintATM.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Close Window"
      Top             =   2940
      Width           =   795
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
      Height          =   735
      Left            =   990
      MouseIcon       =   "PrintATM.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "PrintATM.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Print Report"
      Top             =   2940
      Width           =   795
   End
End
Attribute VB_Name = "frmHRMSPrintATM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsATMdet, rsEmpInfo                                               As ADODB.Recordset
Attribute rsEmpInfo.VB_VarUserMemId = 1073938432
Dim rsEmpInfoDup, rsSignatories                                       As ADODB.Recordset
Attribute rsEmpInfoDup.VB_VarUserMemId = 1073938434
Attribute rsSignatories.VB_VarUserMemId = 1073938434
Dim FromDate, ToDate                                                  As String
Attribute FromDate.VB_VarUserMemId = 1073938436
Attribute ToDate.VB_VarUserMemId = 1073938436
Dim SBMan, SBAcctNo                                                   As String
Attribute SBMan.VB_VarUserMemId = 1073938438
Attribute SBAcctNo.VB_VarUserMemId = 1073938438
Dim GManager, President                                               As String
Attribute GManager.VB_VarUserMemId = 1073938440
Attribute President.VB_VarUserMemId = 1073938440
Dim CorpSec, Treasurer                                                As String
Attribute CorpSec.VB_VarUserMemId = 1073938442
Attribute Treasurer.VB_VarUserMemId = 1073938442

Function FindPrevMonth() As String
    If cboMOnth.Text = "January" Then FindPrevMonth = "12"
    If cboMOnth.Text = "Febuary" Then FindPrevMonth = "1"
    If cboMOnth.Text = "March" Then FindPrevMonth = "2"
    If cboMOnth.Text = "April" Then FindPrevMonth = "3"
    If cboMOnth.Text = "May" Then FindPrevMonth = "4"
    If cboMOnth.Text = "June" Then FindPrevMonth = "5"
    If cboMOnth.Text = "July" Then FindPrevMonth = "6"
    If cboMOnth.Text = "August" Then FindPrevMonth = "7"
    If cboMOnth.Text = "September" Then FindPrevMonth = "8"
    If cboMOnth.Text = "October" Then FindPrevMonth = "9"
    If cboMOnth.Text = "November" Then FindPrevMonth = "10"
    If cboMOnth.Text = "December" Then FindPrevMonth = "11"
End Function

Function FindNextMonth() As String
    If cboMOnth.Text = "January" Then FindNextMonth = "2"
    If cboMOnth.Text = "Febuary" Then FindNextMonth = "3"
    If cboMOnth.Text = "March" Then FindNextMonth = "4"
    If cboMOnth.Text = "April" Then FindNextMonth = "5"
    If cboMOnth.Text = "May" Then FindNextMonth = "6"
    If cboMOnth.Text = "June" Then FindNextMonth = "7"
    If cboMOnth.Text = "July" Then FindNextMonth = "8"
    If cboMOnth.Text = "August" Then FindNextMonth = "9"
    If cboMOnth.Text = "September" Then FindNextMonth = "10"
    If cboMOnth.Text = "October" Then FindNextMonth = "11"
    If cboMOnth.Text = "November" Then FindNextMonth = "12"
    If cboMOnth.Text = "December" Then FindNextMonth = "1"
End Function

Sub PrintATM()

    On Error Resume Next
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & COMPANY_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Payroll Period From <b>" & Format(FromDate, "DD-MMM-YYYY") & "</b> To <b>" & Format(ToDate, "DD-MMM-YYYY") & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=10%><font size = 2>" & Day(Date) & " " & The_month(MONTH(LOGDATE)) & " " & YEAR(LOGDATE) & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2><b>" & SBMan & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>  </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2></td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2></td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    rsATMdet.MoveFirst

    Do While Not rsATMdet.EOF
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive='A' and emplevel = " & N2Str2Null(rsATMdet!EMPLEVEL) & " and empno = '" & rsATMdet!EMPNO & "'", gconDMIS
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            EmpNetAmt = Format(N2Str2Zero(rsATMdet!netamount), "###,###,##0.00")
            TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
        End If
        rsATMdet.MoveNext
    Loop
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & SBAcctNo & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, "###,###,##0.00") & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    rsATMdet.MoveFirst
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Do While Not rsATMdet.EOF
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive='A' and emplevel = " & N2Str2Null(rsATMdet!EMPLEVEL) & " and empno = '" & rsATMdet!EMPNO & "'", gconDMIS
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            I = I + 1
            NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
            EmpActNo = Null2String(rsATMdet!acctno)
            EmpNetAmt = N2Str2Zero(rsATMdet!netamount)
            TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            Print #1, "<tr ALIGN = RIGHT>"
            Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
            Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
            Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
            Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, "###,###,##0.00") & "</td>"
            Print #1, "</tr>"
        End If
        rsATMdet.MoveNext
    Loop
    Print #1, "</table>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = right width=10%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, "###,###,##0.00") & "</u></b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = right width=10%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 3></td>"
    Print #1, "</tr>"
    Print #1, "</table><p><br><br></p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2>APPROVED FOR PAYMENT</td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2>NOTED BY:</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2><br></td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3></td>"
    Print #1, "<td align = left width=30%><font size = 3></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2><b>" & GManager & "</B></td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2><b>" & President & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3>General Manager</td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3>President</td>"
    Print #1, "</tr>"
    Print #1, "</table><p><br><br></p>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1
    DoEvents
    On Error GoTo Errorcode
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
    Else
        Close #1
        '        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        'browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    '    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintCorpATM()
    On Error Resume Next
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double

    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & CompanyName & "</td>"
    'Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Payroll Period From <b>" & Format(FromDate, "DD-MMM-YYYY") & "</b> To <b>" & Format(ToDate, "DD-MMM-YYYY") & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=10%><font size = 2>" & Day(Date) & " " & The_month(MONTH(LOGDATE)) & " " & YEAR(LOGDATE) & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2><b>" & SBMan & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2> </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2> </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2></td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    rsATMdet.MoveFirst
    Do While Not rsATMdet.EOF
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive='A' and emplevel = " & N2Str2Null(rsATMdet!EMPLEVEL) & " and empno = '" & rsATMdet!EMPNO & "'", gconDMIS
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            EmpNetAmt = Format(N2Str2Zero(rsATMdet!netamount), "###,###,##0.00")
            TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
        End If
        rsATMdet.MoveNext
    Loop
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & SBAcctNo & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, "###,###,##0.00") & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    rsATMdet.MoveFirst
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Do While Not rsATMdet.EOF
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive='A' and emplevel = " & N2Str2Null(rsATMdet!EMPLEVEL) & " and empno = '" & rsATMdet!EMPNO & "'", gconDMIS
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            I = I + 1
            NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
            EmpActNo = Null2String(rsATMdet!acctno)
            EmpNetAmt = N2Str2Zero(rsATMdet!netamount)
            TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            Print #1, "<tr ALIGN = RIGHT>"
            Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
            Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
            Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
            Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, "###,###,##0.00") & "</td>"
            Print #1, "</tr>"
        End If
        rsATMdet.MoveNext
    Loop
    Print #1, "</table>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = right width=10%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, "###,###,##0.00") & "</u></b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = right width=10%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 3></td>"
    Print #1, "</tr>"
    Print #1, "</table><p><br><br></p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2>APPROVED FOR PAYMENT</td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2>NOTED BY:</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2><br></td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3></td>"
    Print #1, "<td align = left width=30%><font size = 3></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2><b>" & GManager & "</B></td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2><b>" & CorpSec & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3>General Manager</td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3>Corporate Secretary</td>"
    Print #1, "</tr>"
    Print #1, "</table><p><br><br></p>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1
    DoEvents
    On Error GoTo Errorcode
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        '       DoEvents

        'browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        '       Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintPresCorpATM()
    On Error Resume Next
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double

    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & COMPANY_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Payroll Period From <b>" & Format(FromDate, "DD-MMM-YYYY") & "</b> To <b>" & Format(ToDate, "DD-MMM-YYYY") & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=10%><font size = 2>" & Day(Date) & " " & The_month(MONTH(LOGDATE)) & " " & YEAR(LOGDATE) & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2><b>" & SBMan & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2></td>"
    Print #1, "<td align = LEFT width=30%><font size = 2> </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2> </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2> </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    rsATMdet.MoveFirst
    Do While Not rsATMdet.EOF
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive='A' and emplevel = " & N2Str2Null(rsATMdet!EMPLEVEL) & " and empno = '" & rsATMdet!EMPNO & "'", gconDMIS
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            EmpNetAmt = Format(N2Str2Zero(rsATMdet!netamount), "###,###,##0.00")
            TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
        End If
        rsATMdet.MoveNext
    Loop
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & SBAcctNo & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, "###,###,##0.00") & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    rsATMdet.MoveFirst
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Do While Not rsATMdet.EOF
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive='A' and emplevel = " & N2Str2Null(rsATMdet!EMPLEVEL) & " and empno = '" & rsATMdet!EMPNO & "'", gconDMIS
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            I = I + 1
            NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
            EmpActNo = Null2String(rsATMdet!acctno)
            EmpNetAmt = N2Str2Zero(rsATMdet!netamount)
            TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            Print #1, "<tr ALIGN = RIGHT>"
            Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
            Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
            Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
            Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, "###,###,##0.00") & "</td>"
            Print #1, "</tr>"
        End If
        rsATMdet.MoveNext
    Loop
    Print #1, "</table>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = right width=10%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, "###,###,##0.00") & "</u></b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = right width=10%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 3></td>"
    Print #1, "</tr>"
    Print #1, "</table><p><br><br></p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2>APPROVED FOR PAYMENT</td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2>NOTED BY:</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2><br></td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3></td>"
    Print #1, "<td align = left width=30%><font size = 3></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2><b>" & President & "</B></td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2><b>" & CorpSec & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3>President</td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3>Corporate Secretary</td>"
    Print #1, "</tr>"
    Print #1, "</table><p><br><br></p>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1
    DoEvents
    On Error GoTo Errorcode
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        'browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintTreasATM()
    'On Error Resume Next
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double

    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & COMPANY_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Payroll Period From <b>" & Format(FromDate, "DD-MMM-YYYY") & "</b> To <b>" & Format(ToDate, "DD-MMM-YYYY") & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=10%><font size = 2>" & Day(Date) & " " & The_month(MONTH(LOGDATE)) & " " & YEAR(LOGDATE) & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2><b>" & SBMan & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2> </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2> </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2> </td>"
    'Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    rsATMdet.MoveFirst
    Do While Not rsATMdet.EOF
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive='A' and emplevel = " & N2Str2Null(rsATMdet!EMPLEVEL) & " and empno = '" & rsATMdet!EMPNO & "'", gconDMIS
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            EmpNetAmt = Format(N2Str2Zero(rsATMdet!netamount), "###,###,##0.00")
            TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
        End If
        rsATMdet.MoveNext
    Loop
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & SBAcctNo & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, "###,###,##0.00") & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    rsATMdet.MoveFirst
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Do While Not rsATMdet.EOF
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive='A' and emplevel = " & N2Str2Null(rsATMdet!EMPLEVEL) & " and empno = '" & rsATMdet!EMPNO & "'", gconDMIS
        If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
            I = I + 1
            NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
            EmpActNo = Null2String(rsATMdet!acctno)
            EmpNetAmt = N2Str2Zero(rsATMdet!netamount)
            TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            Print #1, "<tr ALIGN = RIGHT>"
            Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
            Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
            Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
            Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, "###,###,##0.00") & "</td>"
            Print #1, "</tr>"
        End If
        rsATMdet.MoveNext
    Loop
    Print #1, "</table>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = right width=10%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, "###,###,##0.00") & "</u></b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = right width=10%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 2></td>"
    Print #1, "<td align = right width=30%><font size = 3></td>"
    Print #1, "</tr>"
    Print #1, "</table><p><br><br></p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2>APPROVED FOR PAYMENT</td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2>NOTED BY:</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2><br></td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3></td>"
    Print #1, "<td align = left width=30%><font size = 3></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2><b>" & GManager & "</B></td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 2><b>" & Treasurer & "</b></td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = left>"
    Print #1, "<td align = left width=10%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3>General Manager</td>"
    Print #1, "<td align = left width=30%><font size = 2></td>"
    Print #1, "<td align = left width=30%><font size = 3>Treasurer</td>"
    Print #1, "</tr>"
    Print #1, "</table><p><br><br></p>"
    Print #1, "</body>"
    Print #1, "</html>"
    Close #1
    DoEvents
    'On Error GoTo ErrorCode

    Open App.Path & "\ATM.html" For Input As #1

    If EOF(1) Then
        Screen.MousePointer = 0
        MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        'browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

    'ErrorCode:
    '   browATMAdvice.Refresh
    '   Resume Next
End Sub

Private Sub Check1_Click()
    If chkCopyCorpSec.Value = 1 Then
        chkCopyCorpSec.Value = 0
    End If
End Sub

Private Sub chkCopyCorpSec_Click()
    If Check1.Value = 1 Then
        Check1.Value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    Dim MM, ddFROM, YY                                                As String
    Dim vYEAR                                                         As String
    MM = What_month(cboMOnth)
    YY = cboyear.Text
    If cboQuensina.Text = "1st Cut-Off" Then
        If PAYROLL_CODE = 1 Then
            FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM1)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO1)
        Else
            If cboMOnth.Text = "January" Then
                vYEAR = CDbl(cboyear) - 1
            Else
                vYEAR = CDbl(cboyear)
            End If
            FromDate = DateSerial(vYEAR, FindPrevMonth, PAYROLLCODE_FROM1)
            ToDate = DateSerial(vYEAR, MM, PAYROLLCODE_TO1)
        End If
    Else
        FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM2)
        ToDate = DateSerial(YY, MM, PAYROLLCODE_TO2)
    End If

    Dim FILTER                                                        As String
    Set rsATMdet = New ADODB.Recordset
    If OptProbReg.Value = True Then
        rsATMdet.Open "select * from HRMS_ATMdet where emplevel = 'E' AND (deyt = '" & Format(ToDate, "Short Date") & "') order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf OptContractual.Value = True Then
        rsATMdet.Open "select * from HRMS_ATMdet where emplevel = 'C' AND (deyt = '" & Format(ToDate, "Short Date") & "') order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf OptAllowanceBase.Value = True Then
        rsATMdet.Open "select * from HRMS_ATMdet where emplevel = 'A' AND (deyt = '" & Format(ToDate, "Short Date") & "') order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsATMdet.Open "select * from HRMS_ATMdet where emplevel = 'M' AND (deyt = '" & Format(ToDate, "Short Date") & "') order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsATMdet.EOF And Not rsATMdet.BOF Then
        If chkCopyCorpSec.Value = 1 Then
            PrintCorpATM
        ElseIf Check1.Value = 1 Then
            PrintTreasATM
        ElseIf Check2.Value = 1 Then
            PrintPresCorpATM
        Else
            PrintATM
        End If
    Else
        ShowNoRecord
    End If
    Screen.MousePointer = 0
    On Error GoTo adder:
    Dim C                                                             As Object
    Set C = CreateObject("internetexplorer.application")
    C.Navigate (App.Path & " /ATM.html")
    C.ToolBar = False
    C.StatusBar = False
    C.Width = Me.ScaleWidth
    C.HEIGHT = Me.ScaleHeight
    C.Left = 0
    C.Top = 0
    C.Visible = True
    C.ExecWB 7, 1
    Set C = Nothing
    Exit Sub
adder:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub Command1_Click()
    'While Not browATMAdvice.ReadyState = READYSTATE_COMPLETE
    '        Debug.Print "NO"
    '    Wend
    'Dim cx As InternetExplorer
    'cx.ExecWB OLECMDID_PRINTPREVIEW,OLECMDEXECOPT_PROMPTUSER
    'cx.ExecWB OLECMDID_PRINT,OLECMDEXECOPT_DODEFAULT
    '    browATMAdvice.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    'On Error Resume Next
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If LOGLEVEL <> "ADM" Then
        OptManager.Enabled = False
    End If
    Set rsSignatories = New ADODB.Recordset
    rsSignatories.Open "select * from ALL_PROFILE Where ModuleName = 'HRMS'", gconDMIS
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        SBMan = Null2String(rsSignatories!BankManager)
        SBAcctNo = Null2String(rsSignatories!ACCOUNTNO)
        GManager = Null2String(rsSignatories!APPROVEDBY)
        President = Null2String(rsSignatories!NotedBy1)
        CorpSec = Null2String(rsSignatories!NotedBy1)
        Treasurer = Null2String(rsSignatories!NOTEDBY2)
    End If
    cboQuensina.AddItem "1st Cut-Off"
    cboQuensina.AddItem "2nd Cut-Off"
    fillcbomonth cboMOnth
    'FillcboYear cboYear
    fillcombo_up cboyear
    
    If Day(Date) > 15 Then
        cboQuensina.Text = "2nd Cut-Off"
    Else
        cboQuensina.Text = "1st Cut-Off"
    End If
    cboyear.Text = YEAR(LOGDATE)
    cboMOnth.Text = The_month(MONTH(LOGDATE))

    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHRMSPrintATM = Nothing
End Sub


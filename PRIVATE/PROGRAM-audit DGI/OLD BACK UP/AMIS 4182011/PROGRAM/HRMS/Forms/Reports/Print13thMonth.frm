VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPrint13thMonth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print 13th Month Pay"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4080
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Print13thMonth.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   4080
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
      Left            =   1845
      MouseIcon       =   "Print13thMonth.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "Print13thMonth.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Close Window"
      Top             =   3015
      Width           =   885
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
      Left            =   975
      MouseIcon       =   "Print13thMonth.frx":09DF
      MousePointer    =   99  'Custom
      Picture         =   "Print13thMonth.frx":0B31
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print Report"
      Top             =   3015
      Width           =   885
   End
   Begin VB.CheckBox chkContractuals 
      Caption         =   "Print Contractuals Copy"
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
      Height          =   285
      Left            =   870
      TabIndex        =   9
      Top             =   2370
      Width           =   2835
   End
   Begin VB.CheckBox chkBonus 
      Caption         =   "Print Bonus"
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
      Height          =   285
      Left            =   870
      TabIndex        =   8
      Top             =   2100
      Width           =   2835
   End
   Begin VB.CheckBox chkManagers 
      Caption         =   "Print Managers Copy"
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
      Height          =   285
      Left            =   870
      TabIndex        =   10
      Top             =   2640
      Width           =   2835
   End
   Begin SHDocVwCtl.WebBrowser browAtmAdvice 
      Height          =   3495
      Left            =   600
      TabIndex        =   11
      Top             =   4770
      Width           =   7035
      ExtentX         =   12409
      ExtentY         =   6165
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
   Begin VB.CheckBox chkMidYear 
      Caption         =   "Mid Year Pay Only"
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
      Height          =   285
      Left            =   870
      TabIndex        =   3
      Top             =   750
      Width           =   2835
   End
   Begin VB.CheckBox chkPayslip 
      Caption         =   "Print Payslip"
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
      Height          =   285
      Left            =   870
      TabIndex        =   7
      Top             =   1830
      Width           =   2835
   End
   Begin VB.CheckBox chkCorpSec 
      Caption         =   "Another Sheet for Corp. Sec."
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
      Height          =   285
      Left            =   870
      TabIndex        =   6
      Top             =   1560
      Width           =   2835
   End
   Begin VB.CheckBox chkATMAdvice 
      Caption         =   "Print ATM Advice"
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
      Height          =   285
      Left            =   870
      TabIndex        =   5
      Top             =   1290
      Width           =   2835
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
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
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
      Left            =   3090
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   945
   End
   Begin VB.ComboBox cboDay 
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
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1065
   End
   Begin Crystal.CrystalReport rpt13thMonth 
      Left            =   90
      Top             =   1530
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
   Begin VB.CheckBox chkInclude 
      Caption         =   "Include Resignees"
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
      Height          =   285
      Left            =   870
      TabIndex        =   4
      Top             =   1020
      Width           =   2835
   End
End
Attribute VB_Name = "frmHRMSPrint13thMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsYTDDETAILS, rsSignatories                                       As ADODB.Recordset
Attribute rsSignatories.VB_VarUserMemId = 1073938432
Dim ToDate                                                            As String
Attribute ToDate.VB_VarUserMemId = 1073938434
Dim SBMan, SBAcctNo                                                   As String
Attribute SBMan.VB_VarUserMemId = 1073938435
Attribute SBAcctNo.VB_VarUserMemId = 1073938435
Dim GManager, President                                               As String
Attribute GManager.VB_VarUserMemId = 1073938437
Attribute President.VB_VarUserMemId = 1073938437
Dim CorpSec, Treasurer                                                As String
Attribute CorpSec.VB_VarUserMemId = 1073938439
Attribute Treasurer.VB_VarUserMemId = 1073938439

Sub PrintATM()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>13th Month Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,t13thmonth from HRMS_YTDDetails where emplevel = 'E' and yeer = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A' and emplevel ='E'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive = 'A' and emplevel = 'E' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select t13thmonth,empno from HRMS_YTDDetails where emplevel = 'E' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMManager()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>13th Month Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,t13thmonth from HRMS_YTDDetails where emplevel = 'M' and yeer = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A' and emplevel = 'M'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive = 'A' and emplevel = 'M' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select t13thmonth,empno from HRMS_YTDDetails where emplevel = 'M' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMContractuals()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>13th Month Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,t13thmonth from HRMS_YTDDetails where emplevel = 'C' and yeer = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A' and emplevel = 'C'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive = 'A' and emplevel = 'C' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select t13thmonth,empno from HRMS_YTDDetails where emplevel = 'C' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMMidYearCorpSec()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Mid Year Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,midyear from HRMS_YTDDetails where emplevel = 'E' and YEER = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'E' and empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'E' and activeinactive = 'A' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select midyear,empno from HRMS_YTDDetails where emplevel = 'E' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMMidYear()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Mid Year Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,midyear from HRMS_YTDDetails where emplevel = 'E' and YEER = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'E' and empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'E' and activeinactive = 'A' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select midyear,empno from HRMS_YTDDetails where emplevel = 'E' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMManagersMidYearCorpSec()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATMMan.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>(MANAGERS MID-YEAR) ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Mid Year Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,midyear from HRMS_YTDDetails where emplevel = 'M' and YEER = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'M' and empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'M' and activeinactive = 'A' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select midyear,empno from HRMS_YTDDetails where emplevel = 'M' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATMMan.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATMMan.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMManagersMidYear()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATMMan.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>(MANAGERS MID-YEAR) ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Mid Year Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,midyear from HRMS_YTDDetails where emplevel = 'M' and YEER = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'M' and empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'M' and activeinactive = 'A' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select midyear,empno from HRMS_YTDDetails where emplevel = 'M' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATMMan.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATMMan.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMContractualsMidYearCorpSec()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATMCon.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>(CONTRACTUALS MID-YEAR) ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Mid Year Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,midyear from HRMS_YTDDetails where emplevel = 'C' and YEER = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'C' and empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'C' and activeinactive = 'A' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select midyear,empno from HRMS_YTDDetails where emplevel = 'C' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATMCon.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATMCon.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMContractualsMidYear()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATMCon.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>(CONTRACTUALS MID-YEAR) ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Mid Year Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,midyear from HRMS_YTDDetails where emplevel = 'C' and YEER = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'C' and empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'C' and activeinactive = 'A' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select midyear,empno from HRMS_YTDDetails where emplevel = 'C' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!midyear), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATMCon.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATMCon.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMBonus()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Christmas Bonus Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,Bonus from HRMS_YTDDetails where emplevel = 'E' and yeer = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A' and emplevel='E'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!bonus), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive = 'A' and emplevel = 'E' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select Bonus,empno from HRMS_YTDDetails where emplevel = 'E' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!bonus), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMManagerBonus()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Christmas Bonus Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,Bonus from HRMS_YTDDetails where emplevel = 'M' and yeer = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'M' and empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!bonus), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive = 'A' and emplevel = 'M' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select Bonus,empno from HRMS_YTDDetails where emplevel = 'M' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!bonus), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Sub PrintATMContractualBonus()
    Dim I                                                             As Integer
    Dim NeymEmp, EmpActNo                                             As String
    Dim EmpNetAmt, TotEmpNetAmt                                       As Double
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Close #1
    Open App.Path & "\ATM.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<html>"
    Print #1, "<head><title> ATM Advice </title>"
    Print #1, "</head>"
    Print #1, "<body><p><br><br><br></p>"
    Print #1, "<table alignment = center border= 0cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=10%><font size = 4>" & SYSTEM_OWNER_NAME & "</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 4>ATM ADVICE</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = center>"
    Print #1, "<td align = center width=30%><font size = 2>Christmas Bonus Release Date <b>" & Format(LOGDATE, "DD-MMM-YYYY") & "</b></td>"
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
    Print #1, "<td align = LEFT width=30%><font size = 2>EQUITABLE PCI BANK</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Caceres Naga Branch</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>Naga City</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Dim rsYTDDETAILS                                                  As ADODB.Recordset
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select empno,Bonus from HRMS_YTDDetails where emplevel = 'C' and yeer = '" & cboyear.Text & "' order by empno asc", gconDMIS
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        Do While Not rsYTDDETAILS.EOF
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = 'C' and empno = '" & rsYTDDETAILS!EMPNO & "' and activeinactive = 'A'", gconDMIS
            If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!bonus), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
            End If
            rsYTDDETAILS.MoveNext
        Loop
    End If
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>This is to authorize your bank to debit from account no. <B><U>" & ACCOUNT_NO & "</U></B> the amount of <b><u>Php " & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b> to be credited</td>"
    Print #1, "</tr>"
    Print #1, "<tr ALIGN = LEFT>"
    Print #1, "<td align = LEFT width=30%><font size = 2>to the following accounts:</td>"
    Print #1, "</tr></table><p>"
    Print #1, "<table alignment = right border=1 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
    Print #1, "<tr ALIGN = RIGHT>"
    Print #1, "<td align = center width=10%><font size = 2></td>"
    Print #1, "<td align = center width=30%><font size = 2>NAME</td>"
    Print #1, "<td align = center width=30%><font size = 2>ACCOUNT NUMBER</td>"
    Print #1, "<td align = center width=30%><font size = 2>NET AMOUNT</td>"
    Print #1, "</tr>"
    I = 0
    TotEmpNetAmt = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where activeinactive = 'A' and emplevel = 'C' order by lastname asc", gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Do While Not rsEmpInfo.EOF
            Set rsYTDDETAILS = New ADODB.Recordset
            rsYTDDETAILS.Open "select Bonus,empno from HRMS_YTDDetails where emplevel = 'C' and empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboyear.Text & "'", gconDMIS
            If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                I = I + 1
                NeymEmp = Null2String(rsEmpInfo!lastname) + ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
                EmpActNo = Null2String(rsEmpInfo!ACCOUNTNO)
                EmpNetAmt = Format(N2Str2Zero(rsYTDDETAILS!bonus), MAXIMUM_DIGIT)
                TotEmpNetAmt = TotEmpNetAmt + EmpNetAmt
                Print #1, "<tr ALIGN = RIGHT>"
                Print #1, "<td align = center width=10%><font size = 2>" & I & "</td>"
                Print #1, "<td align = left width=30%><font size = 2>" & NeymEmp & "</td>"
                Print #1, "<td align = center width=30%><font size = 2>" & EmpActNo & "</td>"
                Print #1, "<td align = right width=30%><font size = 2>" & Format(EmpNetAmt, MAXIMUM_DIGIT) & "</td>"
                Print #1, "</tr>"
            End If
            rsEmpInfo.MoveNext
        Loop
        Print #1, "</table>"
        Print #1, "<table alignment = right border=0 cellpadding=0 cellspacing=0 bordercolor=#000000 width=100% id=AutoNumber1>"
        Print #1, "<tr ALIGN = RIGHT>"
        Print #1, "<td align = right width=10%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 2></td>"
        Print #1, "<td align = right width=30%><font size = 3><b><u>" & Format(TotEmpNetAmt, MAXIMUM_DIGIT) & "</u></b></td>"
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
    End If
    DoEvents
    On Error Resume Next
    Open App.Path & "\ATM.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgSpeechBox "File Not Found!"
    Else
        Close #1
        browAtmAdvice.Navigate App.Path & "\ATM.html"
        DoEvents
        browAtmAdvice.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    browAtmAdvice.Refresh
    Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:47
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "Acess_Print", "13TH MONTH PAY") = False Then Exit Sub

    Dim MM, DD, YY, FILTER                                            As String
    YY = cboyear.Text
    MM = What_month(cboMOnth.Text)
    DD = cboDay.Text
    ToDate = DateSerial(YY, MM, DD)
    rpt13thMonth.WindowState = crptMaximized

    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select * from HRMS_YTDDetails where (ytdCutoffdate = '" & ToDate & "') order by empno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        Screen.MousePointer = 11
        rpt13thMonth.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
        rpt13thMonth.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
        rpt13thMonth.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
        If chkPaySlip.Value = 0 Then
            rpt13thMonth.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
            rpt13thMonth.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
            rpt13thMonth.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
        End If
        If chkMidYear.Value = 1 Then
            If chkCorpSec.Value = 1 Then
                If chkATMAdvice.Value = 1 Then
                    If chkContractuals.Value = 0 And chkManagers.Value = 0 Then
                        PrintATMMidYearCorpSec
                    End If
                    If chkContractuals.Value = 1 Then
                        PrintATMContractualsMidYearCorpSec
                    End If
                    If chkManagers.Value = 1 Then
                        If LOGLEVEL = "ADM" Then PrintATMManagersMidYearCorpSec
                    End If
                End If
            Else
                If chkATMAdvice.Value = 1 Then
                    If chkContractuals.Value = 0 And chkManagers.Value = 0 Then
                        PrintATMMidYear
                    End If
                    If chkContractuals.Value = 1 Then
                        PrintATMContractualsMidYear
                    End If
                    If chkManagers.Value = 1 Then
                        If LOGLEVEL = "ADM" Then PrintATMManagersMidYear
                    End If
                Else
                    If chkPaySlip.Value = 1 Then
                        PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "PayslipMidyear.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                    Else
                        PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "midyearA.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                        If chkContractuals.Value = 1 Then
                            PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "midyearC.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                        End If
                        If chkManagers.Value = 1 Then
                            If LOGLEVEL = "ADM" Then PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "midyearM.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                        End If
                    End If
                End If
            End If
        ElseIf chkInclude.Value = 1 Then
            PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "13thmonthNA.rpt", "{ytddetails.yeer} = '" & YY & "'", DMIS_REPORT_Connection, 1
        ElseIf chkCorpSec.Value = 1 Then
            PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "ATM13thCorp.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
        ElseIf chkManagers.Value = 1 Then
            If chkBonus.Value = 1 Then
                If chkATMAdvice.Value = 1 Then
                    PrintATMManagerBonus
                Else
                    If chkPaySlip.Value = 1 Then
                        PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "PayslipBonusM.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                    Else
                        PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "BonusM.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                    End If
                End If
            ElseIf chkATMAdvice.Value = 1 Then
                PrintATMManager
            Else
                If chkPaySlip.Value = 1 Then
                    PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "Payslip13thM.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                Else
                    PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "13thmonthM.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                End If
            End If
        ElseIf chkContractuals.Value = 1 Then
            If chkBonus.Value = 1 Then
                If chkATMAdvice.Value = 1 Then
                    PrintATMContractualBonus
                Else
                    If chkPaySlip.Value = 1 Then
                        PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "PayslipBonusC.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                    Else
                        PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "BonusC.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                    End If
                End If
            ElseIf chkATMAdvice.Value = 1 Then
                PrintATMContractuals
            Else
                If chkPaySlip.Value = 1 Then
                    PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "Payslip13thC.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                Else
                    PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "13thmonthC.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                End If
            End If
        ElseIf chkBonus.Value = 1 Then
            If chkATMAdvice.Value = 1 Then
                PrintATMBonus
            Else
                If chkPaySlip.Value = 1 Then
                    PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "PayslipBonus.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                Else
                    PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "Bonus.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                End If
            End If
        Else
            If chkATMAdvice.Value = 1 Then
                PrintATM
            Else
                If chkPaySlip.Value = 1 Then
                    PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "Payslip13th.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                Else
                    PrintSQLReport rpt13thMonth, HRMS_REPORT_PATH & "13thmonthA.rpt", "month({ytddetails.ytdcutoffdate}) = " & MONTH(rsYTDDETAILS!ytdcutoffdate) & " AND day({ytddetails.ytdcutoffdate}) = " & Day(rsYTDDETAILS!ytdcutoffdate) & " AND year({ytddetails.ytdcutoffdate}) = " & YEAR(rsYTDDETAILS!ytdcutoffdate), DMIS_REPORT_Connection, 1
                End If
            End If
        End If
        Screen.MousePointer = 0
    Else
        ShowNoRecord
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
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    fillcboDay cboDay
    fillcbomonth cboMOnth
    'FillcboYear cboyear
    fillcombo_up cboyear
    If LOGLEVEL <> "ADM" Then chkManagers.Enabled = False
    Set rsSignatories = New ADODB.Recordset
    rsSignatories.Open "select * from ALL_PROFILE", gconDMIS
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        SBMan = Null2String(rsSignatories!BankManager)
        SBAcctNo = Null2String(rsSignatories!ACCOUNTNO)
        GManager = Null2String(rsSignatories!APPROVEDBY)
        President = Null2String(rsSignatories!NotedBy1)
        CorpSec = Null2String(rsSignatories!NotedBy1)
        Treasurer = Null2String(rsSignatories!NOTEDBY2)
    Else
        SBMan = "MS. AVA JEAN BRUTAS"
        SBAcctNo = "0493-020677-001"
        GManager = "Sotero E. Dionisio"
        President = "Cresencio Fernandez"
        CorpSec = "Jose Llanera"
        Treasurer = "Bonifacio Sison"
    End If
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select * from HRMS_YTDDetails order by ytdcutoffdate desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        rsYTDDETAILS.MoveFirst
        cboDay.Text = Day(rsYTDDETAILS!ytdcutoffdate)
        cboMOnth.Text = The_month(MONTH(rsYTDDETAILS!ytdcutoffdate))
        cboyear.Text = YEAR(rsYTDDETAILS!ytdcutoffdate)
    Else
        cboDay.Text = Day(LOGDATE)
        cboMOnth.Text = The_month(MONTH(LOGDATE))
        cboyear.Text = YEAR(LOGDATE)
    End If
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub


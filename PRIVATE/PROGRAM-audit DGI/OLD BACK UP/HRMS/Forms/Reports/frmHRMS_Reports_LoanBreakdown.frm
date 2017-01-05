VERSION 5.00
Begin VB.Form frmHRMS_Reports_LaonBreakdown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Loans Breakdown"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_Reports_LoanBreakdown.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   4065
   Begin VB.ComboBox cboName 
      Height          =   330
      Left            =   1455
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   510
      Width           =   2535
   End
   Begin VB.ComboBox cboYear 
      Height          =   330
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1230
      Width           =   1695
   End
   Begin VB.ComboBox cboMOnth 
      Height          =   330
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   870
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Summary"
      Height          =   285
      Left            =   2460
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Per Employee"
      Height          =   285
      Left            =   210
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   2055
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
      Height          =   765
      Left            =   3300
      MouseIcon       =   "frmHRMS_Reports_LoanBreakdown.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Reports_LoanBreakdown.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1650
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
      Height          =   765
      Left            =   2640
      MouseIcon       =   "frmHRMS_Reports_LoanBreakdown.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Reports_LoanBreakdown.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1650
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE NAME"
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONTH"
      Height          =   210
      Index           =   1
      Left            =   870
      TabIndex        =   5
      Top             =   930
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      Height          =   210
      Index           =   0
      Left            =   975
      TabIndex        =   4
      Top             =   1320
      Width           =   435
   End
End
Attribute VB_Name = "frmHRMS_Reports_LaonBreakdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                           As Excel.Application
Dim xlbook                                          As Excel.Workbook
Dim xlsheet                                         As Excel.Worksheet

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    If Option1.Value = True Then
        Call PRINTINEXCEL_EMP
    Else
        Call PRINTEXCEL_DEPT
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call FillEmployeeName
    Call fillcbomonth(cboMOnth)
    cboMOnth.Text = MonthName(MONTH(Date))
    'Call FillcboYear(cboyear)
    Call fillcombo_up(cboyear)
    
End Sub

Sub FillEmployeeName()
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A' AND RESIGNED IS NULL ORDER BY LASTNAME")
    cboName.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboName.AddItem Null2String(RSTMP!FULLNAME)
            RSTMP.MoveNext
        Loop
        cboName.ListIndex = 0
    End If
    Set RSTMP = Nothing
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        cboName.Visible = True
        Label1(2).Visible = True
    Else
        cboName.Visible = False
        Label1(2).Visible = False
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        cboName.Visible = False
        Label1(2).Visible = False
    Else
        cboName.Visible = True
        Label1(2).Visible = True
    End If
End Sub

Sub PRINTEXCEL_DEPT()
    Dim RSTMP                                       As New ADODB.Recordset
    Dim RSDEPT                                      As New ADODB.Recordset
    Dim RSLOAN                                      As New ADODB.Recordset
    Dim RSDET                                       As New ADODB.Recordset
    Dim SQL                                         As String
    Dim xcnt                                        As Integer
    Dim XLOC                                        As Integer
    Dim XHED                                        As Integer
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "Loan BreakDown sum.xlt")
    Set xlsheet = xlbook.Worksheets(1)
    xlsheet.Cells(2, "A") = "Summary of Employee Loans"
    XLOC = 5
    Set RSDEPT = gconDMIS.Execute("SELECT DISTINCT DEPTCODE FROM HRMS_EMPINFO")
    If Not (RSDEPT.BOF And RSDEPT.EOF) Then
        Do While Not RSDEPT.EOF
            xlsheet.Range("A" & XLOC & ":" & "F" & XLOC).Merge
            xlsheet.Range("A" & XLOC & ":" & "H" & XLOC).BorderAround ColorIndex:=1, WEIGHT:=xlThin, Color:=&H800080
            xlsheet.Range("A" & XLOC & ":" & "H" & XLOC).Interior.Color = &HFFFFC0
            xlsheet.Cells(XLOC, "A") = FindDepartmentName(Null2String(RSDEPT!DEPTCODE))
            
            Set RSTMP = gconDMIS.Execute("SELECT EMPNO, LASTNAME + ', ' + FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE DEPTCODE = " & N2Str2Null(RSDEPT!DEPTCODE) & " ORDER BY FULLNAME")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                Do While Not RSTMP.EOF
                    XHED = 0
                    Set RSLOAN = gconDMIS.Execute("SELECT * FROM HRMS_LOANMAS WHERE EMPNO = '" & Null2String(RSTMP!EMPNO) & "' AND YEAR(DATEGRANTED) = " & cboyear & " ORDER BY ID")
                    If Not (RSLOAN.BOF And RSLOAN.EOF) Then
                        XLOC = XLOC + 1
                        'xlsheet.Range("B" & XLOC & ":" & "H" & XLOC).Merge
                        xlsheet.Range("B" & XLOC & ":" & "H" & XLOC).BorderAround ColorIndex:=1, WEIGHT:=xlThin, Color:=&H800080
                        xlsheet.Range("B" & XLOC & ":" & "H" & XLOC).Interior.Color = &HC0FFC0
                        xlsheet.Cells(XLOC, "B") = Null2String(RSTMP!FULLNAME)
                        Do While Not RSLOAN.EOF
                            XHED = XHED + 1
                            XLOC = XLOC + 1
                            xlsheet.Cells(XLOC, "C") = Null2String(RSLOAN!LOANTYPE)
                            xlsheet.Cells(XLOC, "D") = Null2String(RSLOAN!DATEGRANTED)
                            xlsheet.Cells(XLOC, "E") = Null2String(RSLOAN!DATESTARTED)
                            xlsheet.Cells(XLOC, "F") = Null2String(RSLOAN!MATURITYDATE)
                            xlsheet.Cells(XLOC, "G") = Null2String(RSLOAN!AMOUNTLOANED)
                            xlsheet.Cells(XLOC, "H") = Null2String(RSLOAN!LoanBalance)
                            
                            RSLOAN.MoveNext
                        Loop
                        xlsheet.Range("B" & XLOC - XHED & ":" & "H" & XLOC).BorderAround ColorIndex:=1, WEIGHT:=xlThin, Color:=&H800080
                    End If
                    Set RSLOAN = Nothing
                    
                    RSTMP.MoveNext
                Loop
            End If
            Set RSTMP = Nothing
            XLOC = XLOC + 1
            RSDEPT.MoveNext
        Loop
    End If

    xlApp.Windows.ITEM(1).Caption = "Employee Loans Summary"
    xlApp.Visible = True
    Set xlApp = Nothing
    Set xlsheet = Nothing
    Set xlbook = Nothing
End Sub

Sub PRINTINEXCEL_EMP()

End Sub

Function FindDepartmentName(XXX As String) As String
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DEPTNAME FROM HRMS_DEPARTMENT WHERE DEPTCODE = " & N2Str2Null(XXX) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindDepartmentName = Null2String(RSTMP!DEPTNAME)
    End If
    Set RSTMP = Nothing
End Function

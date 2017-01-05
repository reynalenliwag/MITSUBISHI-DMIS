VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMS_GovermentForms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Philhealth Quartely Remittance Forms(RF1)"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "GovermentForms.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   5595
   Begin VB.Frame Frame1 
      Caption         =   "Applicable Quarter"
      Height          =   1545
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   5355
      Begin VB.OptionButton opt_4 
         Caption         =   "Quarter I on Ending December"
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   1110
         Width           =   3435
      End
      Begin VB.OptionButton opt_3 
         Caption         =   "Quarter III on Ending September"
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   850
         Width           =   3435
      End
      Begin VB.OptionButton opt_2 
         Caption         =   "Quarter II on Ending Jun"
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   590
         Width           =   3435
      End
      Begin VB.OptionButton opt_1 
         Caption         =   "Quarter I on Ending March"
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   3435
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
      Left            =   4680
      MouseIcon       =   "GovermentForms.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "GovermentForms.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   2220
      Width           =   765
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
      Left            =   3930
      MouseIcon       =   "GovermentForms.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "GovermentForms.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   2220
      Width           =   765
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptSSSMonthly 
      Left            =   4200
      Top             =   240
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   330
      TabIndex        =   1
      Top             =   240
      Width           =   435
   End
End
Attribute VB_Name = "frmHRMS_GovermentForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim xlApp                                                         As Excel.Application
    Dim xlsheet                                                       As Excel.Worksheet
    Dim xlbook                                                        As Excel.Workbook
    '------------------------------------------------------------
    Dim RS_HEADER                                                     As ADODB.Recordset
    '------------------------------------------------------------
    Set RS_HEADER = gconDMIS.Execute("SELECT * FROM ALL_PROFILE WHERE MODULENAME='HRMS'")
    If Not (RS_HEADER.EOF Or Not RS_HEADER.BOF) Then
        PHIC_NO = ""
        EMPY_TIN = ""
        EMPY_SSS = ""
        EMPY_TYPE = ""
        EMPY_NAME = ""
        EMPY_ADDRESS = ""
        EMPY_TELEPHONE = ""
    End If
    Dim I                                                             As Integer
    Dim EmployeeShare_1                                               As Double
    Dim EmployerShare_1                                               As Double
    Dim EmployeeShare_2                                               As Double
    Dim EmployerShare_2                                               As Double
    Dim EmployeeShare_3                                               As Double
    Dim EmployerShare_3                                               As Double
    Dim rsPhilHealth                                                  As ADODB.Recordset
    Dim rsPhilHealth2                                                 As ADODB.Recordset
    Dim rsPhilHealth3                                                 As ADODB.Recordset
    Dim rsEmpInfo                                                     As ADODB.Recordset

    Set rsPhilHealth = New ADODB.Recordset
    Set rsPhilHealth2 = New ADODB.Recordset
    Set rsPhilHealth3 = New ADODB.Recordset

    Set rsEmpInfo = New ADODB.Recordset

    Call rsEmpInfo.Open("SELECT   EMPNO,  LASTNAME,FIRSTNAME ,PHNO,LEFT(MIDDLENAME,1) MIDDLENAME  FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE='A' AND PHNO IS NOT NULL ", gconDMIS, adOpenKeyset, adLockReadOnly)
    If (rsEmpInfo.EOF Or rsEmpInfo.BOF) Then
        MsgSpeechBox "No Employee Record"
        Exit Sub
    End If
    Set xlApp = New Excel.Application
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "RF-1.xlt")
    Set xlsheet = xlbook.Worksheets(1)

    xlsheet.Cells(11, "N") = "" & Mid(COMPANY_TIN, 1, 1) & ""
    xlsheet.Cells(11, "O") = "" & Mid(COMPANY_TIN, 2, 1) & ""
    xlsheet.Cells(11, "P") = "" & Mid(COMPANY_TIN, 3, 1) & ""
    xlsheet.Cells(11, "R") = "" & Mid(COMPANY_TIN, 5, 1) & ""
    xlsheet.Cells(11, "S") = "" & Mid(COMPANY_TIN, 6, 1) & ""
    xlsheet.Cells(11, "T") = "" & Mid(COMPANY_TIN, 7, 1) & ""
    xlsheet.Cells(11, "V") = "" & Mid(COMPANY_TIN, 9, 1) & ""
    xlsheet.Cells(11, "W") = "" & Mid(COMPANY_TIN, 10, 1) & ""
    xlsheet.Cells(11, "X") = "" & Mid(COMPANY_TIN, 11, 1) & ""
    xlsheet.Cells(15, "R") = "" & COMPANY_NAME & ""
    xlsheet.Cells(16, "R") = "" & COMPANY_ADDRESS & ""

    I = 0
    While Not rsEmpInfo.EOF
        If opt_1.Value = True Then
            Set rsPhilHealth = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT1, SUM(PHILHEALTHR) AS EMPRAMT1 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='1' AND PAY_YEAR=" & cboyear)
            Set rsPhilHealth2 = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT2, SUM(PHILHEALTHR) AS EMPRAMT2 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='2' AND PAY_YEAR=" & cboyear)
            Set rsPhilHealth3 = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT3, SUM(PHILHEALTHR) AS EMPRAMT3 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='3' AND PAY_YEAR=" & cboyear)
        ElseIf opt_2.Value = True Then
            Set rsPhilHealth = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT1, SUM(PHILHEALTHR) AS EMPRAMT1 FROM   HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH = '4' AND PAY_YEAR='" & cboyear & "'")
            Set rsPhilHealth2 = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT2, SUM(PHILHEALTHR) AS EMPRAMT2 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='5' AND PAY_YEAR='" & cboyear & "'")
            Set rsPhilHealth3 = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT3, SUM(PHILHEALTHR) AS EMPRAMT3 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='6' AND PAY_YEAR='" & cboyear & "'")
        ElseIf opt_3.Value = True Then
            Set rsPhilHealth = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT1, SUM(PHILHEALTHR) AS EMPRAMT1 FROM   HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH = '7' AND PAY_YEAR=" & cboyear)
            Set rsPhilHealth2 = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT2, SUM(PHILHEALTHR) AS EMPRAMT2 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='8' AND PAY_YEAR=" & cboyear)
            Set rsPhilHealth3 = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT3, SUM(PHILHEALTHR) AS EMPRAMT3 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='9' AND PAY_YEAR=" & cboyear)
        ElseIf opt_4.Value = True Then
            Set rsPhilHealth = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT1, SUM(PHILHEALTHR) AS EMPRAMT1 FROM   HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH = '10' AND PAY_YEAR=" & cboyear)
            Set rsPhilHealth2 = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT2, SUM(PHILHEALTHR) AS EMPRAMT2 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='11' AND PAY_YEAR=" & cboyear)
            Set rsPhilHealth3 = gconDMIS.Execute("SELECT SUM(PHILHEALTHE) AS EMPEAMT3, SUM(PHILHEALTHR) AS EMPRAMT3 FROM HRMS_PAYROLL WHERE EMPNO=" & N2Str2Null(rsEmpInfo!EMPNO) & " AND PAY_MONTH ='12' AND PAY_YEAR=" & cboyear)
        End If

        EmployeeShare_1 = 0
        EmployerShare_1 = 0
        EmployeeShare_2 = 0
        EmployerShare_2 = 0
        EmployeeShare_3 = 0
        EmployerShare_3 = 0

        If Not rsPhilHealth.EOF And Not rsPhilHealth.BOF Then
            EmployeeShare_1 = N2Str2Zero(rsPhilHealth!EMPEAMT1)
            EmployerShare_1 = N2Str2Zero(rsPhilHealth!EMPRAMT1)
        End If
        If Not rsPhilHealth2.EOF And Not rsPhilHealth2.BOF Then
            EmployeeShare_2 = N2Str2Zero(rsPhilHealth2!EMPEAMT2)
            EmployerShare_2 = N2Str2Zero(rsPhilHealth2!EMPRAMT2)
        End If
        If Not rsPhilHealth3.EOF And Not rsPhilHealth3.BOF Then
            EmployeeShare_3 = N2Str2Zero(rsPhilHealth3!EMPEAMT3)
            EmployerShare_3 = N2Str2Zero(rsPhilHealth3!EMPRAMT3)
        End If

        I = I + 1
        j = j + 1
        xlsheet.Cells(23 + j, "D") = Null2String(rsEmpInfo!lastname)
        xlsheet.Cells(23 + j, "Q") = Null2String(rsEmpInfo!FIRSTNAME)
        xlsheet.Cells(23 + j, "ad") = Null2String(rsEmpInfo!MIDDLENAME)
        xlsheet.Cells(23 + j, "AF") = Null2String(Replace(rsEmpInfo!PHNO, "-", ""))
        xlsheet.Cells(23 + j, "AR") = EmployeeShare_1
        xlsheet.Cells(23 + j, "AW") = EmployerShare_1
        xlsheet.Cells(23 + j, "BB") = EmployeeShare_2
        xlsheet.Cells(23 + j, "BG") = EmployerShare_2
        xlsheet.Cells(23 + j, "BL") = EmployeeShare_3
        xlsheet.Cells(23 + j, "BQ") = EmployerShare_3

        If I Mod 15 = 0 Then
            xlApp.Visible = True
            Set xlsheet = Nothing
            Set xlbook = Nothing
            Set xlApp = Nothing
            j = 0
            Set xlApp = New Excel.Application
            Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "RF-1.xlt")
            Set xlsheet = xlbook.Worksheets(1)
            xlsheet.Cells(11, "N") = "" & Mid(COMPANY_TIN, 1, 1) & ""
            xlsheet.Cells(11, "O") = "" & Mid(COMPANY_TIN, 2, 1) & ""
            xlsheet.Cells(11, "P") = "" & Mid(COMPANY_TIN, 3, 1) & ""
            xlsheet.Cells(11, "R") = "" & Mid(COMPANY_TIN, 5, 1) & ""
            xlsheet.Cells(11, "S") = "" & Mid(COMPANY_TIN, 6, 1) & ""
            xlsheet.Cells(11, "T") = "" & Mid(COMPANY_TIN, 7, 1) & ""
            xlsheet.Cells(11, "V") = "" & Mid(COMPANY_TIN, 9, 1) & ""
            xlsheet.Cells(11, "W") = "" & Mid(COMPANY_TIN, 10, 1) & ""
            xlsheet.Cells(11, "X") = "" & Mid(COMPANY_TIN, 11, 1) & ""
            xlsheet.Cells(15, "R") = "" & COMPANY_NAME & ""
            xlsheet.Cells(16, "R") = "" & COMPANY_ADDRESS & ""
        End If
        rsEmpInfo.MoveNext
    Wend
    xlApp.Visible = True
    Set xlsheet = Nothing
    Set xlbook = Nothing
    Set xlApp = Nothing
    Set rsPhilHealth = Nothing
    Set rsPhilHealth2 = Nothing
    Set rsPhilHealth3 = Nothing
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'FillcboYear cboyear
    fillcombo_up cboyear
    On Error Resume Next
    cboyear.Text = YEAR(Now)
End Sub


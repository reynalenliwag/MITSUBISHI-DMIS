VERSION 5.00
Begin VB.Form frmHRMS_Reports_R3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS R3 FORM"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_Reports_R3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   6780
   Begin VB.TextBox TXTADDRESS 
      Height          =   405
      Left            =   2370
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   1020
      Width           =   4275
   End
   Begin VB.TextBox TXTTELNO 
      Height          =   405
      Left            =   180
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   1020
      Width           =   2175
   End
   Begin VB.OptionButton OPT_HOUSEHOLD 
      Caption         =   "Household"
      Height          =   255
      Left            =   3180
      TabIndex        =   12
      Top             =   1545
      Width           =   1425
   End
   Begin VB.OptionButton OPT_REGULAR 
      Caption         =   "Regular"
      Height          =   255
      Left            =   2190
      TabIndex        =   11
      Top             =   1545
      Width           =   1095
   End
   Begin VB.TextBox TXTREGEMPLOYERNAME 
      Height          =   405
      Left            =   2370
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   300
      Width           =   4275
   End
   Begin VB.TextBox TXTEMPLOYERIDNUMBER 
      Height          =   405
      Left            =   180
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   300
      Width           =   2145
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2100
      Width           =   1185
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2100
      Width           =   1905
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   795
      Left            =   3120
      MouseIcon       =   "frmHRMS_Reports_R3.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Reports_R3.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Window"
      Top             =   2640
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   795
      Left            =   2430
      MouseIcon       =   "frmHRMS_Reports_R3.frx":0A42
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Reports_R3.frx":0B94
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print this Record"
      Top             =   2640
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2370
      TabIndex        =   16
      Top             =   780
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TEL NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   180
      TabIndex        =   14
      Top             =   780
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TYPE OF EMPLOYEE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   210
      TabIndex        =   10
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "REGISTERED EMPLOYER NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2370
      TabIndex        =   9
      Top             =   60
      Width           =   2595
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "EMPLOYER ID NUMBER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   8
      Top             =   60
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   2250
      TabIndex        =   4
      Top             =   1830
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quarter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   3
      Top             =   1830
      Width           =   720
   End
End
Attribute VB_Name = "frmHRMS_Reports_R3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Call SaveSetting("ADMS 1.0", "HRMS", "SSS_EMPLOYERIDNUMBER", TXTEMPLOYERIDNUMBER)
    Call SaveSetting("ADMS 1.0", "HRMS", "SSS_EMPLOYERNAME", TXTREGEMPLOYERNAME)
    Call SaveSetting("ADMS 1.0", "HRMS", "SSS_TELNO", TXTTELNO)
    Call SaveSetting("ADMS 1.0", "HRMS", "SSS_ADDRESS", TXTADDRESS)

    If OPT_REGULAR.Value = True Then
        Call SaveSetting("ADMS 1.0", "HRMS", "SSS_TYPEOFEMPLOYEE", "R")
    Else
        Call SaveSetting("ADMS 1.0", "HRMS", "SSS_TYPEOFEMPLOYEE", "H")
    End If


    Dim xlApp                                        As Excel.Application
    Dim xlsheet                                      As Excel.Worksheet
    Dim xlbook                                       As Excel.Workbook



    Set xlApp = New Excel.Application
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "R3.xlt")
    Set xlsheet = xlbook.Worksheets(1)


    Dim Y                                            As Integer
    Y = 0


    xlsheet.Cells(10, "C") = TXTEMPLOYERIDNUMBER
    xlsheet.Cells(10, "M") = TXTREGEMPLOYERNAME
    xlsheet.Cells(12, "M") = TXTADDRESS
    xlsheet.Cells(12, "C") = TXTTELNO

    If OPT_REGULAR.Value = True Then
        xlsheet.Cells(12, "P") = "REGULAR"
    Else
        xlsheet.Cells(12, "P") = "HOUSE HOLD"
    End If




    Dim RS_SSS                                       As ADODB.Recordset

    If Combo1 = "QUARTER I" Then
        xlsheet.Cells(10, "P") = "3 " & Combo2
        Set RS_SSS = gconDMIS.Execute("SELECT  SSSNO,EMPNO ,LASTNAME + ', ' + FIRSTNAME +'.' + LEFT(MIDDLENAME,1) AS EMPLOYEENAME ," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=1 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M1," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=2 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M2," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=3 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M3" & _
                                    " From HRMS_EMPINFO")
    ElseIf Combo1 = "QUARTER II" Then
        xlsheet.Cells(10, "P") = "6 " & Combo2
        Set RS_SSS = gconDMIS.Execute("SELECT SSSNO, EMPNO ,LASTNAME + ', ' + FIRSTNAME +'.' + LEFT(MIDDLENAME,1) AS EMPLOYEENAME ," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=4 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M1," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=5 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M2," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=6 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M3" & _
                                    " From HRMS_EMPINFO")
    ElseIf Combo1 = "QUARTER III" Then
        xlsheet.Cells(10, "P") = "9 " & Combo2
        Set RS_SSS = gconDMIS.Execute("SELECT  SSSNO,EMPNO ,LASTNAME + ', ' + FIRSTNAME +'.' + LEFT(MIDDLENAME,1) AS EMPLOYEENAME ," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=7 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M1," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=8 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M2," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=9 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M3" & _
                                    " From HRMS_EMPINFO")
    Else
        xlsheet.Cells(10, "P") = "12 " & Combo2
        Set RS_SSS = gconDMIS.Execute("SELECT SSSNO, EMPNO ,LASTNAME + ', ' + FIRSTNAME +'.' + LEFT(MIDDLENAME,1) AS EMPLOYEENAME ," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=10 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M1," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=11 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M2," & _
                                      "(SELECT SUM(SSSE) FROM HRMS_PAYROLL WHERE PAY_MONTH=12 AND PAY_YEAR=" & Combo2 & " AND EMPNO=HRMS_EMPINFO.EMPNO ) AS M3" & _
                                    " From HRMS_EMPINFO")
    End If



    Dim SSSNO                                        As String




    If Not RS_SSS.EOF And Not RS_SSS.BOF Then
        RS_SSS.MoveFirst
        Y = 15
        While Not RS_SSS.EOF
            CNTR = CNTR + 1
            If CNTR Mod 15 = 0 Then
                xlApp.Visible = True
                Set xlApp = New Excel.Application
                Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "R3.xlt")
                Set xlsheet = xlbook.Worksheets(1)

                xlsheet.Cells(10, "C") = TXTEMPLOYERIDNUMBER
                xlsheet.Cells(10, "M") = TXTREGEMPLOYERNAME
                xlsheet.Cells(12, "M") = TXTADDRESS
                xlsheet.Cells(12, "C") = TXTTELNO

                If OPT_REGULAR.Value = True Then
                    xlsheet.Cells(12, "P") = "REGULAR"
                Else
                    xlsheet.Cells(12, "P") = "HOUSE HOLD"
                End If


                If Combo1 = "QUARTER I" Then
                    xlsheet.Cells(10, "P") = "3 " & Combo2
                ElseIf Combo1 = "QUARTER II" Then
                    xlsheet.Cells(10, "P") = "6 " & Combo2
                ElseIf Combo1 = "QUARTER III" Then
                    xlsheet.Cells(10, "P") = "9 " & Combo2
                Else
                    xlsheet.Cells(10, "P") = "12 " & Combo2
                End If
                Y = 15
            End If

            SSSNO = Replace(Null2String(RS_SSS!SSSNO), "-", "")
            If Len(SSSNO) > 0 Then
                xlsheet.Cells(Y, "C") = Mid(SSSNO, 1, 1)
                xlsheet.Cells(Y, "D") = Mid(SSSNO, 2, 1)
                xlsheet.Cells(Y, "E") = Mid(SSSNO, 3, 1)
                xlsheet.Cells(Y, "F") = Mid(SSSNO, 4, 1)
                xlsheet.Cells(Y, "G") = Mid(SSSNO, 5, 1)
                xlsheet.Cells(Y, "H") = Mid(SSSNO, 6, 1)
                xlsheet.Cells(Y, "I") = Mid(SSSNO, 7, 1)
                xlsheet.Cells(Y, "J") = Mid(SSSNO, 8, 1)
                xlsheet.Cells(Y, "K") = Mid(SSSNO, 9, 1)
                xlsheet.Cells(Y, "L") = Mid(SSSNO, 10, 1)
            End If
            xlsheet.Cells(Y, "M") = RS_SSS!EMPLOYEENAME
            xlsheet.Cells(Y, "P") = RS_SSS!M1
            xlsheet.Cells(Y, "Q") = RS_SSS!M2
            xlsheet.Cells(Y, "R") = RS_SSS!M3

            Y = Y + 1
            RS_SSS.MoveNext
        Wend
    End If



    Set RSTOTAL = Nothing
    Set rsTemp = Nothing

    xlApp.Visible = True
    Set xlsheet = Nothing
    Set xlbook = Nothing
    Set xlApp = Nothing

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Combo1.AddItem "QUARTER I"
    Combo1.AddItem "QUARTER II"
    Combo1.AddItem "QUARTER III"
    Combo1.AddItem "QUARTER IV"
    Combo1.ListIndex = 0
    TXTREGEMPLOYERNAME = GetSetting("ADMS 1.0", "HRMS", "SSS_EMPLOYERNAME", COMPANY_NAME)
    TXTEMPLOYERIDNUMBER = GetSetting("ADMS 1.0", "HRMS", "SSS_EMPLOYERIDNUMBER", "")
    TXTTELNO = GetSetting("ADMS 1.0", "HRMS", "SSS_TELNO", "")
    TXTADDRESS = GetSetting("ADMS 1.0", "HRMS", "SSS_ADDRESS", COMPANY_ADDRESS)

    If GetSetting("ADMS 1.0", "HRMS", "SSS_TYPEOFEMPLOYEE", "") = "R" Then
        OPT_REGULAR.Value = 1
        OPT_HOUSEHOLD.Value = 0
    ElseIf GetSetting("ADMS 1.0", "HRMS", "SSS_TYPEOFEMPLOYEE", "") = "H" Then
        OPT_REGULAR.Value = 0
        OPT_HOUSEHOLD.Value = 1
    Else
        OPT_REGULAR.Value = 0
        OPT_HOUSEHOLD.Value = 0
    End If
    'FillcboYear Combo2
    fillcombo_up Combo2
    Combo2.Text = YEAR(LOGDATE)
End Sub

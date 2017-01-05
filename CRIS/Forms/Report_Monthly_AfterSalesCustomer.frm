VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmCRIS_Report_AfterSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "After Sales Report"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_Monthly_AfterSalesCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3675
   Begin VB.OptionButton Option1 
      Caption         =   "SALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   870
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1860
      Width           =   2355
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   990
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "9999"
      Top             =   2340
      Width           =   2325
   End
   Begin VB.ComboBox Combo1 
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
      ForeColor       =   &H00400000&
      Height          =   360
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1380
      Width           =   2355
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   825
      Left            =   2145
      MouseIcon       =   "Report_Monthly_AfterSalesCustomer.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Report_Monthly_AfterSalesCustomer.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   2850
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   825
      Left            =   1275
      MouseIcon       =   "Report_Monthly_AfterSalesCustomer.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "Report_Monthly_AfterSalesCustomer.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   2850
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   4440
      Top             =   270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "MMPC Monthly Inventory Control"
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
   Begin VB.OptionButton Option2 
      Caption         =   "SERVICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   870
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   120
      Picture         =   "Report_Monthly_AfterSalesCustomer.frx":19D0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   405
      TabIndex        =   8
      Top             =   2400
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   225
      TabIndex        =   7
      Top             =   1860
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1380
      Picture         =   "Report_Monthly_AfterSalesCustomer.frx":24D5
      Top             =   2880
      Width           =   1500
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   825
      Left            =   0
      TabIndex        =   4
      Top             =   -30
      Width           =   3795
      _Version        =   655364
      _ExtentX        =   6694
      _ExtentY        =   1455
      _StockProps     =   14
      Caption         =   "After-Sales Report    "
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   2
      ForeColor       =   4194304
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SAE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   450
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "frmCRIS_Report_AfterSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportType                          As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    If IsNumeric(txtYear) = False Then: MsgSpeech (" Error In Date"): txtYear.SetFocus: Exit Sub
    Screen.MousePointer = 11
    frmSplash.Show
    frmSplash.labCon = "Extracting Data to Excel... Please Wait"
    If ReportType = "SALES" Then
        'PRINTSERVICE
    'Else
        PRINTSALES
    End If
    Unload frmSplash
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))

    txtYear.Text = Year(LOGDATE)

'    If ReportType = "SERVICE" Then
'        Me.Caption = "After Sales Report:Customer Directory-Service"
'        Combo_Loadval Combo1, gconDMIS.Execute("SELECT distinct upper(WRITER)  from CSMS_REPAIRORDER ")
'        Combo1.AddItem "ALL", 0
'        Label3.Caption = "SA"
'        Combo1.ListIndex = 0
'    ElseIf ReportType = "SALES" Then
        ReportType = "SALES"
        Label3.Caption = "SAE"
        Me.Caption = "After Sales Report:Customer Directory-Sales"
        Combo_Loadval Combo1, gconDMIS.Execute("SELECT distinct upper(SALESAE) from SMIS_SALESORDER")
        Combo1.AddItem "ALL", 0
        Combo1.ListIndex = 0
'    End If

    Screen.MousePointer = 0
End Sub
Sub PRINTSERVICE()
    Dim SQLTXT As String
    Dim RSTMP As New ADODB.Recordset
    
    Dim xlApp
    Dim xlbook
    Dim xlSheet1
    Dim xlSheet2
    Set xlApp = CreateObject("Excel.Application")
    Set xlbook = xlApp.Workbooks.Open(App.Path & "\CRIS_EXCEL\AfterSalesReportsSERVICE.xlt")
    Set xlSheet1 = xlbook.Worksheets(1)
    Set xlSheet2 = xlbook.Worksheets(2)

    xlSheet1.Cells(3, 1) = "SERVICE : " & cboMonth & " " & txtYear
    xlSheet2.Cells(3, 1) = "SERVICE : " & cboMonth & " " & txtYear
    If Combo1.Text <> "ALL" Then
        xlSheet1.Cells(3, 4) = "SERVICE ADVISOR: " & Combo1
        xlSheet2.Cells(3, 4) = "SERVICE ADVISOR: " & Combo1
    End If

    SQLTXT = "SELECT ROW_NUMBER() OVER (ORDER BY DTE_FINISHED) AS [NO],FIRSTNAME,LASTNAME," & vbCrLf
    SQLTXT = SQLTXT & "ADDRESS1 , ADDRESS2, HOMEPHONE, TELEPHONENO, DTE_FINISHED, Model, VIN, DTE_RECD" & vbCrLf
    SQLTXT = SQLTXT & "From" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT  B.CUSCDE,B.FIRSTNAME,B.LASTNAME,B.ADDRESS1,B.ADDRESS2,B.HOMEPHONE,B.TELEPHONENO," & vbCrLf
    SQLTXT = SQLTXT & "B.DTE_FINISHED,A.MODEL,A.VIN,A.PLATE_NO,RECD_BY,JOBTYPE,WRITER,CUSTYPE," & vbCrLf
    SQLTXT = SQLTXT & "Case JOBTYPE" & vbCrLf
    SQLTXT = SQLTXT & "WHEN 'PMS' THEN ISNULL(DTE_RECD,'')" & vbCrLf
    SQLTXT = SQLTXT & "END As DTE_RECD" & vbCrLf
    SQLTXT = SQLTXT & "FROM CSMS_CUSVEH A INNER JOIN" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT X.CUSCDE,ISNULL(FIRSTNAME,'') AS FIRSTNAME,ISNULL(X.LASTNAME,'') AS LASTNAME,Y.PLATE_NO," & vbCrLf
    SQLTXT = SQLTXT & "ISNULL(X.CUSTOMERADD,'') AS ADDRESS1,ISNULL(X.PROVINCIALADD,'') AS ADDRESS2,X.CUSTYPE," & vbCrLf
    SQLTXT = SQLTXT & "ISNULL(HOMEPHONE,'') AS HOMEPHONE,ISNULL(TELEPHONENO,'') AS TELEPHONENO, DTE_COMP AS DTE_FINISHED,JOBTYPE,DTE_RECD,RECD_BY,WRITER" & vbCrLf
    SQLTXT = SQLTXT & "FROM ALL_CUSTOMER_TABLE X INNER JOIN" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT WRITER,A.REP_OR,PLATE_NO,ACCT_NO,MAX(ISNULL(DTE_COMP,'')) AS DTE_COMP,MAX(ISNULL(DTE_RECD,'')) AS DTE_RECD,JOBTYPE,RECD_BY  FROM" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT REP_OR,A.PLATE_NO,A.ACCT_NO,A.DTE_COMP,A.DTE_RECD,A.RECD_BY ,B.WRITER FROM CSMS_REPOR A INNER JOIN CSMS_REPAIRORDER B" & vbCrLf
    SQLTXT = SQLTXT & "ON A.REP_OR = B.RO_NO AND A.PLATE_NO = B.PLATE_NO" & vbCrLf
    SQLTXT = SQLTXT & "WHERE A.TRANSTYPE ='R') A LEFT OUTER JOIN" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT REP_OR,JOBTYPE,TECHCODE FROM CSMS_RO_DET WHERE LIVIL = '1' AND JOBTYPE = 'PMS'" & vbCrLf
    SQLTXT = SQLTXT & ") B ON A.REP_OR = B.REP_OR WHERE MONTH(DTE_COMP)= " & What_month(cboMonth) & " AND YEAR(DTE_COMP)=" & txtYear & "" & vbCrLf
    SQLTXT = SQLTXT & "GROUP BY PLATE_NO,ACCT_NO,JOBTYPE,RECD_BY,A.REP_OR,WRITER" & vbCrLf
    SQLTXT = SQLTXT & ")Y ON X.CUSCDE = Y.ACCT_NO" & vbCrLf
    SQLTXT = SQLTXT & ")B ON A.PLATE_NO = B.PLATE_NO" & vbCrLf
    SQLTXT = SQLTXT & ")T WHERE " & vbCrLf
     If Combo1.Text <> "ALL" Then
        SQLTXT = SQLTXT & " WRITER =" & N2Str2Null(Combo1) & " AND " & vbCrLf
    End If
    SQLTXT = SQLTXT & " CUSTYPE = 'P'  ORDER BY DTE_FINISHED "
    
    Set RSTMP = gconDMIS.Execute(SQLTXT)
    
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        xlSheet1.Cells(7, 1).CopyFromRecordset RSTMP
    Else
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If
    
    Set RSTMP = Nothing
    SQLTXT = ""
    
    SQLTXT = "SELECT ROW_NUMBER() OVER (ORDER BY DTE_FINISHED) AS [NO],FIRSTNAME,LASTNAME," & vbCrLf
    SQLTXT = SQLTXT & "LASTNAME , ADDRESS1,ADDRESS2, HOMEPHONE, TELEPHONENO, DTE_FINISHED, Model, VIN, DTE_RECD" & vbCrLf
    SQLTXT = SQLTXT & "From" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT  B.CUSCDE,B.FIRSTNAME,B.LASTNAME,B.ADDRESS1,B.ADDRESS2,B.HOMEPHONE,B.TELEPHONENO," & vbCrLf
    SQLTXT = SQLTXT & "B.DTE_FINISHED,A.MODEL,A.VIN,A.PLATE_NO,RECD_BY,JOBTYPE,WRITER,CUSTYPE," & vbCrLf
    SQLTXT = SQLTXT & "Case JOBTYPE" & vbCrLf
    SQLTXT = SQLTXT & "WHEN 'PMS' THEN ISNULL(DTE_RECD,'')" & vbCrLf
    SQLTXT = SQLTXT & "END As DTE_RECD" & vbCrLf
    SQLTXT = SQLTXT & "FROM CSMS_CUSVEH A INNER JOIN" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT X.CUSCDE,ISNULL(FIRSTNAME,'') AS FIRSTNAME,ISNULL(X.LASTNAME,'') AS LASTNAME,Y.PLATE_NO," & vbCrLf
    SQLTXT = SQLTXT & "ISNULL(X.CUSTOMERADD,'') AS ADDRESS1,ISNULL(X.PROVINCIALADD,'') AS ADDRESS2,X.CUSTYPE," & vbCrLf
    SQLTXT = SQLTXT & "ISNULL(HOMEPHONE,'') AS HOMEPHONE,ISNULL(TELEPHONENO,'') AS TELEPHONENO, DTE_COMP AS DTE_FINISHED,JOBTYPE,DTE_RECD,RECD_BY,WRITER" & vbCrLf
    SQLTXT = SQLTXT & "FROM ALL_CUSTOMER_TABLE X INNER JOIN" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT WRITER,A.REP_OR,PLATE_NO,ACCT_NO,MAX(ISNULL(DTE_COMP,'')) AS DTE_COMP,MAX(ISNULL(DTE_RECD,'')) AS DTE_RECD,JOBTYPE,RECD_BY  FROM" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT REP_OR,A.PLATE_NO,A.ACCT_NO,A.DTE_COMP,A.DTE_RECD,A.RECD_BY ,B.WRITER FROM CSMS_REPOR A INNER JOIN CSMS_REPAIRORDER B" & vbCrLf
    SQLTXT = SQLTXT & "ON A.REP_OR = B.RO_NO AND A.PLATE_NO = B.PLATE_NO" & vbCrLf
    SQLTXT = SQLTXT & "WHERE A.TRANSTYPE ='R') A LEFT OUTER JOIN" & vbCrLf
    SQLTXT = SQLTXT & "(" & vbCrLf
    SQLTXT = SQLTXT & "SELECT REP_OR,JOBTYPE,TECHCODE FROM CSMS_RO_DET WHERE LIVIL = '1' AND JOBTYPE = 'PMS'" & vbCrLf
    SQLTXT = SQLTXT & ") B ON A.REP_OR = B.REP_OR WHERE MONTH(DTE_COMP)= " & What_month(cboMonth) & " AND YEAR(DTE_COMP)=" & txtYear & "" & vbCrLf
    SQLTXT = SQLTXT & "GROUP BY PLATE_NO,ACCT_NO,JOBTYPE,RECD_BY,A.REP_OR,WRITER" & vbCrLf
    SQLTXT = SQLTXT & ")Y ON X.CUSCDE = Y.ACCT_NO" & vbCrLf
    SQLTXT = SQLTXT & ")B ON A.PLATE_NO = B.PLATE_NO" & vbCrLf
    SQLTXT = SQLTXT & ")T WHERE " & vbCrLf
     If Combo1.Text <> "ALL" Then
        SQLTXT = SQLTXT & " WRITER =" & N2Str2Null(Combo1) & " AND " & vbCrLf
    End If
    SQLTXT = SQLTXT & " CUSTYPE IN ('F','P C','C P','G','C', ' ') ORDER BY DTE_FINISHED "
    
    Set RSTMP = gconDMIS.Execute(SQLTXT)
    
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        xlSheet2.Cells(7, 1).CopyFromRecordset RSTMP
    Else
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If
    
    
    xlApp.Visible = True
    Set xlbook = Nothing
    Set xlSheet1 = Nothing
    Set xlSheet2 = Nothing
    Set xlApp = Nothing
    Set RSTMP = Nothing
End Sub



Sub PRINTSERVICE1()
    Dim SQL                             As String
    Dim rsCust                          As ADODB.Recordset

    On Error GoTo ErrorCode                                  'AXP063119:34

    SQL = " SELECT"
    SQL = SQL & " C.NIYM, C.CUSCDE, C.VIN, C.MAKE, C.MODEL, CUST.LASTNAME, "
    SQL = SQL & " CUST.FIRSTNAME , CUST.PROVINCIALADD, CUST.CUSTOMERADD, CUST.MOBILE ,CUST.CITY, CUST.HOMEPHONE, CUST.TELEPHONENO, CUST.CUSTYPE ,"
    SQL = SQL & " REPAIR.DTE_COMP AS DATEFINISH, RO.WRITER "
    SQL = SQL & " FROM  CSMS_CUSVEH C "
    SQL = SQL & " Inner Join "
    SQL = SQL & " CSMS_REPOR REPAIR ON C.PLATE_NO = REPAIR.PLATE_NO "
    SQL = SQL & " Inner Join "
    SQL = SQL & " CSMS_REPAIRORDER RO ON RO.RO_NO= REPAIR.REP_OR  "
    SQL = SQL & " Inner Join "
    SQL = SQL & " ALL_CUSTOMER_TABLE CUST ON C.CUSCDE = CUST.CUSCDE WHERE MONTH(DTE_COMP)= " & What_month(cboMonth) & " AND YEAR(DTE_COMP)=" & txtYear
    If Combo1.Text <> "ALL" Then
        SQL = SQL & " AND WRITER =" & N2Str2Null(Combo1)
    End If
    SQL = SQL & " ORDER BY REPAIR.DATEFINISH "
    Set rsCust = gconDMIS.Execute(SQL)

    If rsCust.EOF Or rsCust.BOF Then
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If
    If rsCust.EOF Or rsCust.BOF Then
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If
    Dim xlApp
    Dim xlbook
    Dim xlSheet1
    Dim xlSheet2
    Set xlApp = CreateObject("Excel.Application")
    Set xlbook = xlApp.Workbooks.Open(App.Path & "\CRIS_EXCEL\AfterSalesReportsSERVICE.xlt")
    Set xlSheet1 = xlbook.Worksheets(1)
    Set xlSheet2 = xlbook.Worksheets(2)

    Dim i                               As Integer
    Dim j                               As Integer


    xlSheet1.Cells(3, 1) = "SERVICE : " & cboMonth & " " & txtYear
    xlSheet2.Cells(3, 1) = "SERVICE : " & cboMonth & " " & txtYear
    If Combo1.Text <> "ALL" Then
        xlSheet1.Cells(3, 4) = "SERVICE ADVISOR: " & Combo1
        xlSheet2.Cells(3, 4) = "SERVICE ADVISOR: " & Combo1

    End If

    If Not rsCust.EOF And Not rsCust.BOF Then
        Do While Not rsCust.EOF
            If Null2String(rsCust!CUSTYPE) = "P" Or Null2String(rsCust!CUSTYPE) = "" Then
                xlSheet1.Cells(7 + j, 1) = j + 1
                xlSheet1.Cells(7 + j, 2) = UCase(Null2String(rsCust!FIRSTNAME))
                xlSheet1.Cells(7 + j, 3) = UCase(Null2String(rsCust!lastname))
                xlSheet1.Cells(7 + j, 4) = Null2String(rsCust!CUSTOMERADD)
                xlSheet1.Cells(7 + j, 5) = Null2String(rsCust!provincialadd) & " " & Null2String(rsCust!CITY)
                xlSheet1.Cells(7 + j, 6) = Null2String(rsCust!HomePhone)
                xlSheet1.Cells(7 + j, 7) = Null2String(rsCust!Mobile)
                xlSheet1.Cells(7 + j, 8) = Null2String(rsCust!datefinish)
                xlSheet1.Cells(7 + j, 9) = IIf(Null2String(rsCust!Make) = "", Null2String(rsCust!Model), Null2String(rsCust!Make) & " " & Null2String(rsCust!Model))
                xlSheet1.Cells(7 + j, 10) = Null2String(rsCust!Vin)
                xlSheet1.Cells(7 + j, 11) = GetLastPMSDate(rsCust.Fields(1))
                j = j + 1
            Else
                xlSheet2.Cells(7 + i, 1) = i + 1
                xlSheet2.Cells(7 + i, 2) = UCase(Null2String(rsCust!FIRSTNAME))
                xlSheet2.Cells(7 + i, 3) = UCase(Null2String(rsCust!lastname))
                xlSheet2.Cells(7 + i, 4) = UCase(Null2String(rsCust!lastname))
                xlSheet2.Cells(7 + i, 5) = Null2String(rsCust!CUSTOMERADD)
                xlSheet2.Cells(7 + i, 6) = Null2String(rsCust!provincialadd) & " " & Null2String(rsCust!CITY)
                xlSheet2.Cells(7 + i, 7) = Null2String(rsCust!HomePhone)
                xlSheet2.Cells(7 + i, 8) = Null2String(rsCust!Mobile)
                xlSheet2.Cells(7 + i, 9) = Null2String(rsCust!datefinish)
                xlSheet1.Cells(7 + j, 9) = IIf(Null2String(rsCust!Make) = "", Null2String(rsCust!Model), Null2String(rsCust!Make) & " " & Null2String(rsCust!Model))
                xlSheet2.Cells(7 + i, 11) = Null2String(rsCust!Vin)
                i = i + 1
            End If
            rsCust.MoveNext
        Loop
    End If
    xlApp.Visible = True
    Set xlbook = Nothing
    Set xlSheet1 = Nothing
    Set xlSheet2 = Nothing
    Set xlApp = Nothing



    Exit Sub
ErrorCode:
    ShowVBError

End Sub
'---------------------------------------------------------------------------------------
' Procedure : PRINTSALES
' DateTime  : 10/24/2007 15:35
' Author    : Ashish
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function GetLastPMSDate(Xcode As String) As String
    Dim RSTMP                           As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT ISNULL(MAX(A.DTE_RECD),''), A.ACCT_NO FROM CSMS_REPOR A INNER JOIN CSMS_RO_DET B ON A.REP_OR = B.REP_OR " & _
        " WHERE B.JOBTYPE = 'PMS' AND B.LIVIL = 1 " & _
        " AND A.ACCT_NO = " & N2Str2Null(Xcode) & _
        " GROUP BY A.ACCT_NO")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetLastPMSDate = Null2String(RSTMP.Fields(0))
    End If
    Set RSTMP = Nothing
End Function

Sub PRINTSALES()

    Dim rsCust                          As ADODB.Recordset
    Dim SQL                             As String
    On Error GoTo ErrorCode                                  'AXP063115:35

    Set rsCust = New ADODB.Recordset
    SQL = "SELECT FIRSTNAME, LASTNAME,CUSTOMERADD, PROVINCIALADD, CITY ,HOMEPHONE, MOBILE, "
    SQL = SQL & " MODELDESCRIPTION,VINO, DateReleased ,CUSTYPE FROM SMIS_SALESORDER "
    SQL = SQL & " INNER JOIN ALL_CUSTOMER ON ALL_CUSTOMER.CUSCDE=SMIS_SALESORDER.CODE WHERE MONTH(DateReleased)= " & What_month(cboMonth) & " AND YEAR(DateReleased )=" & txtYear

    If Combo1.Text <> "ALL" Then
        SQL = SQL & " AND SALESAE=" & N2Str2Null(Combo1)
    End If
    SQL = SQL & " ORDER BY DateReleased "
    Set rsCust = gconDMIS.Execute(SQL)

    If rsCust.EOF Or rsCust.BOF Then
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If

    Dim xlApp
    Dim xlbook
    Dim xlSheet1
    Dim xlSheet2
    Set xlApp = CreateObject("Excel.Application")

    Set xlbook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "\SMIS_EXCEL\AfterSalesReportsSales.xlt")
    Set xlSheet1 = xlbook.Worksheets(1)
    Set xlSheet2 = xlbook.Worksheets(2)

    Dim i                               As Integer
    Dim j                               As Integer
    xlSheet1.Cells(2, 1) = "DEALER NAME : " & COMPANY_NAME
    xlSheet2.Cells(2, 1) = "DEALER NAME : " & COMPANY_NAME
    xlSheet1.Cells(3, 1) = "SALES : " & cboMonth & " " & txtYear
    xlSheet2.Cells(3, 1) = "SALES : " & cboMonth & " " & txtYear

    If Combo1.Text <> "ALL" Then
        xlSheet1.Cells(3, 4) = "SALES AGENT: " & Combo1
        xlSheet2.Cells(3, 4) = "SALES AGENT: " & Combo1

    End If

    If Not rsCust.EOF And Not rsCust.BOF Then
        Do While Not rsCust.EOF
            If Null2String(rsCust!CUSTYPE) = "P" Or Null2String(rsCust!CUSTYPE) = "" Then
                xlSheet1.Cells(7 + j, 1) = j + 1
                xlSheet1.Cells(7 + j, 2) = Null2String(rsCust!FIRSTNAME)
                xlSheet1.Cells(7 + j, 3) = Null2String(rsCust!lastname)
                xlSheet1.Cells(7 + j, 4) = Null2String(rsCust!CUSTOMERADD)
                xlSheet1.Cells(7 + j, 5) = Null2String(rsCust!provincialadd) & " " & Null2String(rsCust!CITY)
                xlSheet1.Cells(7 + j, 6) = Null2String(rsCust!HomePhone)
                xlSheet1.Cells(7 + j, 7) = Null2String(rsCust!Mobile)
                xlSheet1.Cells(7 + j, 8) = Null2String(rsCust!DateReleased)
                xlSheet1.Cells(7 + j, 9) = Null2String(rsCust!modeldescription)
                xlSheet1.Cells(7 + j, 10) = Null2String(rsCust!Vino)
                j = j + 1
            Else
                xlSheet2.Cells(7 + i, 1) = i + 1
                xlSheet2.Cells(7 + i, 2) = Null2String(rsCust!FIRSTNAME)
                xlSheet2.Cells(7 + i, 3) = Null2String(rsCust!lastname)
                xlSheet2.Cells(7 + i, 4) = Null2String(rsCust!lastname)
                xlSheet2.Cells(7 + i, 5) = Null2String(rsCust!CUSTOMERADD)
                xlSheet2.Cells(7 + i, 6) = Null2String(rsCust!provincialadd) & " " & Null2String(rsCust!CITY)
                xlSheet2.Cells(7 + i, 7) = Null2String(rsCust!HomePhone)
                xlSheet2.Cells(7 + i, 8) = Null2String(rsCust!Mobile)
                xlSheet2.Cells(7 + i, 9) = Null2String(rsCust!DateReleased)
                xlSheet2.Cells(7 + i, 10) = Null2String(rsCust!modeldescription)
                xlSheet2.Cells(7 + i, 11) = Null2String(rsCust!Vino)
                i = i + 1
            End If
            rsCust.MoveNext
        Loop
    End If
    xlApp.Visible = True
    Set xlbook = Nothing
    Set xlSheet1 = Nothing
    Set xlSheet2 = Nothing
    Set xlApp = Nothing


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub ServiceReport()
    ReportType = "SERVICE"
End Sub
Sub SalesReport()
    ReportType = "SALES"
End Sub

Private Sub Option1_Click()
    Label3.Caption = "SAE"
    Me.Caption = "After Sales Report:Customer Directory-Sales"
    Combo_Loadval Combo1, gconDMIS.Execute("SELECT distinct UPPER(SALESAE) from SMIS_SALESORDER")
    Combo1.AddItem "ALL", 0

    ReportType = "SALES"
    Combo1.ListIndex = 0
    Combo1.SetFocus
End Sub

Private Sub Option2_Click()
    Me.Caption = "After Sales Report:Customer Directory-Service"
    Combo_Loadval Combo1, gconDMIS.Execute("SELECT distinct UPPER(WRITER)  from CSMS_REPAIRORDER ")
    Combo1.AddItem "ALL", 0

    Label3.Caption = "SA"
    ReportType = "SERVICE"
    Combo1.ListIndex = 0
    Combo1.SetFocus
End Sub

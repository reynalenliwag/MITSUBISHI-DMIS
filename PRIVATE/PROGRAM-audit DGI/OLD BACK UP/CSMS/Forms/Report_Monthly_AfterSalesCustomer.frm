VERSION 5.00
Begin VB.Form frmCSMS_Report_AfterSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "After Sales Report"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   3975
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   60
      TabIndex        =   2
      Top             =   810
      Width           =   3885
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   825
         Left            =   2460
         MouseIcon       =   "Report_Monthly_AfterSalesCustomer.frx":0E42
         MousePointer    =   99  'Custom
         Picture         =   "Report_Monthly_AfterSalesCustomer.frx":0F94
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Close Window"
         Top             =   2310
         Width           =   885
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
         TabIndex        =   5
         Top             =   780
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
         TabIndex        =   4
         Text            =   "9999"
         Top             =   1740
         Width           =   2325
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
         TabIndex        =   3
         Top             =   1260
         Width           =   2355
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   825
         Left            =   1620
         MouseIcon       =   "Report_Monthly_AfterSalesCustomer.frx":13DF
         MousePointer    =   99  'Custom
         Picture         =   "Report_Monthly_AfterSalesCustomer.frx":1531
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Print Report"
         Top             =   2310
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Service Sales Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   3735
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
         Left            =   390
         TabIndex        =   10
         Top             =   840
         Width           =   495
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   720
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
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   510
      End
   End
   Begin VB.PictureBox rptReleased 
      Height          =   480
      Left            =   4440
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   270
      Width           =   1200
   End
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
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   690
      Left            =   120
      Picture         =   "Report_Monthly_AfterSalesCustomer.frx":19D0
      Top             =   30
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1380
      Picture         =   "Report_Monthly_AfterSalesCustomer.frx":2272
      Top             =   2880
      Width           =   1500
   End
End
Attribute VB_Name = "frmCSMS_Report_AfterSales"
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

    If Len(Dir(App.Path & "\AfterSalesReportsSERVICE.xlt")) <= 0 Then
        If EXTRACT_FILES(111, "\AfterSalesReportsSERVICE.xlt") = False Then
            MsgBox "Please Put AfterSalesReportsSERVICE.XLT on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If

    If IsNumeric(txtYear) = False Then: MsgSpeech (" Error In Date"): txtYear.SetFocus: Exit Sub
    Screen.MousePointer = 11
    frmSplash.Show
    frmSplash.labCon = "Extracting Data to Excel... Please Wait"
    If ReportType = "SERVICE" Then
        PRINTSERVICE
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
    ReportType = "SERVICE"
    If ReportType = "SERVICE" Then
        Me.Caption = "After Sales Report:Customer Directory-Service"
        Combo_Loadval Combo1, gconDMIS.Execute("SELECT distinct upper(WRITER)  from CSMS_REPAIRORDER ")
        Combo1.AddItem "ALL", 0
        Label3.Caption = "SA"
        Combo1.ListIndex = 0
    End If

    Screen.MousePointer = 0
End Sub
Sub PRINTSERVICE()
    Dim SQLTXT As String
    Dim rstmp As New ADODB.Recordset
    
    Dim xlApp
    Dim xlBook
    Dim xlSheet1
    Dim xlSheet2
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\AfterSalesReportsSERVICE.xlt")
    Set xlSheet1 = xlBook.Worksheets(1)

    xlSheet1.Cells(8, 2) = "SERVICE : " & cboMonth & " " & txtYear
    If Combo1.Text <> "ALL" Then
        xlSheet1.Cells(3, 4) = "SERVICE ADVISOR: " & Combo1
       'xlSheet2.Cells(3, 4) = "SERVICE ADVISOR: " & Combo1
    End If

'    SQLTXT = "SELECT ROW_NUMBER() OVER (ORDER BY DTE_FINISHED) AS [NO],FIRSTNAME,LASTNAME," & vbCrLf
'    SQLTXT = SQLTXT & "ADDRESS1 , ADDRESS2, HOMEPHONE, TELEPHONENO, DTE_FINISHED, Model, VIN, DTE_RECD" & vbCrLf
'    SQLTXT = SQLTXT & "From" & vbCrLf
'    SQLTXT = SQLTXT & "(" & vbCrLf
'    SQLTXT = SQLTXT & "SELECT  B.CUSCDE,B.FIRSTNAME,B.LASTNAME,B.ADDRESS1,B.ADDRESS2,B.HOMEPHONE,B.TELEPHONENO," & vbCrLf
'    SQLTXT = SQLTXT & "B.DTE_FINISHED,A.MODEL,A.VIN,A.PLATE_NO,RECD_BY,JOBTYPE,WRITER,CUSTYPE," & vbCrLf
'    SQLTXT = SQLTXT & "Case JOBTYPE" & vbCrLf
'    SQLTXT = SQLTXT & "WHEN 'PMS' THEN ISNULL(DTE_RECD,'')" & vbCrLf
'    SQLTXT = SQLTXT & "END As DTE_RECD" & vbCrLf
'    SQLTXT = SQLTXT & "FROM CSMS_CUSVEH A INNER JOIN" & vbCrLf
'    SQLTXT = SQLTXT & "(" & vbCrLf
'    SQLTXT = SQLTXT & "SELECT X.CUSCDE,ISNULL(FIRSTNAME,'') AS FIRSTNAME,ISNULL(X.LASTNAME,'') AS LASTNAME,Y.PLATE_NO," & vbCrLf
'    SQLTXT = SQLTXT & "ISNULL(X.CUSTOMERADD,'') AS ADDRESS1,ISNULL(X.PROVINCIALADD,'') AS ADDRESS2,X.CUSTYPE," & vbCrLf
'    SQLTXT = SQLTXT & "ISNULL(HOMEPHONE,'') AS HOMEPHONE,ISNULL(TELEPHONENO,'') AS TELEPHONENO, DTE_COMP AS DTE_FINISHED,JOBTYPE,DTE_RECD,RECD_BY,WRITER" & vbCrLf
'    SQLTXT = SQLTXT & "FROM ALL_CUSTOMER_TABLE X INNER JOIN" & vbCrLf
'    SQLTXT = SQLTXT & "(" & vbCrLf
'    SQLTXT = SQLTXT & "SELECT WRITER,A.REP_OR,PLATE_NO,ACCT_NO,MAX(ISNULL(DTE_COMP,'')) AS DTE_COMP,MAX(ISNULL(DTE_RECD,'')) AS DTE_RECD,JOBTYPE,RECD_BY  FROM" & vbCrLf
'    SQLTXT = SQLTXT & "(" & vbCrLf
'    SQLTXT = SQLTXT & "SELECT REP_OR,A.PLATE_NO,A.ACCT_NO,A.DTE_COMP,A.DTE_RECD,A.RECD_BY ,B.WRITER FROM CSMS_REPOR A INNER JOIN CSMS_REPAIRORDER B" & vbCrLf
'    SQLTXT = SQLTXT & "ON A.REP_OR = B.RO_NO AND A.PLATE_NO = B.PLATE_NO" & vbCrLf
'    SQLTXT = SQLTXT & "WHERE A.TRANSTYPE ='R') A LEFT OUTER JOIN" & vbCrLf
'    SQLTXT = SQLTXT & "(" & vbCrLf
'    SQLTXT = SQLTXT & "SELECT REP_OR,JOBTYPE,TECHCODE FROM CSMS_RO_DET WHERE LIVIL = '1' AND JOBTYPE = 'PMS'" & vbCrLf
'    SQLTXT = SQLTXT & ") B ON A.REP_OR = B.REP_OR WHERE MONTH(DTE_COMP)= " & What_month(cboMonth) & " AND YEAR(DTE_COMP)=" & txtYear & "" & vbCrLf
'    SQLTXT = SQLTXT & "GROUP BY PLATE_NO,ACCT_NO,JOBTYPE,RECD_BY,A.REP_OR,WRITER" & vbCrLf
'    SQLTXT = SQLTXT & ")Y ON X.CUSCDE = Y.ACCT_NO" & vbCrLf
'    SQLTXT = SQLTXT & ")B ON A.PLATE_NO = B.PLATE_NO" & vbCrLf
'    SQLTXT = SQLTXT & ")T WHERE " & vbCrLf


'     If Combo1.Text <> "ALL" Then
'        SQLTXT = SQLTXT & " WRITER =" & N2Str2Null(Combo1) & " AND " & vbCrLf
'    End If
'    SQLTXT = SQLTXT & " CUSTYPE = 'P'  ORDER BY DTE_FINISHED "
'
'    Set rstmp = gconDMIS.Execute(SQLTXT)
'
'    If Not (rstmp.EOF And rstmp.BOF) Then
'        xlSheet1.Cells(10, 1).CopyFromRecordset rstmp
'    Else
'        MsgSpeechBox " There Are No Records for the Specified Date"
'        Exit Sub
'    End If
    
    Set rstmp = Nothing
    SQLTXT = ""


    SQLTXT = "  SELECT ROW_NUMBER() OVER (ORDER BY DTE_FINISHED) AS [NO],DTE_RECD, " & vbCrLf
    SQLTXT = SQLTXT & "LASTNAME , FIRSTNAME, (ADDRESS1) as Complete_Address ,EMAIL, (isnull(HOMEPHONE,'NONE') + ' / ' + isnull(TELEPHONENO,'NONE')) as contact_number,PLATE_NO,VIN,Model,D_SOLD,Cus_type,CITY,DTE_REL,KM_RDG,null as remarks" & vbCrLf
    SQLTXT = SQLTXT & " From " & vbCrLf
    SQLTXT = SQLTXT & " ( " & vbCrLf
    SQLTXT = SQLTXT & " SELECT  B.CUSCDE,B.FIRSTNAME,B.LASTNAME,B.EMAIL,B.ADDRESS1,B.ADDRESS2,B.HOMEPHONE,B.TELEPHONENO," & vbCrLf
    SQLTXT = SQLTXT & " B.DTE_FINISHED,A.MODEL,A.VIN,A.PLATE_NO,RECD_BY,JOBTYPE,WRITER,CUSTYPE,D_SOLD,Cus_type,CITY,DTE_REL,KM_RDG," & vbCrLf
    SQLTXT = SQLTXT & " Case JOBTYPE " & vbCrLf
    SQLTXT = SQLTXT & " when 'PMS' THEN ISNULL(DTE_RECD,'')" & vbCrLf
    SQLTXT = SQLTXT & " end As DTE_RECD" & vbCrLf
    SQLTXT = SQLTXT & " FROM CSMS_CUSVEH A INNER JOIN" & vbCrLf
    SQLTXT = SQLTXT & " ( " & vbCrLf
    SQLTXT = SQLTXT & " SELECT X.CUSCDE,ISNULL(FIRSTNAME,'') AS FIRSTNAME,isnull(email,'') as EMAIL,ISNULL(X.LASTNAME,'') AS LASTNAME,Y.PLATE_NO," & vbCrLf
    SQLTXT = SQLTXT & " ISNULL(X.CUSTOMERADD,'') AS ADDRESS1,ISNULL(X.PROVINCIALADD,'') AS ADDRESS2,X.CUSTYPE," & vbCrLf
    SQLTXT = SQLTXT & " ISNULL(HOMEPHONE,'') AS HOMEPHONE,ISNULL(TELEPHONENO,'') AS TELEPHONENO, DTE_COMP AS DTE_FINISHED,JOBTYPE,DTE_RECD,RECD_BY,WRITER,CITY,DTE_REL,KM_RDG," & vbCrLf
    SQLTXT = SQLTXT & " Case CUSTYPE" & vbCrLf
    SQLTXT = SQLTXT & " when 'P' then 'PERSONAL'" & vbCrLf
    SQLTXT = SQLTXT & " when 'C' then 'Company/Agency'" & vbCrLf
    SQLTXT = SQLTXT & " when 'F' then 'Fleet Account'" & vbCrLf
    SQLTXT = SQLTXT & " when 'G' then 'Government'" & vbCrLf
    SQLTXT = SQLTXT & " end As Cus_type" & vbCrLf
    SQLTXT = SQLTXT & " FROM ALL_CUSTOMER_TABLE X INNER JOIN" & vbCrLf
    SQLTXT = SQLTXT & " ( " & vbCrLf
    SQLTXT = SQLTXT & " SELECT WRITER,A.REP_OR,PLATE_NO,ACCT_NO,MAX(ISNULL(DTE_COMP,'')) AS DTE_COMP,MAX(ISNULL(DTE_RECD,'')) AS DTE_RECD,JOBTYPE,RECD_BY ,MAX(ISNULL(DTE_REL,'')) AS DTE_REL,max(isnull(KM_RDG,'')) as KM_RDG FROM" & vbCrLf
    SQLTXT = SQLTXT & " ( " & vbCrLf
    SQLTXT = SQLTXT & " SELECT REP_OR,A.PLATE_NO,A.ACCT_NO,A.DTE_COMP,A.DTE_RECD,DTE_REL,A.RECD_BY ,B.WRITER,A.KM_RDG FROM CSMS_REPOR A INNER JOIN CSMS_REPAIRORDER B" & vbCrLf
    SQLTXT = SQLTXT & " ON A.REP_OR = B.RO_NO AND A.PLATE_NO = B.PLATE_NO" & vbCrLf
    SQLTXT = SQLTXT & " WHERE A.TRANSTYPE ='R') A LEFT OUTER JOIN" & vbCrLf
    SQLTXT = SQLTXT & " ( " & vbCrLf
    SQLTXT = SQLTXT & " SELECT REP_OR,JOBTYPE,TECHCODE FROM CSMS_RO_DET WHERE LIVIL = '1' AND JOBTYPE = 'PMS'" & vbCrLf
    SQLTXT = SQLTXT & " )       B ON A.REP_OR = B.REP_OR WHERE MONTH(DTE_COMP)= " & What_month(cboMonth) & " AND YEAR(DTE_COMP)=" & txtYear & "" & vbCrLf
    SQLTXT = SQLTXT & " GROUP BY PLATE_NO,ACCT_NO,JOBTYPE,RECD_BY,A.REP_OR,WRITER" & vbCrLf
    SQLTXT = SQLTXT & " )Y ON X.CUSCDE = Y.ACCT_NO" & vbCrLf
    SQLTXT = SQLTXT & " )B ON A.PLATE_NO = B.PLATE_NO" & vbCrLf
    SQLTXT = SQLTXT & " )T WHERE" & vbCrLf
     
     If Combo1.Text <> "ALL" Then
        SQLTXT = SQLTXT & " WRITER =" & N2Str2Null(Combo1) & " AND " & vbCrLf
    End If
    SQLTXT = SQLTXT & " CUSTYPE IN ('F','P C','C P','G','C', ' ') ORDER BY DTE_FINISHED "
    
    Set rstmp = gconDMIS.Execute(SQLTXT)
    
    If Not (rstmp.EOF And rstmp.BOF) Then
        xlSheet1.Cells(10, 1).CopyFromRecordset rstmp
    Else
        MsgSpeechBox " There Are No Records for the Specified Date"
        Exit Sub
    End If
    
    
    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlSheet2 = Nothing
    Set xlApp = Nothing
    Set rstmp = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PRINTSALES
' DateTime  : 10/24/2007 15:35
' Author    : Ashish
' Purpose   :
'---------------------------------------------------------------------------------------
'

Sub ServiceReport()
    ReportType = "SERVICE"
End Sub


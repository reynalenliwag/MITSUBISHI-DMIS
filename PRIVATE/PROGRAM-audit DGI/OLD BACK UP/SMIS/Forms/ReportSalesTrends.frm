VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_SalesLead2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Trend"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportSalesTrends.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   3420
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
      Left            =   1800
      MouseIcon       =   "ReportSalesTrends.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportSalesTrends.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   1650
      Width           =   885
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Lead Source"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1140
      TabIndex        =   5
      Top             =   1350
      Width           =   1785
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Age Group"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1140
      TabIndex        =   4
      Top             =   1080
      Width           =   1785
   End
   Begin VB.OptionButton Option2 
      Caption         =   "by Color"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1140
      TabIndex        =   3
      Top             =   810
      Width           =   1785
   End
   Begin VB.OptionButton Option1 
      Caption         =   "by Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1140
      TabIndex        =   2
      Top             =   570
      Value           =   -1  'True
      Width           =   1785
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "9999"
      Top             =   60
      Width           =   1185
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   30
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "List of Registrations"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
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
      Left            =   930
      MouseIcon       =   "ReportSalesTrends.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportSalesTrends.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   1650
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
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
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmSMIS_Report_SalesLead2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    Dim SQL                                                           As String
    Set rsPurchAgree = New ADODB.Recordset
    SQL = "SELECT * FROM SMIS_MRRINV INNER JOIN ALL_CUSTOMER ON CUSCDE=CUSTOMERCODE"
    SQL = SQL & "  WHERE year(DateReleased) = " & txtYear
    rsPurchAgree.Open SQL, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptGenREP.Formulas(2) = "Month = '" & txtYear & "'"

        If Option1.Value = True Then
            rptGenREP.WindowTitle = "Sales Trend By Gender"
            rptGenREP.ReportTitle = "Sales Trend By Gender"
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "SalesTrend1.rpt", " year({SQ.DateReleased}) = " & txtYear, DMIS_REPORT_Connection, 1
        ElseIf Option2.Value = True Then
            rptGenREP.WindowTitle = "Sales Trend By Color"
            rptGenREP.ReportTitle = "Sales Trend By  Color"
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "SalesTrend2.rpt", " YEAR({SQ.DateReleased}) = " & txtYear, DMIS_REPORT_Connection, 1

        ElseIf Option3.Value = True Then
            rptGenREP.WindowTitle = "Sales Trend According to Age Group"
            rptGenREP.ReportTitle = "Sales Trend According to Age Group"
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "SalesTrend3.rpt", "  year({SQ.DateReleased}) = " & txtYear, DMIS_REPORT_Connection, 1
        ElseIf Option4.Value = True Then
            rptGenREP.WindowTitle = "Sales Trend By Lead Source"
            rptGenREP.ReportTitle = "Sales Trend By Lead Source"
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "SalesTrend4.rpt", " year({SQ.DateReleased}) = " & txtYear, DMIS_REPORT_Connection, 1

        End If

        
        If Option1.Value = True Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES TREND ", "", "", "", "BY GENDER " & txtYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

        ElseIf Option2.Value = True Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES TREND ", "", "", "", "BY COLOR " & txtYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ElseIf Option3.Value = True Then
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES TREND ", "", "", "", "BY AGE GROUP " & txtYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Else
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES TREND ", "", "", "", "BY LEADSOURCE " & txtYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        End If

        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Year of " & txtYear
    End If
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

'
'Private Sub Command1_Click()
'    Dim RsLeadClass                    As ADODB.Recordset
'    Dim RsCRISProspect                 As ADODB.Recordset
'    Dim TEMPSQL                        As String
'    Dim SO, LO, TD, QU, LETTER, EMAIL, APPOINTMENT As Integer
'    Dim LeadClass                      As String
'    Set RsLeadClass = gconDMIS.Execute("SELECT * FROM CRIS_LEADCLASS")
'
'    'UPDATE DMIS.dbo.CRIS_LGM
'    'SET
'    'ProspectName=<ProspectName,varchar(50),>,
'    'CUSCDE=<CUSCDE,char(8),>,
'    'CUSNAME=<CUSNAME,varchar(100),>,
'    'ADDRESS=<ADDRESS,varchar(100),>,
'    'LEADCLASS=<LEADCLASS,varchar(80),>, MODELINQUIRY=<MODELINQUIRY,varchar(80),>,
'    'STATE=<STATE,varchar(20),>, SAE=<SAE,varchar(100),>
'    'WHERE <Search conditions,,>
'
'    gconDMIS.Execute ("DELETE FROM CRIS_LGM")
'
'
'    TEMPSQL = "INSERT INTO CRIS_LGM"
'    TEMPSQL = TEMPSQL & " SELECT " & vbCrLf
'    TEMPSQL = TEMPSQL & " P.AcctName , " & vbCrLf
'    TEMPSQL = TEMPSQL & " P.CUSCDE , " & vbCrLf
'    TEMPSQL = TEMPSQL & " (SELECT CV.CUSTOMERNAME from  CRIS_vw_ALLprofile CV WHERE CV.CUSCDE=P.CUSCDE) ," & vbCrLf
'    TEMPSQL = TEMPSQL & " P.ADDRESS , " & vbCrLf
'    TEMPSQL = TEMPSQL & " NULL," & vbCrLf
'    TEMPSQL = TEMPSQL & " P.VARIANT , " & vbCrLf
'    TEMPSQL = TEMPSQL & " P.Status , " & vbCrLf
'    TEMPSQL = TEMPSQL & " P.SAE , P.ProspectID" & vbCrLf
'    TEMPSQL = TEMPSQL & " FROM CRIS_PROSPECTS P" & vbCrLf
'    gconDMIS.Execute TEMPSQL
'    TEMPSQL = vbNullString
'
'
'    'gconDMIS.Execute ("INSERT INTO CRIS_LGM")
'    While Not RsLeadClass.EOF
'        LeadClass = RsLeadClass("LCLASS")
'        SO = RsLeadClass("SO")
'        LO = RsLeadClass("LO")
'        TD = RsLeadClass("TD")
'        QU = RsLeadClass("QU")
'        LETTER = RsLeadClass("LETTER")
'        EMAIL = RsLeadClass("EMAIL")
'        APPOINTMENT = RsLeadClass("APPOINTMENT")
'        ''FIND SUCH PROSPECT NOW
'        TEMPSQL = "select * from CRIS_PROSPECTS WHERE "
'        TEMPSQL = TEMPSQL & " (IsDate (LOGSO) =" & x(SO)
'        TEMPSQL = TEMPSQL & " AND IsDate (LOGAPPLICATION) =" & x(LO)
'        TEMPSQL = TEMPSQL & " AND IsDate (LOGTESTDRIVE) =" & x(TD)
'        TEMPSQL = TEMPSQL & " AND IsDate (LOGQUOTE) =" & x(QU)
'        TEMPSQL = TEMPSQL & " AND IsDate (LOGLETTER) =" & x(LETTER)
'        TEMPSQL = TEMPSQL & " AND IsDate (LOGEMAIL) =" & x(EMAIL)
'        TEMPSQL = TEMPSQL & " AND IsDate (LOGAPPOINTMENT) =" & x(APPOINTMENT) & ")"
'
'
'        Set RsCRISProspect = gconDMIS.Execute(TEMPSQL)
'        'UPDATE DMIS.dbo.CRIS_LGM
'        'SET
'        'ProspectName=<ProspectName,varchar(50),>,
'        'CUSCDE=<CUSCDE,char(8),>,
'        'CUSNAME=<CUSNAME,varchar(100),>,
'        'ADDRESS=<ADDRESS,varchar(100),>,
'        'LEADCLASS=<LEADCLASS,varchar(80),>, MODELINQUIRY=<MODELINQUIRY,varchar(80),>,
'        'STATE=<STATE,varchar(20),>, SAE=<SAE,varchar(100),>
'        'WHERE <Search conditions,,>
'
'        If Not (RsCRISProspect.EOF Or RsCRISProspect.BOF) Then
'            gconDMIS.Execute ("UPDATE CRIS_LGM SET LEADCLASS=" & N2Str2Null(RsLeadClass("LCLASS")) & " WHERE ProspectID=" & RsCRISProspect("PROSPECTID"))
'
'        End If
'
'
'        RsLeadClass.MoveNext
'    Wend
'
'End Sub
'Function x(ss As Variant) As Integer
'    If ss = True Then: x = 1: Else x = 0
'
'End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES TREND )"
            Call frmALL_AuditInquiry.DisplayHistory("", "SALES TREND ", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"


    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub


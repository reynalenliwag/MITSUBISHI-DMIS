VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_SAEPer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAE Performance"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportSAEPer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   4545
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1725
      Left            =   0
      ScaleHeight     =   1725
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   1710
      Width           =   4515
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
         Left            =   2160
         MouseIcon       =   "ReportSAEPer.frx":0E42
         MousePointer    =   99  'Custom
         Picture         =   "ReportSAEPer.frx":0F94
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Close Window"
         Top             =   780
         Width           =   885
      End
      Begin VB.ComboBox cboYear 
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
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   1365
      End
      Begin VB.ComboBox cboMonth2 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   1455
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
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   1515
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
         Left            =   1290
         MouseIcon       =   "ReportSAEPer.frx":13DF
         MousePointer    =   99  'Custom
         Picture         =   "ReportSAEPer.frx":1531
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print Report"
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   -30
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   30
         Width           =   300
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CheckBox Check1 
         Caption         =   "Add Date Range"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   1410
         Value           =   1  'Checked
         Width           =   3465
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   990
         Width           =   4395
      End
      Begin VB.ComboBox cboSalesAE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   4395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Models"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   13
         Top             =   690
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SAE Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   7020
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Sales Executive Performance"
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
End
Attribute VB_Name = "frmSMIS_Report_SAEPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset
Dim rsSrep                                                            As ADODB.Recordset
Dim REPORTNAME                                                        As String

Sub FillcboSAE()
    Set rsSrep = New ADODB.Recordset
    rsSrep.Open "select * from SMIS_vw_Srep ", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSrep.EOF And Not rsSrep.BOF Then
        cboSalesAE.Clear
        cboSalesAE.AddItem "ALL"
        cboSalesAE.Text = "ALL"
        rsSrep.MoveFirst
        Do While Not rsSrep.EOF
            cboSalesAE.AddItem Null2String(rsSrep!Name)
            rsSrep.MoveNext
        Loop
    End If

    Combo_Loadval Combo1, gconDMIS.Execute("SELECT DISTINCT UPPER(MODEL) MODEL  FROM ALL_MODEL")
    Combo1.AddItem "ALL"
    Combo1.Text = "ALL"
End Sub

Sub ShowSAEVsPROSPECT()
    REPORTNAME = "SAECLOSING"
End Sub

Sub ShowSAETeamVsPROSPECT()
    REPORTNAME = "SAECLOSINGTEAM"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim FILTER                                                        As String
    On Error GoTo errcode:
    If What_month(cboMonth) > What_month(cboMonth2) Then
        MsgSpeechBox "Error In From - To Months"
        Exit Sub
    End If
    rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    If REPORTNAME = "SAECLOSING" Then

        'PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeclosing.rpt", "month({CP.LogInitialInquiry}) >= " & What_month(cboMonth) & " AND month({CP.LogInitialInquiry}) <= " & What_month(cboMonth2) & " AND YEAR({CP.LogInitialInquiry}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        FILTER = "(YEAR({CP.LOGCLOSINGDATE})=" & cboYear.Text & " AND MONTH({CP.LOGCLOSINGDATE})>=" & What_month(cboMonth) & " Or Year({CP.LOGINITIALINQUIRY})>= " & cboYear & " And Month({CP.LOGINITIALINQUIRY}) >= " & What_month(cboMonth) & ") AND (YEAR({CP.LOGCLOSINGDATE})=" & cboYear.Text & " AND MONTH({CP.LOGCLOSINGDATE})<=" & What_month(cboMonth) & " Or Year({CP.LOGINITIALINQUIRY})<= " & cboYear & " And Month({CP.LOGINITIALINQUIRY}) <= " & What_month(cboMonth) & ")"

        'FILTER = "YEAR({CP.LOGCLOSINGDATE})=2008 AND MONTH({CP.LOGCLOSINGDATE})=7 OR YEAR({CP.LOGINITIALINQUIRY})=2008 AND MONTH({CP.LOGINITIALINQUIRY})=7"

        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae\saeclosing.rpt", FILTER, DMIS_REPORT_Connection, 1
        rptGenREP.PageZoom 90

    ElseIf REPORTNAME = "SAECLOSINGTEAM" Then
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeclosingbyteam.rpt", "month({CP.LogInitialInquiry}) >= " & What_month(cboMonth) & " AND month({CP.LogInitialInquiry}) <= " & What_month(cboMonth2) & " AND YEAR({CP.LogInitialInquiry}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        rptGenREP.PageZoom 90
    Else
        If Check1.Value = 1 Then
            If cboSalesAE.Text <> "ALL" And Combo1.Text <> "ALL" Then
                Set rsPurchAgree = New ADODB.Recordset
                rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = " & cboYear.Text & " AND month(datereleased) >= " & What_month(cboMonth) & " AND month(datereleased) <=" & What_month(cboMonth2) & " AND salesae = '" & cboSalesAE.Text & "' and model='" & Combo1 & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboSalesAE.Text = "ALL" And Combo1.Text <> "ALL" Then
                Set rsPurchAgree = New ADODB.Recordset
                rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = " & cboYear.Text & " AND month(datereleased) >= " & What_month(cboMonth) & " AND month(datereleased) <=" & What_month(cboMonth2) & " and model='" & Combo1 & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboSalesAE.Text <> "ALL" And Combo1.Text = "ALL" Then
                Set rsPurchAgree = New ADODB.Recordset
                rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = " & cboYear.Text & " AND month(datereleased) >= " & What_month(cboMonth) & " AND month(datereleased) <=" & What_month(cboMonth2) & " AND salesae = '" & cboSalesAE.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                Set rsPurchAgree = New ADODB.Recordset
                rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = " & cboYear.Text & " AND month(datereleased) >= " & What_month(cboMonth) & " AND month(datereleased) <=" & What_month(cboMonth2), gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            If cboSalesAE.Text <> "ALL" And Combo1.Text <> "ALL" Then
                Set rsPurchAgree = New ADODB.Recordset
                rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE salesae = '" & cboSalesAE.Text & "' and model='" & Combo1 & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboSalesAE.Text = "ALL" And Combo1.Text <> "ALL" Then
                Set rsPurchAgree = New ADODB.Recordset
                rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE model='" & Combo1 & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboSalesAE.Text <> "ALL" And Combo1.Text = "ALL" Then
                Set rsPurchAgree = New ADODB.Recordset
                rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE salesae = '" & cboSalesAE.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                Set rsPurchAgree = New ADODB.Recordset
                rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = " & cboYear.Text & " AND month(datereleased) >= " & What_month(cboMonth) & " AND month(datereleased) <=" & What_month(cboMonth2), gconDMIS, adOpenForwardOnly, adLockReadOnly

            End If


        End If
        If Not rsPurchAgree.BOF And Not rsPurchAgree.EOF Then
            Screen.MousePointer = 11
            If Check1.Value = 1 Then
                If cboSalesAE.Text <> "ALL" And Combo1.Text <> "ALL" Then
                    PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeper.rpt", "month({purchagree.datereleased}) >= " & What_month(cboMonth) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonth2) & " AND YEAR({purchagree.datereleased}) = " & cboYear.Text & " AND {purchagree.model} = '" & Combo1.Text & "'" & " AND {purchagree.salesae} = '" & cboSalesAE.Text & "' ", DMIS_REPORT_Connection, 1
                ElseIf cboSalesAE.Text = "ALL" And Combo1.Text <> "ALL" Then
                    PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeper.rpt", "month({purchagree.datereleased}) >= " & What_month(cboMonth) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonth2) & " AND YEAR({purchagree.datereleased}) = " & cboYear.Text & " AND {purchagree.model} = '" & Combo1.Text & "'", DMIS_REPORT_Connection, 1
                ElseIf cboSalesAE.Text <> "ALL" And Combo1.Text = "ALL" Then
                    PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeper.rpt", "month({purchagree.datereleased}) >= " & What_month(cboMonth) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonth2) & " AND YEAR({purchagree.datereleased}) = " & cboYear.Text & " AND {purchagree.salesae} = '" & cboSalesAE.Text & "'", DMIS_REPORT_Connection, 1
                Else
                    ' PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeper.rpt", "month({purchagree.datereleased}) >= " & What_month(cboMonth) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonth2) & " AND YEAR({purchagree.datereleased}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                    PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae\saeper.rpt", "Month({MRRINV.datereleased})>=" & What_month(cboMonth) & " AND MONTH({MRRINV.datereleased})<=" & What_month(cboMonth2) & " AND Year({MRRINV.datereleased})=" & cboYear.Text & "", DMIS_REPORT_Connection, 1
                End If
            Else
                If cboSalesAE.Text <> "ALL" And Combo1.Text <> "ALL" Then
                    PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeper.rpt", "{purchagree.model} = '" & Combo1.Text & "'" & " AND {purchagree.salesae} = '" & cboSalesAE.Text & "' ", DMIS_REPORT_Connection, 1
                ElseIf cboSalesAE.Text = "ALL" And Combo1.Text <> "ALL" Then
                    PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeper.rpt", "{purchagree.model} = '" & Combo1.Text & "'", DMIS_REPORT_Connection, 1
                ElseIf cboSalesAE.Text <> "ALL" And Combo1.Text = "ALL" Then
                    PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae/saeper.rpt", "{purchagree.salesae} = '" & cboSalesAE.Text & "'", DMIS_REPORT_Connection, 1
                Else
                    PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "sae\saeper.rpt", "", DMIS_REPORT_Connection, 1
                End If
            End If
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 5:00
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
             Call NEW_LogAudit("V", "SALES EXECUTIVE PERFORMANCE", "", "", "", cboSalesAE & " " & Combo1 & "FROM " & cboMonth & "TO " & cboMonth2 & " " & cboYear, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for the Month of " & cboMonth.Text
        End If
    End If
    Exit Sub
errcode:
    ShowVBError

End Sub

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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES EXECUTIVE PERFORMANCE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SALES EXECUTIVE PERFORMANCE", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    fillcbomonth cboMonth2
    fillcbomoreyear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboMonth2.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)


    If REPORTNAME = "SAECLOSINGTEAM" Then
        Picture1.Visible = False
        Me.Caption = "Sales Executive Performance Sales Productivity Per Group"
    ElseIf REPORTNAME = "SAECLOSING" Then
        Picture1.Visible = False
        Me.Caption = "Sales Executive Performance Sales Productivity Per SC"
    Else
        Me.Caption = "SAE Performance"
        Picture1.Visible = True
        FillcboSAE
    End If

    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    REPORTNAME = ""
End Sub


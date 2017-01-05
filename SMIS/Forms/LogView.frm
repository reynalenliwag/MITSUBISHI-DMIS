VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Inquiry_ViewLog 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   9210
   Begin XtremeReportControl.ReportControl lvInq 
      Height          =   4605
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   9135
      _Version        =   655364
      _ExtentX        =   16113
      _ExtentY        =   8123
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnSort =   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1350
      TabIndex        =   5
      Top             =   450
      Width           =   4515
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   8490
      MouseIcon       =   "LogView.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "LogView.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   5520
      Width           =   705
   End
   Begin Crystal.CrystalReport rptLog 
      Left            =   180
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Sales Executive Performance"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowNavigationCtls=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   795
      Left            =   7800
      MouseIcon       =   "LogView.frx":079A
      MousePointer    =   99  'Custom
      Picture         =   "LogView.frx":08EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   705
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   795
      Left            =   7110
      MouseIcon       =   "LogView.frx":0C52
      MousePointer    =   99  'Custom
      Picture         =   "LogView.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Select "
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   " Search Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   510
      Width           =   4635
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9765
      _Version        =   655364
      _ExtentX        =   17224
      _ExtentY        =   741
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_ViewLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PROSPECTID                                                        As Long
Dim AcctName                                                          As String
Dim Action                                                            As String
Dim SAEID                                                             As Long
Dim SAENAME                                                           As String
Dim CustomerCode                                                      As String
Event EventSelected(EventName As String, xProspectID As Long, AppointmentID As Long)

Sub ShowCustomerAppointment(xProspectID As Long, CustomersName As String)
    Action = "PROSPECTAPP"
    PROSPECTID = xProspectID
    ShortcutCaption.Caption = "SHOWING SALES APPPOINTMENT DETAIL OF " & CustomersName
End Sub

Sub SHOWCUSTOMERLOG(XCUSTOMERCODE As String, CUSTOMERNAME As String)
    CustomerCode = XCUSTOMERCODE
    Action = "LOG:CUSTOMER"
    ShortcutCaption.Caption = "SHOWING LOG DETAIL FOR NAME : " & CUSTOMERNAME
End Sub

Sub SHOWPROSPECTLOG(xProspectID As Long, PROSPECTNAME As String)
    PROSPECTID = xProspectID
    Action = "LOG:PROSPECT"
    ShortcutCaption.Caption = "SHOWING LOG DETAIL FOR PROSPECT: " & PROSPECTNAME
End Sub

Sub SHOWSAEAPPOINTMENTDETAIL(XSAENAME As String)
    Dim XSAEID                                                        As Integer
    Action = "SAEAPPDETAIL"
    SAEID = XSAEID
    SAENAME = XSAENAME
    ShortcutCaption.Caption = "SHOWING SALES APPPOINTMENT DETAIL OF : " & XSAENAME
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Select Case Action
        Case "LOG:PROSPECT"
            rptLog.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptLog.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptLog, SMIS_REPORT_PATH & "ProspectLog.rpt", "{CRIS_VIEWLOG.PROSPECTID}=" & PROSPECTID, DMIS_REPORT_Connection, 1
    End Select
    'MsgBox "eiu ini c form"
    NEW_LogAudit "V", "PROSPECT INQUIRY", "", N2Str2Null(PROSPECTID), "", "PROSPECTID=" & Null2String(PROSPECTID), "", ""
End Sub

Private Sub cmdSelect_Click()
    RaiseEvent EventSelected("Appointment", PROSPECTID, CLng(lvInq.SelectedRows(0).Record(9).Value))
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PROSPECT INQUIRY)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(PROSPECTID), "PROSPECT INQUIRY")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim SQL                                                           As String
    cmdSelect.Visible = False: cmdPrint.Visible = True
    ReportControlPaintManager lvInq
    lvInq.PaintManager.TextFont.Size = 9
    lvInq.PaintManager.TextFont.Name = "Tahoma"

    Select Case Action
        Case "LOG:PROSPECT"
            ReportControlAddColumnHeader lvInq, "DATE, DETAILS, PARTICULAR"
            ResizeColumnHeader lvInq, "20,50,80"
            flex_FillReportView gconDMIS.Execute("Select deyt, LogName, Particular , CSCDE, PROSPECTID, LOGID  from CRIS_VIEWLOG WHERE PROSPECTID=" & PROSPECTID & "  Order By [deyt] , SN "), lvInq
        Case "LOG:CUSTOMER"
            ReportControlAddColumnHeader lvInq, "DATE, DETAILS, PARTICULAR"
            ResizeColumnHeader lvInq, "20,50,80"

            flex_FillReportView gconDMIS.Execute("Select convert(varchar, deyt,101 ), LogName, Particular , CSCDE, PROSPECTID, LOGID  from CRIS_VIEWLOG WHERE CSCDE=" & N2Str2Null(CustomerCode) & "  Order By [deyt] , SN "), lvInq
        Case "SAEAPPDETAIL"
            With lvInq
                .Columns.Add 0, "AppointmentID", 0, True
                .Columns.Add 1, "DateTime", 0, True
                .Columns.Add 2, "Account Name", 50, True
                .Columns.Add 3, "Model", 100, True
                .Columns.Add 4, "Color", 100, True
                .Columns.Add 5, "Terms", 100, True
                .Columns.Add 6, "ExpectedPurchase", 100, False
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            End With
            SQL = " select AppointmentID, " _
                & " Convert(varchar , StartDateTime,101) as DateTime, " _
                & " (Select AcctName from CRIS_PROSPECTS where CRIS_PROSPECTS.ProspectID=CRIS_SalesAppointments.ProspectID) as [Account Name]," _
                & " Model , Color, Terms, Convert(varchar, ExpectedPurchase,101) as [Expected Purchase]  from CRIS_SalesAppointments" _
                & " Where SAE=" & N2Str2Null(SAENAME)

            flex_FillReportView gconDMIS.Execute(SQL), lvInq
        Case "PROSPECTAPP"
            Call ReportControlAddColumnHeader(lvInq, "Date, AppointmentCode, SAE, ModelInquiry, Variant, Color, ExpectedTerm, ExpectedPurchase")
            SQL = "Select  Convert(varchar, StartDateTime, 101),  SAE, Model, Make, Color, Class, Year,Terms,ExpectedPurchase , AppointmentID from CRIS_SalesAppointments WHere ProspectID=" & PROSPECTID
            flex_FillReportView gconDMIS.Execute(SQL), lvInq
            cmdSelect.Visible = True
            cmdPrint.Visible = False

    End Select
    CenterMe frmMain, Me, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PROSPECTID = 0
    AcctName = vbNullString
    Action = vbNullString
    SAEID = 0
    SAENAME = vbNullString

End Sub

Private Sub lvInq_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Index Mod 2 = 0 Then
        Metrics.BackColor = &HCEF9F5
    End If
End Sub

Private Sub Text1_Change()
    lvInq.FilterText = Text1
    lvInq.Populate
End Sub


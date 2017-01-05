VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Inquiry_ViewLog 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
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
   ScaleHeight     =   7470
   ScaleWidth      =   12270
   Begin XtremeReportControl.ReportControl lvInq 
      Height          =   6075
      Left            =   30
      TabIndex        =   0
      Top             =   540
      Width           =   12195
      _Version        =   655364
      _ExtentX        =   21511
      _ExtentY        =   10716
      _StockProps     =   64
      BorderStyle     =   2
      AllowColumnSort =   0   'False
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   11490
      MouseIcon       =   "LogView.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "LogView.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6660
      Width           =   705
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   795
      Left            =   10800
      MouseIcon       =   "LogView.frx":079A
      MousePointer    =   99  'Custom
      Picture         =   "LogView.frx":08EC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6660
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   795
      Left            =   10800
      MouseIcon       =   "LogView.frx":0C28
      MousePointer    =   99  'Custom
      Picture         =   "LogView.frx":0D7A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6660
      Width           =   705
   End
   Begin Crystal.CrystalReport rptLog 
      Left            =   0
      Top             =   0
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12285
      _Version        =   655364
      _ExtentX        =   21669
      _ExtentY        =   952
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
Dim ProspectID                                         As Long
Dim AcctName                                           As String
Dim Action                                             As String
Dim SAEID                                              As Long
Dim SAENAME                                            As String
Dim CustomerCode                                       As String
Event EventSelected(EventName As String, xProspectID As Long, AppointmentID As Long)

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Select Case Action
        Case "LOG:PROSPECT"
            PrintSQLReport rptLog, SMIS_REPORT_PATH & "ProspectLog.rpt", "{CRIS_VIEWLOG.PROSPECTID}=" & ProspectID, DMIS_REPORT_Connection, 1
    End Select

End Sub

Private Sub cmdSelect_Click()
    RaiseEvent EventSelected("Appointment", ProspectID, CLng(lvInq.SelectedRows(0).Record(9).Value))
End Sub

Private Sub Form_Load()
    cmdSelect.Visible = False: cmdPrint.Visible = True
    ReportControlPaintManager lvInq
    lvInq.PaintManager.TextFont.Size = 9
    lvInq.PaintManager.TextFont.Name = "Arial"

    Select Case Action
        Case "LOG:PROSPECT"
            ReportControlAddColumnHeader lvInq, "DATE, DETAILS, PARTICULAR"
            ResizeColumnHeader lvInq, "20,50,80"
            flex_FillReportView gconDMIS.Execute("Select Deyt, LogName, Particular , CSCDE, PROSPECTID, LOGID  from CRIS_VIEWLOG WHERE PROSPECTID=" & ProspectID & "  Order By [Deyt] , SN "), lvInq
        Case "LOG:CUSTOMER"
            ReportControlAddColumnHeader lvInq, "DATE, DETAILS, PARTICULAR"
            ResizeColumnHeader lvInq, "20,50,80"

            flex_FillReportView gconDMIS.Execute("Select Deyt, LogName, Particular , CSCDE, PROSPECTID, LOGID  from CRIS_VIEWLOG WHERE CSCDE=" & N2Str2Null(CustomerCode) & "  Order By [Date] , SN "), lvInq
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
            SQL = "Select  Convert(varchar, StartDateTime, 101),  SAE, Model, Make, Color, Class, Year,Terms,ExpectedPurchase , AppointmentID from CRIS_SalesAppointments WHere ProspectID=" & ProspectID
            flex_FillReportView gconDMIS.Execute(SQL), lvInq
            cmdSelect.Visible = True
            cmdPrint.Visible = False

    End Select
    CenterMe frmMain, Me, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProspectID = 0
    AcctName = vbNullString
    Action = vbNullString
    SAEID = 0
    SAENAME = vbNullString

End Sub

Private Sub lvInq_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
'    If Row.Record Is Nothing Then Exit Sub
'    If Action = "LOG" Then
'
'        Dim lname                      As String
'        Dim lngID                      As Long
'        lname = UCase(Item.Record(3).Value)
'        lngID = Item.Record(0).Value
'
'
'Debug.Print Item.Record(3).Value
'        If lname = "IN BOUND CALLS" Or lname = "OUT BOUND CALLS" Then
'            'frmSMISLogCall.Show
'        ElseIf lname = "QUOTATION" Then
'            '
'        ElseIf lname = "IN BOUND EMAIL" Or lname = "OUT BOUND EMAIL" Then
'        ElseIf lname = "VISITS" Then
'            '
'        ElseIf lname = "SALES ORDER" Then
'            Load frmSMIS_Trans_SalesOrder
'            Call frmSMIS_Trans_SalesOrder.SearchID(lngID)
'            frmSMIS_Trans_SalesOrder.Show
'
'        ElseIf lname = "SALES INVOICE" Then
'            '
'        Else
'
'        End If
'    End If

End Sub

Sub ShowCustomerAppointment(xProspectID As Long, CustomersName As String)
    Action = "PROSPECTAPP"
    ProspectID = xProspectID
    ShortcutCaption.Caption = "SHOWING SALES APPPOINTMENT DETAIL OF " & CustomersName
End Sub

Sub SHOWCUSTOMERLOG(XCUSTOMERCODE As String, CustomerName As String)
    CustomerCode = XCUSTOMERCODE
    Action = "LOG:CUSTOMER"
    ShortcutCaption.Caption = "SHOWING LOG DETAIL FOR NAME : " & CustomerName
End Sub

Sub SHOWPROSPECTLOG(xProspectID As Long, PROSPECTNAME As String)
    ProspectID = xProspectID
    Action = "LOG:PROSPECT"
    ShortcutCaption.Caption = "SHOWING LOG DETAIL FOR PROSPECT: " & PROSPECTNAME
End Sub

Sub SHOWSAEAPPOINTMENTDETAIL(XSAENAME As String)
    Action = "SAEAPPDETAIL"
    SAEID = XSAEID
    SAENAME = XSAENAME
    ShortcutCaption.Caption = "SHOWING SALES APPPOINTMENT DETAIL OF : " & XSAENAME
End Sub


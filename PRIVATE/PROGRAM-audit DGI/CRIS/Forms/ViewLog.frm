VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_ViewLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl lvInq 
      Height          =   6495
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   9945
      _Version        =   655364
      _ExtentX        =   17542
      _ExtentY        =   11456
      _StockProps     =   64
      BorderStyle     =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   465
      Left            =   7320
      TabIndex        =   3
      Top             =   6990
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   465
      Left            =   8625
      TabIndex        =   2
      Top             =   6990
      Width           =   1245
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9975
      _Version        =   655364
      _ExtentX        =   17595
      _ExtentY        =   714
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
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmCRIS_ViewLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ProspectID                               As Long
Dim AcctName                                 As String
Dim Action As String
Dim SAEID As Long
Dim SAENAME As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    lvInq.PrintOptions.BlackWhiteContrast = 0
    lvInq.PrintOptions.Header.Font.Size = 10
    lvInq.PrintOptions.Header.Font.Underline = True
    lvInq.PrintOptions.Header.Font.Bold = True
    lvInq.PrintOptions.Header.TextCenter = " Log Details " & vbCrLf & AcctName
    lvInq.PrintPreview True

End Sub

Private Sub Form_Load()
Dim temprs                               As ADODB.Recordset



    With lvInq                                                '''''''''''UI
        '.PaintManager.HorizontalGridStyle = xtpGridSmallDots  'xtpGridNoLines
        '.PaintManager.HighlightBackColor = RGB(255, 245, 255)
        '.PaintManager.ShadeSortColor = RGB(229, 229, 229)
        .PaintManager.HideSelection = True
        
        '.PaintManager.HighlightBackColor = RGB(0, 0, 0)
        '.PaintManager.VerticalGridStyle = xtpGridNoLines      ' xtpGridSmallDots
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
    End With
Select Case Action
Case "LOG"
    With lvInq
        .Columns.Add 0, "LOGID", 0, True
        .Columns.Add 1, "ProspectID", 0, True
        .Columns.Add 2, "Date", 50, True
        .Columns.Add 3, "LOGNAME", 80, True
        .Columns.Add 4, "Particular", 100, True
        .Columns(0).Visible = False
        .Columns(1).Visible = False
    End With
    SQL = ("Select LogID, ProspectID, DateEmail  as [Date]    , bound  + ' EMAIL' as LogName, Subject as Particular from CRIS_PROSPECT_EMAIL where ProspectID=@PXID " & vbCrLf & _
           "UNION  " & vbCrLf & _
           "SELECT LogID, ProspectID, DateTimeCall , Bound + ' Calls' , Subject FROM CRIS_Prospect_Calls where ProspectID=@PXID " & vbCrLf & _
           "UNION " & _
           "SELECT LogID, ProspectID, DateTimeVisit, 'Visit' as LogName , Comments  FROM CRIS_Prospect_Visits where ProspectID=@PXID " & vbCrLf & _
           "UNION " & _
           "SELECT QuotationID, ProspectID, QuoteDate,'Quotation', QuotationDescription + ' Cost:' + cast(TotalAmount as varchar)  FROM CRIS_Quote_Header where ProspectID=@PXID " & vbCrLf & _
           "UNION " & vbCrLf & _
           "SELECT SO_NO , ProspectID, Deyt,'Sales Order', Model + ' Amount:' + cast(NetSalesPrice as varchar)  FROM SMIS_SalesOrder where ProspectID=@PXID" & vbCrLf & _
           "Union " & vbCrLf & _
           "SELECT APL_NO , ProspectID, DateApplied ,'Loan Application', Ind_LoanApl_UnitModel + ' Status:' + Status  FROM SMIS_LoanIndiv where ProspectID=@PXID" & vbCrLf & _
           "order by Date, LogName  ")
           
    SQL = Replace(SQL, "@PXID", ProspectID)
    flex_FillReportView gconDMIS.Execute(SQL), lvInq
    
Case "SAEAPPDETAIL"
    With lvInq
        .Columns.Add 0, "AppointmentID", 0, True
        .Columns.Add 1, "DateTime", 0, True
        .Columns.Add 2, "Account Name", 50, True
        .Columns.Add 3, "Model", 100, True
        .Columns.Add 4, "Make", 100, True
        .Columns.Add 5, "Color", 100, True
        .Columns.Add 6, "Terms", 100, True
        .Columns.Add 7, "ExpectedPurchase", 100, False
        
        .Columns(0).Visible = False
        .Columns(1).Visible = False
    End With
SQL = " select AppointmentID, " _
& " Convert(varchar , StartDateTime,101) as DateTime, " _
& " (Select AcctName from CRIS_PROSPECTS where CRIS_PROSPECTS.ProspectID=CRIS_SalesAppointments.ProspectID) as [Account Name]," _
& " Model , Make, Color, Terms, Convert(varchar, ExpectedPurchase,101) as [Expected Purchase]  from CRIS_SalesAppointments" _
& " Where SAE=" & SAEID
   
   flex_FillReportView gconDMIS.Execute(SQL), lvInq
End Select
End Sub

Sub ShowReport(xProspectID As Long, xAcctName As String)
    ProspectID = xProspectID
    AcctName = xAcctName
    Action = "LOG"
    ShortcutCaption.caption = "Showing Log Detail for Account Name : " & xAcctName
    
End Sub


Sub ShowSAEAppointmentDetail(xSAEID As Long, xSAENAME As String)
    Action = "SAEAPPDETAIL"
    SAEID = xSAEID
    SAENAME = xSAENAME
    ShortcutCaption.caption = "Showing Sales Apppointment Detail of : " & xSAENAME
End Sub

Private Sub Form_Unload(Cancel As Integer)
ProspectID = 0
AcctName = vbNullString
Action = vbNullString
SAEID = 0
SAENAME = vbNullString

End Sub


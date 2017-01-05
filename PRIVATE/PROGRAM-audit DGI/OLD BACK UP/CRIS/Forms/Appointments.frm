VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_Inquiry_SalesAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Appointment Calendar"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Appointments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12270
   Begin XtremeReportControl.ReportControl lvInquiry 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   12030
      _Version        =   655364
      _ExtentX        =   21220
      _ExtentY        =   11668
      _StockProps     =   64
      BorderStyle     =   4
      ShowGroupBox    =   -1  'True
      AllowColumnRemove=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.ComboBox cboYear 
      Height          =   345
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   420
      Width           =   2445
   End
   Begin VB.ComboBox cboMonth 
      Height          =   345
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   60
      Width           =   2475
   End
   Begin VB.PictureBox picAddEdit 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   0
      Width           =   2835
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   675
         Left            =   1950
         MouseIcon       =   "Appointments.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "Appointments.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   645
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   675
         Left            =   1320
         MouseIcon       =   "Appointments.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "Appointments.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   645
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   675
         Left            =   690
         MouseIcon       =   "Appointments.frx":0EBF
         MousePointer    =   99  'Custom
         Picture         =   "Appointments.frx":1011
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   645
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   675
         Left            =   60
         MouseIcon       =   "Appointments.frx":136D
         MousePointer    =   99  'Custom
         Picture         =   "Appointments.frx":14BF
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   645
      End
   End
   Begin VB.ComboBox cboSAE 
      Height          =   345
      Left            =   4380
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8340
      TabIndex        =   10
      Top             =   480
      Width           =   2085
   End
   Begin VB.Label Label2 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   60
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "Filter By SAE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3180
      TabIndex        =   2
      Top             =   180
      Width           =   2085
   End
End
Attribute VB_Name = "frmCRIS_Inquiry_SalesAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAppointment                                      As ADODB.Recordset
Dim LOGACTION                                          As String
Dim ReportTitle                                        As String
Dim WithEvents FormSearch                              As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1


Private Sub cboMonth_Click()
    ShowMonthlyAppointments
End Sub

Private Sub cboYear_Click()
    ShowMonthlyAppointments
End Sub

Private Sub cmdAdd_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND ISDATE(LOGSO)=0")
    'Call FormSearch.SearchForProspects("'P','C','G','I'", "isdate(logso)=0")
    LOGACTION = "SALESAPPOINTMENT"
    FormSearch.Show 1
End Sub

Private Sub cmdDelete_Click()
    If ShowConfirmDelete = False Then
    Exit Sub
    End If

    Screen.MousePointer = 11
    If lvInquiry.SelectedRows.Count = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("Delete from CRIS_SALESAPPOINTMENTS where AppointmentID=" & lvInquiry.SelectedRows.Row(0).Record(12).Value)
    UpdateLog lvInquiry.SelectedRows.Row(0).Record(13).Value
    ShowMonthlyAppointments
    Screen.MousePointer = 0
End Sub

Private Sub cmdEdit_Click()
    Screen.MousePointer = 11
    If lvInquiry.SelectedRows.Count = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    Screen.MousePointer = 0
    Call frmCRIS_Log_SalesAppointment.EditAppointment(lvInquiry.SelectedRows.Row(0).Record(12).Value, lvInquiry.SelectedRows.Row(0).Record(13).Value)
    frmCRIS_Log_SalesAppointment.Show

End Sub

Private Sub cboSAE_Click()
    ShowMonthlyAppointments
End Sub

Private Sub cmdPrint_Click()
    If lvInquiry.Records.Count = 0 Then
        MsgSpeechBox "No Record to Print"
        Exit Sub
        With lvInquiry
            .PaintManager.HorizontalGridStyle = xtpGridNoLines
            .PaintManager.VerticalGridStyle = xtpGridNoLines
        End With
        lvInquiry.PrintOptions.BlackWhiteContrast = 0
        lvInquiry.PrintOptions.BlackWhitePrinting = True
        lvInquiry.PrintOptions.Header.Font.Size = "14"
        lvInquiry.PrintOptions.Header.TextCenter = ReportTitle
        lvInquiry.PrintPreview True
        With lvInquiry
            .PaintManager.HorizontalGridStyle = xtpGridSmallDots
            .PaintManager.VerticalGridStyle = xtpGridSmallDots
        End With
    End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And lvInquiry.Records.Count > 0 Then
        Call frmSMIS_Mis_Filter.ConfigGrid(lvInquiry, 3)
        frmSMIS_Mis_Filter.Show vbModeless
    ElseIf KeyCode = vbKeyF8 And lvInquiry.Records.Count > 0 Then
        lvInquiry.FilterText = vbNullString
        lvInquiry.Populate
        lvInquiry.Columns(4).FooterText = vbNullString
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1

    Set rsAppointment = New ADODB.Recordset
    Set FormSearch = New frmSMIS_Mis_SearchMaster

    InitListView
    InitData
    
End Sub


Sub InitData()

    Dim rsSAE                                          As Recordset
    Set rsSAE = gconDMIS.Execute("select name, id FROM SMIS_vw_Srep")
    While Not rsSAE.EOF
        cboSAE.AddItem rsSAE!Name
        cboSAE.ItemData(cboSAE.NewIndex) = rsSAE!id
        rsSAE.MoveNext
    Wend
    Set rsSAE = Nothing
    fillcbomonth cboMonth
    FillcboYear cboYear
    cboMonth.AddItem "ALL", 0
    Call cboSAE.AddItem("ALL", 0)
End Sub

Sub InitListView()
    ReportControlAddColumnHeader lvInquiry, _
                                 "Date, Time, ProspectName, Make, Color, SAE"
    ResizeColumnHeader lvInquiry, "10,15,10,10,10,10,10"
    lvInquiry.PaintManager.TextFont.Size = 9
    lvInquiry.PaintManager.TextFont.Name = "Arial"
    ReportControlPaintManager lvInquiry
End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    Unload FormSearch

    Select Case LOGACTION
        Case "SALESAPPOINTMENT"
            Screen.MousePointer = 11
            frmCRIS_Log_SalesAppointment.AddSalesAppointment (oCusRs!ProspectID)
            frmCRIS_Log_SalesAppointment.Show
            frmCRIS_Log_SalesAppointment.cmdAdd.Value = True
            Screen.MousePointer = 0
    End Select
End Sub



Sub ShowMonthlyAppointments()
    Dim SQL                                            As String
    Dim oCalEvent                                      As CalendarEvent
    Dim Subject                                        As String
    Dim Importance                                     As String
    SQL = "SELECT "
    SQL = SQL & " Convert(varchar, CSA.StartDateTime,101),   "
    SQL = SQL & " Convert(varchar, CSA.StartDateTime ,108)+ ' - '+ Convert(varchar, CSA.EndDateTime ,108) , "
    SQL = SQL & " CP.AcctName, "
    SQL = SQL & " CSA.ModelDescript, "
    SQL = SQL & " CSA.Color, "
    SQL = SQL & " CSA.SAE,  "
    SQL = SQL & " CP.CUSCDE, StartDateTime, EndDateTime, ExpectedPurchase, Terms , IMPORTANCE, AppointmentID, CP.Prospectid "
    SQL = SQL & " FROM        CRIS_SalesAppointments   CSA "
    SQL = SQL & " INNER JOIN  CRIS_Prospects  "
    SQL = SQL & " CP ON CSA.ProspectID = CP.ProspectID "
    SQL = SQL & " Where Year(StartDateTime) = " & cboYear
    If cboSAE.ListIndex <> 0 And cboSAE.ListIndex <> -1 Then
        SQL = SQL & " AND CSA.SAE='" & cboSAE.Text & "'"
    End If
    If cboMonth.ListIndex <> 0 And cboMonth.ListIndex <> -1 Then
        SQL = SQL & " AND  Month(StartDateTime) = " & What_month(cboMonth)
    End If
    lvInquiry.FilterText = ""
    flex_FillReportView gconDMIS.Execute(SQL), lvInquiry
    If lvInquiry.Rows.Count = 0 Then
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
        cmdEdit.Enabled = True
    End If
End Sub



Sub UpdateLog(ProspectID As Long)

    Dim TSQL                                           As String
    TSQL = " DECLARE @DT DATETIME " & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(StartDateTime) FROM CRIS_SalesAppointments  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGAPPOINTMENT=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGAPPOINTMENT=NULL  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub

Private Sub lvInquiry_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdEdit.Value = True
End Sub

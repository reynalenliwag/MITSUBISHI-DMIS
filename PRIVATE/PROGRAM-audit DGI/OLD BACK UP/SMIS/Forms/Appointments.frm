VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#10.4#0"; "CODEJO~1.OCX"
Begin VB.Form frmSMIS_Inquiry_SalesAppointment 
   BorderStyle     =   3  'Fixed Dialog
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl lvInquiry 
      Height          =   7395
      Left            =   2340
      TabIndex        =   3
      Top             =   0
      Width           =   9930
      _Version        =   655364
      _ExtentX        =   17515
      _ExtentY        =   13044
      _StockProps     =   64
      BorderStyle     =   4
      ShowGroupBox    =   -1  'True
      AllowColumnRemove=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   60
      TabIndex        =   4
      Top             =   2910
      Width           =   2145
      Begin VB.CommandButton Command1 
         Caption         =   "Print "
         Height          =   435
         Left            =   30
         TabIndex        =   9
         ToolTipText     =   "Print"
         Top             =   870
         Width           =   2055
      End
      Begin VB.OptionButton optGridView 
         Caption         =   "Grid View"
         Height          =   405
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Change to Grid View"
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton optCalendar 
         Caption         =   "Calender View"
         Height          =   405
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "View Calendar"
         Top             =   90
         Width           =   2055
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2460
      Width           =   2205
   End
   Begin XtremeCalendarControl.CalendarControl CLEN 
      Height          =   7335
      Left            =   2310
      TabIndex        =   0
      Top             =   30
      Width           =   9930
      _Version        =   655364
      _ExtentX        =   17515
      _ExtentY        =   12938
      _StockProps     =   64
   End
   Begin VB.PictureBox picAddEdit 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   2175
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   675
         Left            =   690
         MouseIcon       =   "Appointments.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Appointments.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Edit Appointment"
         Top             =   90
         Width           =   645
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   675
         Left            =   60
         MouseIcon       =   "Appointments.frx":07B8
         MousePointer    =   99  'Custom
         Picture         =   "Appointments.frx":090A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Add Appointment"
         Top             =   90
         Width           =   645
      End
   End
   Begin XtremeCalendarControl.DatePicker DatePicker1 
      Height          =   2175
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   2235
      _Version        =   655364
      _ExtentX        =   3942
      _ExtentY        =   3836
      _StockProps     =   64
      AutoSize        =   0   'False
      ShowNoneButton  =   0   'False
      ShowWeekNumbers =   -1  'True
      ShowNonMonthDays=   0   'False
      Show3DBorder    =   3
      MaxSelectionCount=   0
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   5520
      ScaleHeight     =   2145
      ScaleWidth      =   3225
      TabIndex        =   13
      Top             =   2280
      Width           =   3255
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
         Height          =   795
         Left            =   2280
         MouseIcon       =   "Appointments.frx":0C1D
         MousePointer    =   99  'Custom
         Picture         =   "Appointments.frx":0D6F
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Close Window"
         Top             =   1200
         Width           =   735
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
         Height          =   795
         Left            =   1560
         MouseIcon       =   "Appointments.frx":11BA
         MousePointer    =   99  'Custom
         Picture         =   "Appointments.frx":130C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print Report"
         Top             =   1200
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   345
         Left            =   840
         TabIndex        =   16
         Top             =   720
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-mm-dd"
         Format          =   51249153
         CurrentDate     =   39963
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   345
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-mm-dd"
         Format          =   51249153
         CurrentDate     =   39934
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   840
         Width           =   765
      End
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
      Left            =   60
      TabIndex        =   8
      Top             =   2220
      Width           =   2085
   End
   Begin VB.Label labMonth 
      Caption         =   "Label1"
      Height          =   465
      Left            =   420
      TabIndex        =   2
      Top             =   6840
      Width           =   1155
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_SalesAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public g_dlgCalendarReminders                                         As New CalendarDialogs
Dim rsAppointment                                                     As ADODB.Recordset
Dim rsSalesOrder                                                      As ADODB.Recordset
Dim rsDelivery                                                        As ADODB.Recordset
Dim LOGACTION                                                         As String
Dim ReportTitle                                                       As String
Dim WithEvents FormSearch                                             As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1
Dim xlApp                                                             As Excel.Application
Dim xlBook                                                            As Excel.Workbook
Dim xlSheet                                                           As Excel.Worksheet
 


Sub InitCalendar()
    With CLEN.Options
        .EnableInPlaceCreateEvent = False
        .EnableInPlaceEditEventSubject_AfterEventResize = False
        .EnableInPlaceEditEventSubject_ByMouseClick = False
        .EnableInPlaceEditEventSubject_ByTab = False
        .EnableInPlaceEditEventSubject_ByF2 = False
        .WorkDayStartTime = "8:00 AM"
        .WorkDayEndTime = "6:00 PM"
        .DayViewTimeScaleShowMinutes = True
        .WorkWeekMask = xtpCalendarDayMo_Fr
    End With
    DatePicker1.AttachToCalendar CLEN
End Sub

Sub InitData()
    optCalendar.Value = True
    Dim RsSAE                                                         As Recordset
    Set RsSAE = gconDMIS.Execute("select name, id FROM SMIS_vw_Srep")
    While Not RsSAE.EOF
        Combo1.AddItem RsSAE!Name
        Combo1.ItemData(Combo1.NewIndex) = RsSAE!ID
        RsSAE.MoveNext
    Wend
    Set RsSAE = Nothing
    Call Combo1.AddItem("ALL", 0)
End Sub

Sub InitListView()
    ReportControlAddColumnHeader lvInquiry, _
                                 "Date, Time, ProspectName, Make, Color, SAE"
    ResizeColumnHeader lvInquiry, "10,15,10,10,10,10,10"
    lvInquiry.PaintManager.TextFont.Size = 9
    lvInquiry.PaintManager.TextFont.Name = "Arial"
    ReportControlPaintManager lvInquiry
End Sub

Sub ShowMonthlyAppointments(MonthDate As Date)
    Dim SQL                                                           As String
    Dim oCalEvent                                                     As CalendarEvent
    Dim Subject                                                       As String
    Dim Importance                                                    As String
    Dim monthx
    Dim I                                                             As Integer
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
    SQL = SQL & " Where Month(StartDateTime) = " & Month(MonthDate)
    SQL = SQL & " And Year(StartDateTime) = " & Year(MonthDate)

    If Combo1.ListIndex <> 0 And Combo1.ListIndex <> -1 Then
        SQL = SQL & " AND CSA.SAE='" & Combo1.Text & "'"
    End If
    If optGridView.Value = True Then
        lvInquiry.FilterText = ""
        flex_FillReportView gconDMIS.Execute(SQL), lvInquiry
    Else
        If Month(MonthDate) <> labMonth Then
            labMonth = monthx
            CLEN.DataProvider.RemoveAllEvents
            Set rsAppointment = gconDMIS.Execute(SQL)
            While Not rsAppointment.EOF
                I = I + 1
                Set oCalEvent = CLEN.DataProvider.CreateEventEx(I)
                oCalEvent.StartTime = rsAppointment!StartDateTime
                oCalEvent.EndTime = rsAppointment!EndDateTime

                If Null2String(rsAppointment!ModelDescript) <> "" Then
                    Subject = " MODEL:" & rsAppointment!ModelDescript
                End If
                If Null2String(rsAppointment!Color) <> "" Then
                    Subject = Subject & " COLOR:" & rsAppointment!Color
                End If
                If Null2String(rsAppointment!ModelDescript) <> "" Then
                    Subject = Subject & " EXPECTED BUY:" & rsAppointment!ExpectedPurchase
                End If
                If Null2String(rsAppointment!Terms) <> "" Then
                    Subject = Subject & " TERMS:" & rsAppointment!Terms
                End If

                oCalEvent.Subject = Subject
                oCalEvent.Location = Null2String(rsAppointment!SAE)
                Importance = Null2String(rsAppointment!Importance)
                If Importance = "" Or Importance = "N" Then
                    oCalEvent.Importance = xtpCalendarImportanceHigh
                ElseIf Importance = "H" Then
                    oCalEvent.Importance = xtpCalendarImportanceHigh
                Else
                    oCalEvent.Importance = xtpCalendarImportanceLow
                End If
                CLEN.DataProvider.AddEvent oCalEvent
                rsAppointment.MoveNext
            Wend
            CLEN.Populate
            CLEN.DayView.ScrollToWorkDayBegin
        End If
    End If
End Sub

Sub ViewType()
    If optCalendar.Value = True Then
        CLEN.Visible = True
        lvInquiry.Visible = False
        picAddEdit.Visible = False
        ReportTitle = "SALES APPOINTMENT CALENDAR "
    Else
        CLEN.Visible = False
        ReportTitle = "SALES APPOINTMENT SCHEDULE"
        lvInquiry.Visible = True
        picAddEdit.Visible = True
    End If
    ShowMonthlyAppointments DatePicker1.FirstVisibleDay
End Sub

Private Sub CLEN_ViewChanged()
    If CLEN.ViewType = xtpCalendarMonthView Then
        CLEN.ViewType = xtpCalendarDayView
    End If
End Sub

Private Sub cmdAdd_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND ISDATE(LOGSO)=0")
    'Call FormSearch.SearchForProspects("'P','C','G','I'", "isdate(logso)=0")
    LOGACTION = "SALESAPPOINTMENT"
    FormSearch.Show 1
End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdCancel_Click()
pic1.Visible = False
pic1.ZOrder 1
End Sub

Private Sub cmdEdit_Click()
    Screen.MousePointer = 11
    If lvInquiry.SelectedRows.Count = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    Screen.MousePointer = 0

    Call frmSMIS_Log_SalesAppointment.EditAppointment(lvInquiry.SelectedRows.Row(0).Record(12).Value, lvInquiry.SelectedRows.Row(0).Record(13).Value)
    frmSMIS_Log_SalesAppointment.Show
    frmSMIS_Log_SalesAppointment.cmdEdit.Value = True
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11

    If Len(Dir(SMIS_REPORT_PATH & "SAS.xlt")) <= 0 Then
        If EXTRACT_FILES(106, "SAS.xlt") = False Then
            MsgBox "Excel file cannot be located. Please add excel Template file in DMIS Folder.", vbInformation, "SMIS"
            Exit Sub
        End If
    End If

    
    
    Dim RS                                          As ADODB.Recordset
    Dim vdate                                       As String
    Dim vtime                                       As String
    Dim vpname                                      As String
    Dim vmodel                                      As String
    Dim vcolor                                      As String
    Dim vsae                                        As String
    Dim vnote                                       As String
    Dim vcounter                                    As Integer
    Dim SQL As String
    
    
    Dim FDate                                          As Date
    Dim TDate                                          As Date

    FDate = CDate(dtpFromDate.Value)
    TDate = CDate(dtpToDate.Value)
    
    
    
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "SAS.xlt")
    Set xlSheet = xlBook.Worksheets(1)
    
    
    Set RS = New ADODB.Recordset
    
    SQL = "SELECT S.*, P.ACCTNAME  FROM CRIS_SalesAppointments S INNER JOIN CRIS_PROSPECTS P ON S.PROSPECTID = P.PROSPECTID where startdatetime  > = '" & FDate & "' and startdatetime <= '" & TDate & "' order by startdatetime asc"
    
    Set RS = gconDMIS.Execute(SQL)
    
    xlSheet.Cells(7, "B") = FDate
    xlSheet.Cells(8, "B") = TDate
    
    
    If Not RS.BOF And Not RS.EOF Then

        RS.MoveFirst: vcounter = 0
        
        Do While Not RS.EOF
        
        vdate = Null2String(RS!StartDateTime)
        vtime = Null2String(RS!StartDateTime)
        vpname = Null2String(RS!AcctName)
        vmodel = Null2String(RS!ModelDescript)
        vcolor = Null2String(RS!Color)
        vsae = Null2String(RS!SAE)
        vnote = Null2String(RS!Notes)
        

        xlSheet.Cells(12 + vcounter, "A") = vdate
        xlSheet.Cells(12 + vcounter, "B") = vtime
        xlSheet.Cells(12 + vcounter, "C") = vpname
        xlSheet.Cells(12 + vcounter, "E") = vmodel
        xlSheet.Cells(12 + vcounter, "F") = vcolor
        xlSheet.Cells(12 + vcounter, "G") = vsae
        xlSheet.Cells(12 + vcounter, "H") = vnote
        
        
        vcounter = vcounter + 1
        RS.MoveNext
        
        Loop
    
    End If

    xlApp.Visible = True
    DoEvents
    Set xlApp = Nothing
    
Screen.MousePointer = 0


End Sub

Private Sub Combo1_Click()
    ShowMonthlyAppointments DatePicker1.Selection(0).DateBegin
End Sub

Private Sub Command1_Click()
    
If COMPANY_CODE = "HGC" Then

    pic1.Visible = True
    pic1.ZOrder 0

Else
    
    
    If CLEN.Visible = True Then
        CLEN.PrintPreview True
    Else
        If lvInquiry.Records.Count = 0 Then
            MsgSpeechBox "No Record to Print"
            Exit Sub
        End If
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

    NEW_LogAudit "V", "f", "", "", "", "Sales Appointment Calendar by: " & Combo1, "", ""

End If
End Sub

Private Sub DatePicker1_MonthChanged()
    ShowMonthlyAppointments DatePicker1.FirstVisibleDay
End Sub

Private Sub DatePicker1_SelectionChanged()
    If optGridView.Value = True Then
        lvInquiry.FilterText = Format(DatePicker1.Selection.Blocks(0).DateBegin, "mm/dd/yyyy")
        lvInquiry.Populate
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Sales Appointment Calendar)"
            Call frmALL_AuditInquiry.DisplayHistory("", "Sales Appointment Calendar", "PRINTING")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set rsAppointment = New ADODB.Recordset
    Set rsDelivery = New ADODB.Recordset
    Set rsSalesOrder = New ADODB.Recordset
    Set FormSearch = New frmSMIS_Mis_SearchMaster
    InitCalendar
    InitListView
    InitData
End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    Unload FormSearch

    Select Case LOGACTION
        Case "SALESAPPOINTMENT"
            frmSMIS_Log_SalesAppointment.AddSalesAppointment (oCusRs!PROSPECTID)
            frmSMIS_Log_SalesAppointment.Show
            frmSMIS_Log_SalesAppointment.cmdAdd.Value = True
    End Select
End Sub

Private Sub lvInquiry_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    cmdEdit.Value = True
End Sub

Private Sub optCalendar_Click()
    ViewType
End Sub

Private Sub optGridView_Click()
    ViewType
End Sub


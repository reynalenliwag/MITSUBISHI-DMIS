VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_ViewAllApointments 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl lvAppointments 
      Height          =   6285
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   9915
      _Version        =   655364
      _ExtentX        =   17489
      _ExtentY        =   11086
      _StockProps     =   64
      BorderStyle     =   2
      ShowFooter      =   -1  'True
   End
   Begin VB.PictureBox Picture 
      Appearance      =   0  'Flat
      BackColor       =   &H00853036&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   2070
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   3
      Top             =   480
      Width           =   315
   End
   Begin VB.PictureBox Picture 
      Appearance      =   0  'Flat
      BackColor       =   &H004A5A8A&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   60
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   2
      Top             =   480
      Width           =   315
   End
   Begin VB.Label lbl 
      Caption         =   "Upcoming Appointment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2490
      TabIndex        =   5
      Top             =   495
      Width           =   2025
   End
   Begin VB.Label lbl 
      Caption         =   "Old Appointments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   450
      TabIndex        =   4
      Top             =   495
      Width           =   1515
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
Attribute VB_Name = "frmCRIS_ViewAllApointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SQL                                      As String

Friend Sub SQLString(s As String, AppointmentType As Integer)
    SQL = "Select  " & _
          "StartTime + '-' + EndTime , " & _
          "StartDate + ' - ' + EndDate , " & _
          "EndDate , " & _
          "Model , " & _
          "ProfileName , " & _
          "SAE " & _
        " from  CRIS_vW_Appointments Where  AppointmentType=@APTYPE and " & _
          "cast(ProfileID as varchar) + ProfileType in  " & _
          "(select cast(ProfileID as varchar) + ProfileType from CRIS_CalendarEvents where AppointmentID in (@QUERY)) order by 3"
    SQL = Replace(SQL, "@QUERY", s)

    If AppointmentType = 1 Then
        SQL = Replace(SQL, "@APTYPE", 1)
        ShortcutCaption.caption = "Test Drive Appointment"
    Else
        SQL = Replace(SQL, "@APTYPE", 2)
        ShortcutCaption.caption = "Sales Appointment History/Summary"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Form_Load()
    With lvAppointments                                       '''''''''''UI
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots  'xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(255, 245, 255)
        .PaintManager.ShadeSortColor = RGB(229, 229, 229)
        .PaintManager.HighlightBackColor = RGB(0, 0, 0)
        .PaintManager.VerticalGridStyle = xtpGridNoLines      ' xtpGridSmallDots
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
    End With



    With lvAppointments
        .Columns.Add 0, "Time", 80, True
        .Columns.Add 1, "Date Time", 100, True
        .Columns.Add 2, "End Date", 0, True
        .Columns.Add 3, "Model", 100, True
        .Columns.Add 4, "Profile Name", 100, True
        .Columns.Add 5, "SAE", 100, True
        .Columns(2).Visible = False
        .Columns(0).FooterText = "F3: Add Filter"
        .Columns(1).FooterText = "F8: Remove Filter"
    End With



    Dim tmprs                                As ADODB.Recordset
    Dim FLD                                  As ADODB.Field
    Dim REC                                  As ReportRecord

    Set tmprs = gconDMIS.Execute(SQL)
    While Not tmprs.EOF
        Set REC = lvAppointments.Records.Add
        For Each FLD In tmprs.Fields
            REC.AddItem FLD.Value

        Next
        tmprs.MoveNext
    Wend

    lvAppointments.Populate
End Sub

Private Sub lvAppointments_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record Is Nothing Then Exit Sub
    If CDate(Row.Record(2).Value) <= CDate(FormatDateTime(Now, vbShortDate)) Then
        'Metrics.Font.Strikethrough =&H004A5A8A& True
        Metrics.ForeColor = RGB(138, 90, 74)
    Else
        Metrics.ForeColor = RGB(54, 48, 133)

    End If
End Sub


Private Sub lvAppointments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And lvAppointments.Records.Count > 0 Then
        Call frmCRIS_Filter.ConfigGrid(lvAppointments, 3)
        
        frmCRIS_Filter.Show vbModeless
    ElseIf KeyCode = vbKeyF8 And lvAppointments.Records.Count > 0 Then
        lvAppointments.FilterText = vbNullString
        lvAppointments.Populate
        lvAppointments.Columns(4).FooterText = vbNullString
    End If
End Sub

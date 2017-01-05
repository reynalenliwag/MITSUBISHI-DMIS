VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmRAMS_Inquiry_Audit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AUDIT INQUIRY"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Inquiry_Audit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl ReportControl1 
      Height          =   4800
      Left            =   1740
      TabIndex        =   22
      Top             =   1230
      Width           =   7995
      _Version        =   655364
      _ExtentX        =   14102
      _ExtentY        =   8467
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnRemove=   0   'False
      AllowColumnReorder=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   30
      ScaleHeight     =   4140
      ScaleWidth      =   1710
      TabIndex        =   10
      Top             =   1200
      Width           =   1710
      Begin VB.CheckBox ChkCheck 
         Caption         =   "GENERATED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   90
         TabIndex        =   21
         Tag             =   "G"
         Top             =   2265
         Width           =   1680
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "INQUIRY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   90
         TabIndex        =   20
         Tag             =   "I"
         Top             =   1425
         Width           =   1680
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "BATCH POSTED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   90
         TabIndex        =   19
         Tag             =   "O"
         Top             =   600
         Width           =   1680
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "CANCELLED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Tag             =   "C"
         Top             =   1980
         Width           =   1365
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "POSTED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Tag             =   "P"
         Top             =   315
         Width           =   1005
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "UN-POSTED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Tag             =   "U"
         Top             =   1155
         Width           =   1410
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "VIEWED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   16
         Tag             =   "V"
         Top             =   870
         Width           =   1230
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "ADDED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   17
         Tag             =   "A"
         Top             =   1710
         Width           =   1230
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "UPDATED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   90
         TabIndex        =   18
         Tag             =   "E"
         Top             =   2535
         Width           =   1230
      End
      Begin VB.CheckBox ChkCheck 
         Caption         =   "DELETED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   15
         Tag             =   "X"
         Top             =   2820
         Width           =   1230
      End
      Begin VB.Label Label6 
         Caption         =   "Select Modules"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   0
         Width           =   1950
      End
   End
   Begin Crystal.CrystalReport rptInternalReminder 
      Left            =   3300
      Top             =   6135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1770
      TabIndex        =   8
      Top             =   780
      Width           =   7845
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   795
      Left            =   8880
      MouseIcon       =   "Inquiry_Audit.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Inquiry_Audit.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6060
      Width           =   765
   End
   Begin VB.CheckBox Check7 
      Caption         =   "IN  DATE RANGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   5
      Top             =   427
      Width           =   1650
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5490
      TabIndex        =   6
      Top             =   360
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483635
      CalendarTitleForeColor=   16777215
      Format          =   48889857
      CurrentDate     =   39218
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   45
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   375
      Width           =   2940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "::"
      Height          =   330
      Left            =   3015
      TabIndex        =   4
      Top             =   382
      Width           =   285
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7650
      TabIndex        =   7
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48889857
      CurrentDate     =   39218
   End
   Begin VB.CommandButton cmdInquire 
      Caption         =   "&Inquiry"
      Height          =   795
      Left            =   8130
      MouseIcon       =   "Inquiry_Audit.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "Inquiry_Audit.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6060
      Width           =   765
   End
   Begin VB.Label Label5 
      Caption         =   "Filter View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   9
      Top             =   870
      Width           =   1950
   End
   Begin VB.Label Label4 
      Caption         =   "Total Result(s)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   23
      Top             =   6270
      Width           =   2400
   End
   Begin VB.Label Label2 
      Caption         =   "For :(Date)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5535
      TabIndex        =   1
      Top             =   120
      Width           =   2010
   End
   Begin VB.Label Label3 
      Caption         =   "TO: (DATE)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7620
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Select User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   1950
   End
End
Attribute VB_Name = "frmRAMS_Inquiry_Audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INQUIRY_TAG                                                       As String

Sub InitMemVars()
    Dim RS                                                            As ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT USERID, upper(USERNAME) as USERNAME FROM ALL_RAMS_USERS")
    While Not RS.EOF
        With Combo1
            .AddItem Null2String(RS!UserName)
            .ItemData(.NewIndex) = RS!USERID
        End With
        RS.MoveNext
    Wend
    If Combo1.ListCount > 0 Then
        Combo1.ListIndex = 0
    End If
    With ReportControl1
        .Columns.Add 0, "Date", 80, True
        .Columns.Add 1, "Time", 80, True

        .Columns.Add 2, "Description", 250, True
        .Columns.Add 3, "User Action", 100, True
        .Columns.Add 4, "Tracking ID", 220, True
    End With
    cmdInquire.Enabled = False
    '    cmdPrint.Enabled = False
End Sub

Private Sub Check7_Click()
    If Check7.Value = 1 Then
        Label3.Visible = True
        DTPicker2.Visible = True
        Label2.Caption = "FROM:(DATE)"


        DTPicker2.MinDate = DTPicker1.Value
        DTPicker2.Value = DateAdd("d", 1, DTPicker1.Value)

    Else
        Label3.Visible = False
        DTPicker2.Visible = False
        Label2.Caption = "FOR:(DATE)"

    End If
End Sub

Private Sub ChkCheck_Click(Index As Integer)


    Dim MyTag                                                         As String
    INQUIRY_TAG = vbNullString
    For i = 0 To ChkCheck.Count - 1
        If ChkCheck(i).Value = 1 Then
            INQUIRY_TAG = INQUIRY_TAG & "'" & ChkCheck(i).Tag & "'" & ","
        End If
    Next

    If INQUIRY_TAG <> vbNullString Then
        MyTag = Left(INQUIRY_TAG, Len(INQUIRY_TAG) - 1)
    End If


    '    ReportControl1.GroupsOrder.DeleteAll
    '   ReportControl1.Columns(3).Visible = True

    If Len(MyTag) > 0 Then
        cmdInquire.Enabled = True
        INQUIRY_TAG = "(" & MyTag & ")"
        '      If Len(MyTag) > 3 Then
        '         ReportControl1.GroupsOrder.Add ReportControl1.Columns(3)
        '        ReportControl1.Columns(3).Visible = False
        '   End If

    Else

        cmdInquire.Enabled = False
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200715:24
Private Sub cmdInquire_Click()
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim lngCount                                                      As Long
    Dim SQL                                                           As String
    Dim DATETAG                                                       As String
    Dim REC                                                           As XtremeReportControl.ReportRecord


    On Error GoTo Errorcode:

    If DTPicker2.Visible = True Then
        DATETAG = " ACTION_DATE between " & N2Date2Null(DTPicker1.Value) & " AND " & N2Date2Null(DTPicker2.Value)
    Else
        DATETAG = " convert(varchar, A.ACTION_DATE,101) = '" & Format(DTPicker1.Value, "mm/dd/yyyy") & "'"
    End If

    SQL = " SELECT " & vbCrLf
    SQL = SQL & "convert(varchar, A.ACTION_DATE,101)," & vbCrLf
    SQL = SQL & "convert(varchar, A.ACTION_DATE,08)," & vbCrLf
    SQL = SQL & " A.MODULE_NAME ," & vbCrLf
    SQL = SQL & " Case a.USER_ACTION " & vbCrLf
    SQL = SQL & " WHEN 'A' THEN 'ADDED'" & vbCrLf
    SQL = SQL & " WHEN 'E' THEN 'EDITED'" & vbCrLf
    SQL = SQL & " WHEN 'P' THEN 'POSTED'" & vbCrLf
    SQL = SQL & " WHEN 'U' THEN 'UNPOSTED'" & vbCrLf
    SQL = SQL & " WHEN 'C' THEN 'CANCELLED'" & vbCrLf
    SQL = SQL & " WHEN 'X' THEN 'DELETED'" & vbCrLf
    SQL = SQL & " WHEN 'V' THEN 'VIEWED'" & vbCrLf
    SQL = SQL & " WHEN 'I' THEN 'INQUIRED'" & vbCrLf
    SQL = SQL & " WHEN 'R' THEN 'PROCESSED'"
    SQL = SQL & " WHEN 'G' THEN 'GENERATED'"
    SQL = SQL & " WHEN 'O' THEN 'BATCH POSTING'"
    SQL = SQL & " END as User_Action ," & vbCrLf
    SQL = SQL & "A.TRACKING_MEMO," & vbCrLf
    SQL = SQL & "C.USERNAME " & vbCrLf
    SQL = SQL & "FROM DMIS_AUDIT.DBO.DMIS_AUDIT A" & vbCrLf
    SQL = SQL & "INNER JOIN DMIS.DBO.ALL_Rams_Users C ON" & vbCrLf
    SQL = SQL & "a.[USER_ID] = C.[UserID] WHERE USERID=" & Combo1.ItemData(Combo1.ListIndex)


    If Len(INQUIRY_TAG) > 0 Then
        SQL = SQL & "  And USER_ACTION in  " & INQUIRY_TAG
    End If

    If Len(DATETAG) > 0 Then
        SQL = SQL & " AND " & DATETAG
    End If



    Set TEMPRS = gconAudit.Execute(SQL)


    ReportControl1.Records.DeleteAll
    While Not TEMPRS.EOF
        Set REC = ReportControl1.Records.Add

        REC.AddItem TEMPRS.Fields(0).Value
        REC.AddItem Format(TEMPRS.Fields(1).Value, "hh:mm:ss:AM/PM")
        REC.AddItem TEMPRS.Fields(2).Value
        REC.AddItem TEMPRS.Fields(3).Value
        REC.AddItem TEMPRS.Fields(4).Value
        REC.AddItem TEMPRS.Fields(5).Value

        TEMPRS.MoveNext
    Wend
    ReportControl1.Populate

    Set REC = Nothing
    Set TEMPRS = Nothing


    lngCount = ReportControl1.Records.Count
    Label4 = "Total Result(s)" & lngCount
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrint_Click()
    '    Dim crystalFilter                  As String
    '    Dim FILTER  As String
    '    Dim i                              As Integer
    '
    '    For i = 0 To ChkCheck.Count - 1
    '        If ChkCheck(i).Value = 1 Then
    '            crystalFilter = crystalFilter & " {AD.ACTION_TYPE} =(" & N2Str2Null(ChkCheck(i).Tag) & ") OR"
    '        End If
    '    Next
    '
    '    crystalFilter = Left(crystalFilter, Len(crystalFilter) - 2)
    '    crystalFilter = crystalFilter & " AND  {AD.USERID} =(" & LOGID & ")"
    '
    '
    '
    '    rptInternalReminder.Formulas(0) = "CompanyName = '" & Company_name & "'"
    '    rptInternalReminder.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
    '
    '
    '
    '    FILTER = " {U.USERID}=" & UID
    '
    '    If chkDataEntry.Value = 1 Then
    '        PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "AccessReportDataEntry.rpt", FILTER, DMIS_REPORT_Connection, 1
    '    End If
    '
    '    If chkTransaction.Value = 1 Then
    '        PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "AccessReportTransaction.rpt", FILTER, DMIS_REPORT_Connection, 1
    '    End If
    '    If chkReport.Value = 1 Then
    '
    '        PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "AccessReportReport.rpt", FILTER, DMIS_REPORT_Connection, 1
    '    End If
    '
    '    If chkProcessing.Value = 1 Then
    '        PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "AccessReportProcessing.rpt", FILTER, DMIS_REPORT_Connection, 1
    '    End If
    '
    '    If chkInquiry.Value = 1 Then
    '        PrintSQLReport rptInternalReminder, CRIS_REPORT_PATH & "AccessReportInquiry.rpt", FILTER, DMIS_REPORT_Connection, 1
    '    End If
    '

End Sub

Private Sub Combo1_LostFocus()
    Combo1.Enabled = False

End Sub

Private Sub Command1_Click()
    If Combo1.ListCount > 0 Then
        Combo1.Enabled = True
    End If
End Sub

Private Sub DTPicker1_Change()
    DTPicker2.MinDate = DTPicker1.Value
    DTPicker2.Value = DateAdd("d", 1, DTPicker1.Value)
    'DTPicker2.Value =
End Sub

Private Sub Form_Load()

    InitMemVars
    DTPicker1.Value = firstDay(Date)
    'DTPicker2.Value =
End Sub

Private Sub Text1_Change()
    ReportControl1.FilterText = Trim(Text1.Text)
    ReportControl1.Populate
End Sub


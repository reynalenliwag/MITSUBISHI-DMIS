VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Begin VB.Form frmInquiry_Audit 
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
   Icon            =   "INQUIRY_AUDIT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl lstAudit_Inquiry 
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
      Height          =   3030
      Left            =   30
      ScaleHeight     =   3030
      ScaleWidth      =   1710
      TabIndex        =   10
      Top             =   1200
      Width           =   1710
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   90
         TabIndex        =   21
         Tag             =   "G"
         Top             =   2265
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   90
         TabIndex        =   20
         Tag             =   "I"
         Top             =   1425
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   90
         TabIndex        =   19
         Tag             =   "O"
         Top             =   600
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Tag             =   "C"
         Top             =   1980
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Tag             =   "P"
         Top             =   315
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Tag             =   "U"
         Top             =   1155
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   16
         Tag             =   "V"
         Top             =   870
         Value           =   1  'Checked
         Width           =   1230
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   17
         Tag             =   "A"
         Top             =   1710
         Value           =   1  'Checked
         Width           =   1230
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   90
         TabIndex        =   18
         Tag             =   "E"
         Top             =   2535
         Value           =   1  'Checked
         Width           =   1230
      End
      Begin VB.CheckBox ChkCheck 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   15
         Tag             =   "X"
         Top             =   2820
         Value           =   1  'Checked
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
   Begin VB.TextBox txtSearchAudit 
      Height          =   405
      Left            =   1770
      TabIndex        =   8
      Top             =   780
      Width           =   7905
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   795
      Left            =   8880
      MouseIcon       =   "INQUIRY_AUDIT.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "INQUIRY_AUDIT.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6060
      Width           =   765
   End
   Begin VB.CheckBox Check7 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3840
      TabIndex        =   5
      Top             =   427
      Width           =   1650
   End
   Begin MSComCtl2.DTPicker dtFrom 
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
      Format          =   113508353
      CurrentDate     =   39218
   End
   Begin VB.ComboBox cboUsers 
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
   Begin VB.CommandButton cmdChangeUser 
      Caption         =   "::"
      Height          =   330
      Left            =   3015
      TabIndex        =   4
      Top             =   382
      Width           =   285
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   7650
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
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
      Format          =   113508353
      CurrentDate     =   39218
   End
   Begin VB.CommandButton cmdInquire 
      Caption         =   "&Inquiry"
      Height          =   795
      Left            =   8130
      MouseIcon       =   "INQUIRY_AUDIT.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "INQUIRY_AUDIT.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6060
      Width           =   765
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
      Left            =   7380
      MouseIcon       =   "INQUIRY_AUDIT.frx":0D40
      MousePointer    =   99  'Custom
      Picture         =   "INQUIRY_AUDIT.frx":0E92
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6060
      Width           =   765
   End
   Begin VB.Label Label7 
      BackColor       =   &H00004000&
      Caption         =   "PRESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   33
      Top             =   4350
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "F5 : INQUIRY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   32
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "F6 : PRINT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   31
      Top             =   5655
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "ESC : EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   3
      Left            =   60
      TabIndex        =   30
      Top             =   4650
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "F7: CHANGE DATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   6
      Left            =   60
      TabIndex        =   29
      Top             =   5910
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "F2: CHANGE USER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   5
      Left            =   60
      TabIndex        =   28
      Top             =   4905
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "F3: SEARCH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   4
      Left            =   60
      TabIndex        =   27
      Top             =   5160
      Width           =   1455
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
Attribute VB_Name = "frmInquiry_Audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INQUIRY_TAG                                             As String

Sub initMemvars()
    Dim RS                                                  As ADODB.Recordset
    On Error GoTo ErrorCode                           'AXP063121:27
    If CHANGE_USER = True Then
        Set RS = gconDMIS.Execute("SELECT USERID, upper(USER_NAME) as USERNAME FROM ALL_RAMS_USERS order by user_name")
    Else
        Set RS = gconDMIS.Execute("SELECT USERID, upper(USERNAME) as USERNAME FROM ALL_RAMS_USERS order by username")
    End If
    While Not RS.EOF
        With cboUsers
            .AddItem Null2String(RS!UserName)
            .ItemData(.NewIndex) = RS!USERID
        End With
        RS.MoveNext
    Wend
    If cboUsers.ListCount > 0 Then
        cboUsers.ListIndex = 0
    End If
    With lstAudit_Inquiry
        .Columns.Add 0, "Date", 80, True
        .Columns.Add 1, "Time", 80, True
        .Columns.Add 2, "Description", 250, True
        .Columns.Add 3, "User Action", 100, True
        .Columns.Add 4, "Tracking ID", 220, True
        .GroupsOrder.Add .Columns(3)
        .Columns(3).Visible = False
    End With
    '    cmdPrint.Enabled = False:cmdInquire.Enabled = False


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Check7_Click()
    If Check7.Value = 1 Then
        Label3.Visible = True
        dtTo.Visible = True
        Label2.Caption = "FROM:(DATE)"


        dtTo.MinDate = dtFrom.Value
        dtTo.Value = DateAdd("d", 1, dtFrom.Value)

    Else
        Label3.Visible = False
        dtTo.Visible = False
        Label2.Caption = "FOR:(DATE)"

    End If
End Sub

Private Sub ChkCheck_Click(Index As Integer)


    Dim MyTag                                               As String
    INQUIRY_TAG = vbNullString
    For i = 0 To ChkCheck.Count - 1
        If ChkCheck(i).Value = 1 Then
            INQUIRY_TAG = INQUIRY_TAG & "'" & ChkCheck(i).Tag & "'" & ","
        End If
    Next

    If INQUIRY_TAG <> vbNullString Then
        MyTag = Left(INQUIRY_TAG, Len(INQUIRY_TAG) - 1)
    End If


    '    lstAudit_inquiry.GroupsOrder.DeleteAll
    '   lstAudit_inquiry.Columns(3).Visible = True

    If Len(MyTag) > 0 Then
        cmdInquire.Enabled = True
        INQUIRY_TAG = "(" & MyTag & ")"
        '      If Len(MyTag) > 3 Then
        '         lstAudit_inquiry.GroupsOrder.Add lstAudit_inquiry.Columns(3)
        '        lstAudit_inquiry.Columns(3).Visible = False
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
    Dim temprs                                              As ADODB.Recordset
    Dim lngCount                                            As Long
    Dim SQL                                                 As String
    Dim DATETAG                                             As String
    Dim REC                                                 As XtremeReportControl.ReportRecord


    On Error GoTo ErrorCode:

    If dtTo.Visible = True Then
        DATETAG = " ACTION_DATE between " & N2Date2Null(DateValue(DateAdd("d", -1, dtFrom.Value))) & " AND " & N2Date2Null(DateValue(DateAdd("d", 1, dtTo.Value)))
    Else
        DATETAG = " convert(varchar, A.ACTION_DATE,101) = '" & Format(dtFrom.Value, "mm/dd/yyyy") & "'"
    End If
    If CHANGE_USER = True Then
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
        SQL = SQL & "C.USER_NAME " & vbCrLf
        SQL = SQL & "FROM DMIS_AUDIT.DBO.DMIS_AUDIT A" & vbCrLf
        SQL = SQL & "INNER JOIN DMIS.DBO.ALL_Rams_Users C ON" & vbCrLf
        SQL = SQL & "a.[USER_ID] = C.[UserID] WHERE USERID=" & cboUsers.ItemData(cboUsers.ListIndex)

    Else
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
        SQL = SQL & "a.[USER_ID] = C.[UserID] WHERE USERID=" & cboUsers.ItemData(cboUsers.ListIndex)
    End If

    If Len(INQUIRY_TAG) > 0 Then
        SQL = SQL & "  And USER_ACTION in  " & INQUIRY_TAG
    End If

    If Len(DATETAG) > 0 Then
        SQL = SQL & " AND " & DATETAG
    End If


    Set temprs = gconAudit.Execute(SQL)


    lstAudit_Inquiry.Records.DeleteAll
    While Not temprs.EOF
        Set REC = lstAudit_Inquiry.Records.Add

        REC.AddItem temprs.Fields(0).Value
        REC.AddItem Format(temprs.Fields(1).Value, "hh:mm:ss:AM/PM")
        REC.AddItem temprs.Fields(2).Value
        REC.AddItem temprs.Fields(3).Value
        REC.AddItem temprs.Fields(4).Value
        REC.AddItem temprs.Fields(5).Value

        temprs.MoveNext
    Wend
    lstAudit_Inquiry.Populate

    Set REC = Nothing
    Set temprs = Nothing


    lngCount = lstAudit_Inquiry.Records.Count
    Label4 = "Total Result(s)" & lngCount
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdPrint_Click()
    If lstAudit_Inquiry.Records.Count <= 0 Then Exit Sub
    lstAudit_Inquiry.PrintOptions.Header.TextCenter = "AUDIT PRINT FOR " & cboUsers.Text
    lstAudit_Inquiry.PrintPreview True

End Sub

Private Sub cboUsers_LostFocus()
    cboUsers.Enabled = False

End Sub

Private Sub cmdChangeUser_Click()
    If cboUsers.ListCount > 0 Then
        cboUsers.Enabled = True
        cboUsers.SetFocus
        SendKeys ("{F4}")
    End If
End Sub

Private Sub dtFrom_Change()
'dtTo.MinDate = dtFrom.Value
'dtTo.Value = DateAdd("d", 1, dtFrom.Value)
'dtTo.Value =
End Sub

Private Sub dtFrom_GotFocus()
    SendKeys ("{F4}")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    ElseIf KeyCode = vbKeyF2 Then
        cmdChangeUser.Value = True
    ElseIf KeyCode = vbKeyF3 Then
        txtSearchAudit.SetFocus
    ElseIf KeyCode = vbKeyF5 And cmdInquire.Enabled = True Then
        cmdInquire.Value = True
        txtSearchAudit.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        cmdPrint.Value = True
        txtSearchAudit.SetFocus
    ElseIf KeyCode = vbKeyF7 Then
        dtFrom.SetFocus
    End If


End Sub

Private Sub Form_Load()
    initMemvars
    dtFrom.Value = Now
    dtTo.Value = Now
End Sub

Private Sub txtSearchAudit_Change()
    lstAudit_Inquiry.FilterText = Trim(txtSearchAudit.Text)
    lstAudit_Inquiry.Populate
End Sub

Private Sub txtSearchAudit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstAudit_Inquiry.Records.Count > 0 Then
            lstAudit_Inquiry.SelectedRows(0).Selected = True
            lstAudit_Inquiry.SetFocus
        End If
    End If

End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Begin VB.Form frmSMIS_Log_Reminder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminders"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogReminders.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   6660
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   6285
      Left            =   0
      ScaleHeight     =   6285
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   390
      Width           =   6795
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   630
         ScaleHeight     =   900
         ScaleWidth      =   6090
         TabIndex        =   23
         Top             =   5100
         Width           =   6090
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   5250
            MouseIcon       =   "LogReminders.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   4560
            MouseIcon       =   "LogReminders.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Delete Selected Reminder"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3870
            MouseIcon       =   "LogReminders.frx":11FF
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":1351
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Edit Selected Reminder"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   3180
            MouseIcon       =   "LogReminders.frx":16AD
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":17FF
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Add Reminder"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   2490
            MouseIcon       =   "LogReminders.frx":1B12
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":1C64
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1800
            MouseIcon       =   "LogReminders.frx":1F5E
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":20B0
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   1110
            MouseIcon       =   "LogReminders.frx":2408
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":255A
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picDataEntry 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5115
         Left            =   0
         ScaleHeight     =   5115
         ScaleWidth      =   6765
         TabIndex        =   2
         Top             =   0
         Width           =   6765
         Begin VB.PictureBox picRemind_Bottom 
            BorderStyle     =   0  'None
            Height          =   1485
            Left            =   60
            ScaleHeight     =   1485
            ScaleWidth      =   6615
            TabIndex        =   18
            Top             =   3720
            Width           =   6615
            Begin VB.TextBox txtFollowUpNotes 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1020
               Left            =   0
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   360
               Width           =   6525
            End
            Begin VB.ComboBox cboStatus 
               Height          =   330
               Left            =   4860
               TabIndex        =   21
               Text            =   "Combo1"
               Top             =   0
               Width           =   1635
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Status:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   4200
               TabIndex        =   20
               Top             =   60
               Width           =   600
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Follow Up Notes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   60
               TabIndex        =   19
               Top             =   60
               Width           =   1350
            End
         End
         Begin VB.PictureBox picRemind_Top 
            BorderStyle     =   0  'None
            Height          =   3645
            Left            =   60
            ScaleHeight     =   3645
            ScaleWidth      =   6585
            TabIndex        =   3
            Top             =   60
            Width           =   6585
            Begin VB.ComboBox cboPriority 
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
               ItemData        =   "LogReminders.frx":28B9
               Left            =   3330
               List            =   "LogReminders.frx":28C6
               TabIndex        =   7
               Top             =   240
               Width           =   3225
            End
            Begin VB.TextBox txtReminder_Subject 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   0
               TabIndex        =   15
               Top             =   1500
               Width           =   6525
            End
            Begin VB.ComboBox cboReminder_AssignedTo 
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
               Left            =   3360
               TabIndex        =   13
               Top             =   900
               Width           =   3195
            End
            Begin VB.TextBox txtReminder_Notes 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1410
               Left            =   0
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   2160
               Width           =   6495
            End
            Begin VB.ComboBox cboReminder_Type 
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
               ItemData        =   "LogReminders.frx":28DD
               Left            =   0
               List            =   "LogReminders.frx":28DF
               TabIndex        =   6
               Top             =   240
               Width           =   3165
            End
            Begin MSComCtl2.DTPicker txtReminder_Date 
               Height          =   345
               Left            =   0
               TabIndex        =   11
               Top             =   870
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarTitleBackColor=   8388608
               CalendarTitleForeColor=   16777215
               Format          =   405012481
               CurrentDate     =   39139
            End
            Begin MSComCtl2.DTPicker txtReminder_Time 
               Height          =   345
               Left            =   1980
               TabIndex        =   12
               Top             =   870
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "hh:mm tt"
               Format          =   405012483
               UpDown          =   -1  'True
               CurrentDate     =   39139
            End
            Begin VB.Label Label7 
               Caption         =   "Time"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   1980
               TabIndex        =   9
               Top             =   630
               Width           =   1380
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Priority"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   3360
               TabIndex        =   5
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Subject"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   0
               TabIndex        =   14
               Top             =   1230
               Width           =   645
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Notes / Tasks"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   0
               TabIndex        =   16
               Top             =   1920
               Width           =   1155
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Reminder Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   0
               TabIndex        =   4
               Top             =   0
               Width           =   1275
            End
            Begin VB.Label Label1 
               Caption         =   "Reminder Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   0
               TabIndex        =   8
               Top             =   600
               Width           =   1380
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Reminder For/Assigned To"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   3360
               TabIndex        =   10
               Top             =   630
               Width           =   2235
            End
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5100
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   32
         Top             =   5100
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   765
            MouseIcon       =   "LogReminders.frx":28E1
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":2A33
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Cancel"
            Top             =   65
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   75
            MouseIcon       =   "LogReminders.frx":2D71
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":2EC3
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Save Reminder"
            Top             =   65
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   60
         TabIndex        =   31
         Top             =   5250
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6285
      Left            =   0
      ScaleHeight     =   6285
      ScaleWidth      =   6795
      TabIndex        =   35
      Top             =   390
      Width           =   6795
      Begin XtremeReportControl.ReportControl ListView1 
         Height          =   5475
         Left            =   30
         TabIndex        =   39
         Top             =   480
         Width           =   6615
         _Version        =   655364
         _ExtentX        =   11668
         _ExtentY        =   9657
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
      End
      Begin VB.CommandButton cmdCANCELSEARCH 
         Caption         =   "X"
         Height          =   405
         Left            =   6240
         TabIndex        =   38
         Top             =   60
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1590
         TabIndex        =   36
         Top             =   60
         Width           =   4635
      End
      Begin VB.Label Label11 
         Caption         =   "Search Keyword"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   1545
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption cap 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _Version        =   655364
      _ExtentX        =   17277
      _ExtentY        =   661
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
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmSMIS_Log_Reminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ENTRY_LOGID                                             As Long
Dim RS                                                      As ADODB.Recordset

Function GetStatus(XXX)
    If XXX = "N" Then
        GetStatus = "Not Started"
    ElseIf XXX = "I" Then
        GetStatus = "In Progress"
    ElseIf XXX = "C" Then
        GetStatus = "Completed"
    ElseIf XXX = "W" Then
        GetStatus = "Waiting"
    ElseIf XXX = "D" Then
        GetStatus = "Deferred"
    Else
        GetStatus = "Not Started"
    End If

End Function

Function SetStatus(XXX)
    If XXX = "Not Started" Then
        SetStatus = "N"
    ElseIf XXX = "In Progress" Then
        SetStatus = "I"
    ElseIf XXX = "Completed" Then
        SetStatus = "C"
    ElseIf XXX = "Waiting" Then
        SetStatus = "W"
    ElseIf XXX = "Deferred" Then
        SetStatus = "D"
    Else
        SetStatus = ""
    End If

End Function

Function SetUserID(XXX)
    Dim temprs                                              As ADODB.Recordset
    'If CHANGE_USER = True Then
    If COMPANY_CODE = COMPANY_VERSION Then
        Set temprs = gconDMIS.Execute("SELECT USER_NAME FROM ALL_RAMS_USERS where USERID=" & XXX)
    Else
        Set temprs = gconDMIS.Execute("SELECT USERNAME FROM ALL_RAMS_USERS where USERID=" & XXX)
    End If
    If Not (temprs.EOF Or temprs.BOF) Then
        SetUserID = Null2String(temprs.Collect(0))
    End If
End Function

Function GetUserID(XXX)
    Dim temprs                                              As ADODB.Recordset
    If COMPANY_CODE = COMPANY_NAME Then
        Set temprs = gconDMIS.Execute("SELECT USERID FROM ALL_RAMS_USERS where USER_NAME='" & XXX & "'")
    Else
        Set temprs = gconDMIS.Execute("SELECT USERID FROM ALL_RAMS_USERS where USERNAME='" & XXX & "'")
    End If
    If Not (temprs.EOF Or temprs.BOF) Then
        GetUserID = Null2String(temprs.Collect(0))
    End If
End Function

Function GetPriority(XXX)
    If XXX = "N" Then
        GetPriority = "Normal"
    ElseIf XXX = "L" Then
        GetPriority = "Low"
    ElseIf XXX = "H" Then
        GetPriority = "High"
    End If
End Function

Function SetPriority(XXX)
    If XXX = "Normal" Then
        SetPriority = "N"
    ElseIf XXX = "Low" Then
        SetPriority = "L"
    ElseIf XXX = "High" Then
        SetPriority = "H"
    End If

End Function

Sub UpdateLog()

End Sub

Sub InitData()
    picDataEntry.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False


    Dim temprs                                              As ADODB.Recordset
    'If CHANGE_USER = True Then
    If COMPANY_CODE = COMPANY_VERSION Then
        Set temprs = gconDMIS.Execute("SELECT USER_NAME FROM ALL_RAMS_USERS order by user_name")
    Else
        Set temprs = gconDMIS.Execute("SELECT USERNAME FROM ALL_RAMS_USERS order by username")
    End If
    cboReminder_AssignedTo.Clear
    If Not (temprs.EOF Or temprs.BOF) Then
        Combo_Loadval cboReminder_AssignedTo, temprs

    End If



End Sub

Sub initMemvars()
    txtReminder_Notes = ""
    txtReminder_Date = DateValue(LOGDATE)
    txtReminder_Time = TimeValue(LOGDATE)
    txtReminder_Subject = ""
    cboPriority = ""
    cboReminder_AssignedTo = ""
    txtFollowUpNotes = ""
    cboStatus = ""
    cboReminder_Type = ""
    Call Combo_Loadval(cboReminder_Type, gconDMIS.Execute("select distinct remindertype from cris_reminders"))
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * From CRIS_Reminders  Where ENTITYTYPE='E' AND (LOGID =" & LOGID & " or userid=" & LOGID & " ) Order BY ID desc", gconDMIS, adOpenKeyset, adLockOptimistic
End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then
        'LogID, ProspectID, DateTimeCall, Duration, Subject, Comments, Bound, CalledBy, PhoneNo
        ENTRY_LOGID = RS!ID
        cboReminder_Type = Null2String(RS!REMINDERTYPE)
        txtReminder_Notes = Null2String(RS!ReminderNotes)
        txtReminder_Date.Value = DateValue(RS!DateTimeRemind)
        txtReminder_Time.Value = TimeValue(RS!DateTimeRemind)
        txtReminder_Subject = Null2String(RS!Subject)
        cboReminder_AssignedTo = SetUserID(RS!USERID)
        cboStatus = GetStatus(Null2String(RS!Status))


        If NumericVal(RS!USERID) = LOGID And NumericVal(RS!LOGID) <> LOGID Then
            picRemind_Top.Enabled = False
            picRemind_Bottom.Enabled = True
            cap.Caption = "Reminders From " & SetUserID(NumericVal(RS!LOGID))
        ElseIf NumericVal(RS!USERID) = LOGID And NumericVal(RS!LOGID) = LOGID Then
            picRemind_Top.Enabled = True
            picRemind_Bottom.Enabled = True
            cap.Caption = "Personal Reminder"

        Else
            picRemind_Top.Enabled = True
            picRemind_Bottom.Enabled = False
            cap.Caption = "Reminders For " & cboReminder_AssignedTo
        End If

        txtFollowUpNotes = Null2String(RS!followupnotes)
        labID = RS!ID
        cboPriority = GetPriority(Null2String(RS!Priority))
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub SearchID(XXX)
    ENTRY_LOGID = XXX
End Sub

Sub FillReminders()
    Dim SQL                                                 As String

    'If CHANGE_USER = True Then
    If COMPANY_CODE = COMPANY_VERSION Then
        SQL = "SELECT DATETIMEREMIND, REMINDERTYPE ,SUBJECT,"
        SQL = SQL & "(SELECT USER_NAME FROM ALL_RAMS_USERS WHERE USERID=LOGID) ,"
        SQL = SQL & "  PRIORITY , CASE STATUS   "
        SQL = SQL & " WHEN 'N' THEN 'Not Started' "
        SQL = SQL & " WHEN 'I' THEN 'In Progress' "
        SQL = SQL & " WHEN 'C' THEN 'Completed' "
        SQL = SQL & " WHEN 'W' THEN 'Waiting' "
        SQL = SQL & " WHEN 'D' THEN 'Deferred'  "
        SQL = SQL & " ELSE 'Not Started' END AS STATUS "
        SQL = SQL & " , ID From CRIS_REMINDERS"
        SQL = SQL & " WHERE ENTITYTYPE='E' and (LOGID=" & LOGID & " OR  USERID=" & LOGID & ")"
    Else
        SQL = "SELECT DATETIMEREMIND, REMINDERTYPE ,SUBJECT,"
        SQL = SQL & "(SELECT USERNAME FROM ALL_RAMS_USERS WHERE USERID=LOGID) ,"
        SQL = SQL & "  PRIORITY , CASE STATUS   "
        SQL = SQL & " WHEN 'N' THEN 'Not Started' "
        SQL = SQL & " WHEN 'I' THEN 'In Progress' "
        SQL = SQL & " WHEN 'C' THEN 'Completed' "
        SQL = SQL & " WHEN 'W' THEN 'Waiting' "
        SQL = SQL & " WHEN 'D' THEN 'Deferred'  "
        SQL = SQL & " ELSE 'Not Started' END AS STATUS "
        SQL = SQL & " , ID From CRIS_REMINDERS"
        SQL = SQL & " WHERE ENTITYTYPE='E' and (LOGID=" & LOGID & " OR  USERID=" & LOGID & ")"
    End If
    Dim RecSet                                              As ADODB.Recordset
    Set RecSet = gconDMIS.Execute(SQL)
    Dim fld                                                 As Field
    Dim j                                                   As Long
    Dim REC                                                 As XtremeReportControl.ReportRecord

    ListView1.Records.DeleteAll
    While Not RecSet.EOF
        j = j + 1
        Set REC = ListView1.Records.Add
        For Each fld In RecSet.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RecSet.MoveNext
    Wend
    ListView1.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RecSet = Nothing
End Sub

Friend Sub AddReminder()
    ENTRY_LOGID = 0
End Sub

Private Sub cboReminder_AssignedTo_LostFocus()
    If cboReminder_AssignedTo = "" Then Exit Sub
    Dim USERID
    USERID = GetUserID(cboReminder_AssignedTo)
    If USERID = "" Then
        ShowIsRequiredMsg " Assigned To"
        On Error Resume Next
        cboReminder_AssignedTo.SetFocus
        Exit Sub
    End If
End Sub

'Upating Code       : AXP-0707200712:03
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    ENTRY_LOGID = 0
    initMemvars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    Me.Refresh
    cap.Caption = "Add Reminders"
    On Error Resume Next
    cboReminder_Type.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False

    ENTRY_LOGID = 0
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200712:03
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_Reminders where ID=" & ENTRY_LOGID
        rsRefresh
        TIMER_REMIND = ""
        StoreMemVars

        LogAudit "X", "REMINDERS"
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200712:03
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    If picRemind_Top.Enabled = True Then
        'cap.Caption = " Edit Reminder"
        On Error Resume Next
        cboReminder_Type.SetFocus
    Else
        'cap.Caption = "Edit Notes & Status"
        On Error Resume Next
        cboStatus.SetFocus
    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    FillReminders
    Picture2.Visible = False
    Picture1.Visible = True
    cap.Caption = "Search Reminders"
End Sub

Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

'Upating Code       : AXP-0707200712:03
Private Sub cmdSave_Click()
    Dim t1                                                  As String
    Dim SQL                                                 As String
    Dim USERID
    On Error GoTo ErrorCode:

    USERID = GetUserID(cboReminder_AssignedTo)
    If USERID = "" Then
        ShowIsRequiredMsg " Assigned To"
        On Error Resume Next
        cboReminder_AssignedTo.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(cboReminder_Type)) = "" Then
        ShowIsRequiredMsg "Reminder Type"
        On Error Resume Next
        cboReminder_Type.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(txtReminder_Subject)) = "" Then
        ShowIsRequiredMsg "Subject Name "
        On Error Resume Next
        txtReminder_Subject.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(cboReminder_AssignedTo)) = "" Then
        ShowIsRequiredMsg "Assigned To"
        On Error Resume Next
        cboReminder_AssignedTo.SetFocus
        Exit Sub
    End If


    t1 = N2Str2Null(DateValue(txtReminder_Date) & " " & TimeValue(txtReminder_Time))

    If ENTRY_LOGID <= 0 Then
        SQL = "INSERT INTO CRIS_Reminders "
        SQL = SQL & " (USERID,  ReminderType, DateTimeRemind, ReminderNotes, Subject,Snoozed,NextTime, LOGID,ENTITYTYPE,Priority) "
        SQL = SQL & " VALUES("
        SQL = SQL & USERID & ","
        SQL = SQL & N2Str2Null(cboReminder_Type) & ","
        SQL = SQL & t1 & ","
        SQL = SQL & "'" & Replace(txtReminder_Notes, "'", "''") & "',"
        SQL = SQL & N2Str2Null(txtReminder_Subject) & ", 0, "
        SQL = SQL & t1 & "," & LOGID & ", 'E' , " & N2Str2Null(SetPriority(cboPriority)) & " )"

    Else
        If picRemind_Bottom.Enabled = True Then

            SQL = "Update CRIS_Reminders SET "
            SQL = SQL & " Followupnotes=" & N2Str2Null(txtFollowUpNotes) & ", "
            SQL = SQL & " Status ='" & SetStatus(cboStatus) & "'"
            If SetStatus(cboStatus) = "C" Then
                SQL = SQL & " ,  snoozed=1 "
            End If
            SQL = SQL & " WHERE ID=" & ENTRY_LOGID
        Else
            SQL = "Update CRIS_Reminders SET "
            SQL = SQL & " USERID=" & USERID & ", "
            SQL = SQL & " DateTimeRemind=" & t1 & ", "
            SQL = SQL & " ReminderType=" & N2Str2Null(cboReminder_Type) & ", "
            SQL = SQL & " Subject=" & N2Str2Null(txtReminder_Subject) & ", "
            SQL = SQL & " ReminderNotes='" & Replace(txtReminder_Notes, "'", "''") & "',"
            SQL = SQL & " Priority =" & N2Str2Null(SetPriority(cboPriority)) & ","
            SQL = SQL & " ENTITYTYPE ='E'"
            SQL = SQL & " WHERE ID=" & ENTRY_LOGID
        End If

    End If

    gconDMIS.Execute (SQL)

    If ENTRY_LOGID <= 0 Then
        MessagePop RecSave, "Record Added ", "New Reminder Added", 1000
    Else
        MessagePop RecSaveOk, "RecordSaved", "Reminder Updated", 1000
    End If

    ReminderModule ""
    RS.Requery
    If ENTRY_LOGID > 0 Then
        RS.Find ("ID=" & ENTRY_LOGID)
    End If
    cmdCancel.Value = True





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdCancelSearch_Click()
    Picture2.Visible = True
    Picture1.Visible = False
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitData
    initMemvars
    With ListView1
        .Columns.Add 0, "Due Date", 100, True
        .Columns.Add 1, "Type", 100, True
        .Columns.Add 2, "Subject", 100, True
        .Columns.Add 3, "Assigned To", 100, True
        .Columns.Add 4, "Priority", 100, True
        .Columns.Add 5, "Status", 100, True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.GroupRowTextBold = True
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
        .SetCustomDraw xtpCustomBeforeDrawRow
    End With
    With cboStatus
        .AddItem "Not Started"
        .AddItem "In Progress"
        .AddItem "Completed"
        .AddItem "Waiting"
        .AddItem "Deferred"
    End With

    rsRefresh
    If ENTRY_LOGID > 0 Then
        RS.Find ("ID=" & ENTRY_LOGID)
    End If
    StoreMemVars
    frmMain.Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Timer1.Enabled = True
    ENTRY_LOGID = 0
End Sub

Private Sub ListView1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)

    If Row.GroupRow = True Then Exit Sub

    If Item.Record(5).Value = "Completed" Then
        Metrics.ForeColor = vbBlack
        Metrics.Font.Bold = False

        Metrics.Font.Strikethrough = True
    Else
        If DateDiff("n", CDate(Item.Record(0).Value), Now) > 0 Then
            Metrics.ForeColor = vbRed
            Metrics.Font.Bold = True
        Else
            Metrics.ForeColor = vbBlack
            Metrics.Font.Bold = False
        End If

    End If

End Sub

Private Sub ListView1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then Exit Sub
    On Error GoTo ErrorCode
    RS.MoveFirst
    RS.Find ("ID=" & Item.Record(6).Value)
    StoreMemVars
    cmdCancelSearch_Click
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Text1_Change()
    ListView1.FilterText = Text1
    ListView1.Populate
End Sub


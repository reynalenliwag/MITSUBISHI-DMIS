VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_Reminder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Reminders"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
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
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4500
      TabIndex        =   9
      Top             =   4410
      Width           =   4500
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   2970
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   11
         Top             =   30
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   765
            MouseIcon       =   "LogReminders.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   65
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   30
            MouseIcon       =   "LogReminders.frx":0D5A
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":0EAC
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   65
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   0
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   14
         Top             =   45
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   3765
            MouseIcon       =   "LogReminders.frx":11FC
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":134E
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3015
            MouseIcon       =   "LogReminders.frx":16B4
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":1806
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   2265
            MouseIcon       =   "LogReminders.frx":1B31
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":1C83
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   1530
            MouseIcon       =   "LogReminders.frx":1FDF
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":2131
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   782
            MouseIcon       =   "LogReminders.frx":2444
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":2596
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   30
            MouseIcon       =   "LogReminders.frx":28EE
            MousePointer    =   99  'Custom
            Picture         =   "LogReminders.frx":2A40
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox picDataEntry 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4410
      Left            =   0
      ScaleHeight     =   4410
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   0
      Width           =   4665
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
         Left            =   210
         TabIndex        =   22
         Top             =   2100
         Width           =   4035
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
         Left            =   180
         TabIndex        =   21
         Top             =   1470
         Width           =   4035
      End
      Begin VB.TextBox txtReminder_Notes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   180
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2790
         Width           =   4035
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
         ItemData        =   "LogReminders.frx":2D9F
         Left            =   180
         List            =   "LogReminders.frx":2DA1
         TabIndex        =   2
         Top             =   240
         Width           =   4035
      End
      Begin MSComCtl2.DTPicker txtReminder_Date 
         Height          =   345
         Left            =   180
         TabIndex        =   4
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
         Format          =   52363265
         CurrentDate     =   39139
      End
      Begin MSComCtl2.DTPicker txtReminder_Time 
         Height          =   345
         Left            =   2040
         TabIndex        =   5
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
         Format          =   52363267
         UpDown          =   -1  'True
         CurrentDate     =   39139
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   23
         Top             =   1830
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   2520
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   1
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Reminder Date/Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   -60
         TabIndex        =   3
         Top             =   600
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Reminder For/Assigned To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   1230
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_Reminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ENTRY_LOGID                        As Long
Dim RS                                 As adodb.Recordset

Friend Sub AddReminder()
    ENTRY_LOGID = 0
End Sub
Private Sub cmdAdd_Click()
    ENTRY_LOGID = 0
    InitMemVars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    
    On Error Resume Next
    cboReminder_Type.SetFocus
End Sub




Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    
    ENTRY_LOGID = 0
    StoreMemvars
End Sub
Sub UpdateLog()

End Sub
Private Sub cmdDelete_Click()
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_Reminders where ID=" & ENTRY_LOGID
        
        rsRefresh
        StoreMemvars
        

    End If
End Sub

Private Sub cmdEdit_Click()
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    
    On Error Resume Next
    cboReminder_Type.SetFocus

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars

End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemvars

End Sub

Private Sub cmdSave_Click()
    Dim t1                             As String
    Dim SQL                            As String
    Dim USERID

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

    USERID = GetUserID(cboReminder_AssignedTo)
    t1 = N2Str2Null(DateValue(txtReminder_Date) & " " & TimeValue(txtReminder_Time))

    If ENTRY_LOGID <= 0 Then
        SQL = "INSERT INTO CRIS_Reminders "
        SQL = SQL & " (USERID,  ReminderType, DateTimeRemind, ReminderNotes, Subject,Snoozed,NextTime, LOGID,ENTITYTYPE) "
        SQL = SQL & " VALUES("
        SQL = SQL & USERID & ","
        SQL = SQL & N2Str2Null(cboReminder_Type) & ","
        SQL = SQL & t1 & ","
        SQL = SQL & N2Str2Null(txtReminder_Notes) & ","
        SQL = SQL & N2Str2Null(txtReminder_Subject) & ", 0, "
        SQL = SQL & t1 & "," & LOGID & ", 'E')"

    Else


        SQL = "Update CRIS_Reminders SET "
        SQL = SQL & " USERID=" & USERID & ", "
        SQL = SQL & " DateTimeRemind=" & t1 & ", "
        SQL = SQL & " ReminderType=" & N2Str2Null(cboReminder_Type) & ", "
        SQL = SQL & " Subject=" & N2Str2Null(txtReminder_Subject) & ", "
        SQL = SQL & " ReminderNotes=" & N2Str2Null(txtReminder_Notes) & ","
        SQL = SQL & " Snoozed =0,"
        SQL = SQL & " ENTITYTYPE ='E'"
        SQL = SQL & " WHERE ID=" & ENTRY_LOGID
    End If

    gconDMIS.Execute (SQL)

    If ENTRY_LOGID <= 0 Then
        MessagePop RecSave, "Record Added ", "New Reminder Added", 1000
    Else
        MessagePop RecSaveOk, "RecordSaved", "Reminder Updated", 1000
    End If

    UpdateLog
    RS.Requery
    If ENTRY_LOGID > 0 Then
        RS.Find ("ID=" & ENTRY_LOGID)
    End If
    cmdCancel.Value = True

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitData
    InitMemVars
    rsRefresh
    StoreMemvars

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ENTRY_LOGID = 0
End Sub

Sub InitData()
    picDataEntry.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    

    Dim TempRs                         As adodb.Recordset
    Set TempRs = gconDMIS.Execute("SELECT USERNAME FROM ALL_RAMS_USERS")
    cboReminder_AssignedTo.Clear
    If Not (TempRs.EOF Or TempRs.BOF) Then
    Combo_Loadval cboReminder_AssignedTo, TempRs
        
    End If


    
End Sub

Sub InitMemVars()
    txtReminder_Notes = ""
    txtReminder_Date = DateValue(Now)
    txtReminder_Time = TimeValue(Now)
    txtReminder_Subject = ""

End Sub

Sub rsRefresh()
    Set RS = New adodb.Recordset
    RS.Open "SELECT * From CRIS_Reminders Where LOGID =" & LOGID & " Order BY ID desc", gconDMIS, adOpenKeyset, adLockReadOnly

End Sub

Sub StoreMemvars()
    If Not RS.EOF And Not RS.BOF Then
        'LogID, ProspectID, DateTimeCall, Duration, Subject, Comments, Bound, CalledBy, PhoneNo
        ENTRY_LOGID = RS!ID
        cboReminder_Type = Null2String(RS!REMINDERTYPE)
        txtReminder_Notes = Null2String(RS!ReminderNotes)
        txtReminder_Date.Value = DateValue(RS!DateTimeRemind)
        txtReminder_Time.Value = TimeValue(RS!DateTimeRemind)
        txtReminder_Subject = Null2String(RS!Subject)
        cboReminder_AssignedTo = SetUserID(RS!USERID)
        labid = RS!ID
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub



Function SetUserID(xxx)
    Dim TempRs                         As adodb.Recordset
    Set TempRs = gconDMIS.Execute("SELECT USERNAME FROM ALL_RAMS_USERS where USERID=" & xxx)
    If Not (TempRs.EOF Or TempRs.BOF) Then
        SetUserID = Null2String(TempRs.Collect(0))
    End If
End Function
Function GetUserID(xxx)
    Dim TempRs                         As adodb.Recordset
    Set TempRs = gconDMIS.Execute("SELECT USERID FROM ALL_RAMS_USERS where USERNAME='" & xxx & "'")
    If Not (TempRs.EOF Or TempRs.BOF) Then
        GetUserID = Null2String(TempRs.Collect(0))
    End If
End Function


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_InternalReminder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internal ReminderSubjects"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogInternalReminder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture5 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4905
      TabIndex        =   18
      Top             =   5445
      Width           =   4905
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   360
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   19
         Top             =   45
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   3755
            MouseIcon       =   "LogInternalReminder.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "LogInternalReminder.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Exit Window"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3065
            MouseIcon       =   "LogInternalReminder.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "LogInternalReminder.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Delete Selected Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   2375
            MouseIcon       =   "LogInternalReminder.frx":11FF
            MousePointer    =   99  'Custom
            Picture         =   "LogInternalReminder.frx":1351
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Edit Selected Reminder"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   1685
            MouseIcon       =   "LogInternalReminder.frx":16AD
            MousePointer    =   99  'Custom
            Picture         =   "LogInternalReminder.frx":17FF
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Add Reminder"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   990
            MouseIcon       =   "LogInternalReminder.frx":1B12
            MousePointer    =   99  'Custom
            Picture         =   "LogInternalReminder.frx":1C64
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Move to Next Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   300
            MouseIcon       =   "LogInternalReminder.frx":1FBC
            MousePointer    =   99  'Custom
            Picture         =   "LogInternalReminder.frx":210E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Move to Previous Record"
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   3360
         ScaleHeight     =   900
         ScaleWidth      =   2580
         TabIndex        =   27
         Top             =   45
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   755
            MouseIcon       =   "LogInternalReminder.frx":246D
            MousePointer    =   99  'Custom
            Picture         =   "LogInternalReminder.frx":25BF
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Cancel"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   60
            MouseIcon       =   "LogInternalReminder.frx":28FD
            MousePointer    =   99  'Custom
            Picture         =   "LogInternalReminder.frx":2A4F
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Save Reminder"
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox picDataEntry 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   5445
      Left            =   0
      ScaleHeight     =   5445
      ScaleWidth      =   4905
      TabIndex        =   0
      Top             =   0
      Width           =   4905
      Begin VB.CommandButton Command2 
         Caption         =   "+ Customer"
         Height          =   345
         Left            =   3780
         TabIndex        =   13
         Top             =   2070
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+ Prospect"
         Height          =   345
         Left            =   3780
         TabIndex        =   12
         Top             =   1740
         Width           =   975
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2070
         Width           =   3585
      End
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
         ItemData        =   "LogInternalReminder.frx":2D9F
         Left            =   1650
         List            =   "LogInternalReminder.frx":2DAC
         TabIndex        =   4
         Text            =   "cboPriority"
         Top             =   510
         Width           =   3105
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
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   2730
         Width           =   4635
      End
      Begin VB.ComboBox cboReminder_AssignedTo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1650
         TabIndex        =   9
         Top             =   1320
         Width           =   3105
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
         Height          =   1950
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   3390
         Width           =   4695
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
         ItemData        =   "LogInternalReminder.frx":2DC3
         Left            =   1650
         List            =   "LogInternalReminder.frx":2DC5
         TabIndex        =   3
         Top             =   120
         Width           =   3105
      End
      Begin MSComCtl2.DTPicker txtReminder_Date 
         Height          =   345
         Left            =   1650
         TabIndex        =   6
         Top             =   900
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
         Format          =   20643841
         CurrentDate     =   39139
      End
      Begin MSComCtl2.DTPicker txtReminder_Time 
         Height          =   345
         Left            =   3510
         TabIndex        =   7
         Top             =   900
         Width           =   1230
         _ExtentX        =   2170
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
         Format          =   20643843
         UpDown          =   -1  'True
         CurrentDate     =   39139
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Attn of / Subject to"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1740
         Width           =   1575
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
         Left            =   210
         TabIndex        =   2
         Top             =   570
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
         Left            =   180
         TabIndex        =   14
         Top             =   2460
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
         Left            =   150
         TabIndex        =   16
         Top             =   3120
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
         Left            =   210
         TabIndex        =   1
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date/Time Due"
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
         Left            =   210
         TabIndex        =   5
         Top             =   900
         Width           =   1230
      End
      Begin VB.Label lblAss 
         AutoSize        =   -1  'True
         Caption         =   "SAE"
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
         Left            =   960
         TabIndex        =   8
         Top             =   1320
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_InternalReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ENTRY_LOGID                                                       As Long
Dim EmployeeOrCustomer                                                As String
Dim RS                                                                As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim AccountCode                                                       As String
Dim WithEvents SEARCHFORM                                             As frmSMIS_Mis_SearchMaster
Attribute SEARCHFORM.VB_VarHelpID = -1

Function SetUserID(XXX)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT USERNAME FROM ALL_RAMS_USERS where USERID=" & XXX)
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        SetUserID = Null2String(TEMPRS.Collect(0))
    End If
End Function

Function GetUserID(XXX)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT USERID FROM ALL_RAMS_USERS where USERNAME='" & XXX & "'")
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        GetUserID = Null2String(TEMPRS.Collect(0))
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

    Dim TEMPRS                                                        As ADODB.Recordset

    Set TEMPRS = gconDMIS.Execute("SELECT name FROM smis_vw_srep ORDER BY 1 ")
    cboReminder_AssignedTo.Clear
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        Combo_Loadval cboReminder_AssignedTo, TEMPRS
    End If

    Set TEMPRS = gconDMIS.Execute("Select Distinct ReminderType from CRIS_Reminders")
    cboReminder_Type.Clear
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        Combo_Loadval cboReminder_Type, TEMPRS
    End If


End Sub

Sub InitMemVars()
    txtReminder_Notes = ""
    txtReminder_Date = DateValue(LOGDATE)
    txtReminder_Time = TimeValue(LOGDATE)
    txtReminder_Subject = ""
    cboPriority = ""
    cboReminder_AssignedTo = ""
    Text1 = ""
    AccountCode = ""
    cboReminder_Type = ""
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * From CRIS_Reminders where EntityType='S' Order BY ID desc", gconDMIS, adOpenKeyset, adLockReadOnly

End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then
        Dim TEMPRS                                                    As ADODB.Recordset
        'LogID, ProspectID, DateTimeCall, Duration, Subject, Comments, Bound, CalledBy, PhoneNo
        ENTRY_LOGID = RS!ID
        cboReminder_Type = Null2String(RS!REMINDERTYPE)
        txtReminder_Notes = Null2String(RS!ReminderNotes)
        txtReminder_Date.Value = DateValue(RS!DateTimeRemind)
        txtReminder_Time.Value = TimeValue(RS!DateTimeRemind)
        txtReminder_Subject = Null2String(RS!Subject)
        cboReminder_AssignedTo = SetSAECode(Null2String(RS!usercode))
        labid = RS!ID
        cboPriority = GetPriority(Null2String(RS!Priority))
        AccountCode = Null2String(RS!CSCDE)
        If Left(AccountCode, 2) = "PR" Then
            Set TEMPRS = gconDMIS.Execute("SELECT TOP 1 ACCTNAME FROM CRIS_PROSPECTS WHERE PROSPECTID=" & Right(AccountCode, Len(AccountCode) - 3))
            If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
                Text1 = TEMPRS!AcctName & ""
            Else
                Text1 = ""
            End If
        ElseIf Left(AccountCode, 2) = "CS" Then
            Set TEMPRS = gconDMIS.Execute("SELECT TOP 1 ACCTNAME FROM ALL_CUSTOMER WHERE CUSCDE='" & Right(AccountCode, Len(AccountCode) - 3) & "'")
            If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
                Text1 = TEMPRS!AcctName & ""
            Else
                Text1 = ""
            End If
        End If

    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Friend Sub EditReminder(XXX, SearchID As Long)
    EmployeeOrCustomer = XXX
    AddorEdit = "EDIT"
    ENTRY_LOGID = SearchID
End Sub

Friend Sub AddReminder(XXX)
    ENTRY_LOGID = 0
    EmployeeOrCustomer = XXX
    AddorEdit = "ADD"
End Sub

Private Sub cboReminder_AssignedTo_GotFocus()
    VBComBoBoxDroppedDown cboReminder_AssignedTo
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:
    AddorEdit = "ADD"
    ENTRY_LOGID = 0
    InitMemVars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    On Error Resume Next
    cboReminder_Type.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False

    ENTRY_LOGID = 0
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_Reminders where ID=" & ENTRY_LOGID
        LogAudit "X", "INTERNAL REMINDERS" & " ATTN :" & Text1 & " SUBJECT " & txtReminder_Subject
        rsRefresh
        TIMER_REMIND = ""
        StoreMemVars
        If FormExist("MainForm") Then
            MainForm.ShowData
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:
    AddorEdit = "EDIT"
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True

    On Error Resume Next
    cboReminder_Type.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError

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

Private Sub cmdSave_Click()
    Dim t1                                                            As String
    Dim SQL                                                           As String
    Dim SALESCODE

    On Error GoTo ErrorCode:

    cboPriority.ListIndex = SetComboIndex(cboPriority)

    SALESCODE = GetSAECode(cboReminder_AssignedTo)
    If SALESCODE = "" Then
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



    t1 = N2Str2Null(DateValue(txtReminder_Date) & " " & TimeValue(txtReminder_Time))
    If AddorEdit = "ADD" Then
        SQL = "INSERT INTO CRIS_Reminders "
        SQL = SQL & " (CSCDE,  USERID, ReminderType, DateTimeRemind, ReminderNotes, Subject,Snoozed,NextTime, USERCODE,ENTITYTYPE,Priority) "
        SQL = SQL & " VALUES("
        SQL = SQL & N2Str2Null(AccountCode) & ","
        SQL = SQL & N2Str2Null(LOGID) & ","
        SQL = SQL & N2Str2Null(cboReminder_Type) & ","
        SQL = SQL & t1 & ","
        SQL = SQL & N2Str2Null(txtReminder_Notes) & ","
        SQL = SQL & N2Str2Null(txtReminder_Subject) & ", 0, "
        SQL = SQL & t1 & "," & N2Str2Null(SALESCODE) & ", 'S' , " & N2Str2Null(SetPriority(cboPriority)) & " )"
        LogAudit "A", "INTERNAL REMINDERS" & " ATTN :" & Text1 & " SUBJECT " & txtReminder_Subject
    Else
        SQL = "Update CRIS_Reminders SET "
        SQL = SQL & " CSCDE=" & N2Str2Null(AccountCode) & ", "
        SQL = SQL & " USERCODE=" & N2Str2Null(SALESCODE) & ", "
        SQL = SQL & " nexttime=" & t1 & ", "
        SQL = SQL & " ReminderType=" & N2Str2Null(cboReminder_Type) & ", "
        SQL = SQL & " Subject=" & N2Str2Null(txtReminder_Subject) & ", "
        SQL = SQL & " ReminderNotes=" & N2Str2Null(txtReminder_Notes) & ","
        SQL = SQL & " Priority =" & N2Str2Null(SetPriority(cboPriority))
        SQL = SQL & " WHERE ID=" & ENTRY_LOGID
        LogAudit "E", "INTERNAL REMINDERS" & " ATTN :" & Text1 & " SUBJECT " & txtReminder_Subject
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
    If FormExist("frmCRIS_Inquiry_TaskList") Then
        frmCRIS_Inquiry_TaskList.FillGrid
    End If

    If FormExist("MainForm") Then
        MainForm.ShowData
    End If

    cmdCancel.Value = True





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    Set SEARCHFORM = New frmSMIS_Mis_SearchMaster


    If LOGSAE <> "" Then
        SEARCHFORM.SearchForProspects " USERCODE='" & LOGSAE & "'"
    Else
        SEARCHFORM.SearchForProspects ""
    End If

    SEARCHFORM.Show 1
    On Error Resume Next
    txtReminder_Subject.SetFocus
End Sub

Private Sub Command2_Click()
    Set SEARCHFORM = New frmSMIS_Mis_SearchMaster
    SEARCHFORM.SearchForCustomers
    SEARCHFORM.Show 1
    On Error Resume Next
    txtReminder_Subject.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    SetComboMaxLength cboReminder_Type, 20
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitData
    InitMemVars
    rsRefresh

    If AddorEdit <> "ADD" Then
        If ENTRY_LOGID > 0 Then
            cmdEdit.Value = True
            RS.Find ("ID=" & ENTRY_LOGID)
        End If
    End If
    StoreMemVars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ENTRY_LOGID = 0
End Sub

Private Sub SEARCHFORM_NoSelectionMade()
    Unload SEARCHFORM
End Sub

Private Sub SEARCHFORM_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    If XSelection = "CUSTOMER" Then
        AccountCode = "CS:" & Null2String(oCusRs!CUSCDE)
        Text1 = Null2String(oCusRs!AcctName)

    ElseIf XSelection = "PROSPECT" Then
        AccountCode = "PR:" & Null2String(oCusRs!PROSPECTID)
        Text1 = Null2String(oCusRs!AcctName)

    End If
    Unload SEARCHFORM
End Sub


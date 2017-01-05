VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_ProspectReminder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prospect Reminder"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogProspectReminder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture5 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4545
      TabIndex        =   9
      Top             =   4515
      Width           =   4545
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
            Left            =   3755
            MouseIcon       =   "LogProspectReminder.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "LogProspectReminder.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Exit Window"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3065
            MouseIcon       =   "LogProspectReminder.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "LogProspectReminder.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Delete Selected Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   2375
            MouseIcon       =   "LogProspectReminder.frx":11FF
            MousePointer    =   99  'Custom
            Picture         =   "LogProspectReminder.frx":1351
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Edit Selected Reminder"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   1685
            MouseIcon       =   "LogProspectReminder.frx":16AD
            MousePointer    =   99  'Custom
            Picture         =   "LogProspectReminder.frx":17FF
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Add Reminder"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   990
            MouseIcon       =   "LogProspectReminder.frx":1B12
            MousePointer    =   99  'Custom
            Picture         =   "LogProspectReminder.frx":1C64
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Move to Next Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   300
            MouseIcon       =   "LogProspectReminder.frx":1FBC
            MousePointer    =   99  'Custom
            Picture         =   "LogProspectReminder.frx":210E
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Move to Previous Record"
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   3030
         ScaleHeight     =   900
         ScaleWidth      =   2580
         TabIndex        =   11
         Top             =   45
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   755
            MouseIcon       =   "LogProspectReminder.frx":246D
            MousePointer    =   99  'Custom
            Picture         =   "LogProspectReminder.frx":25BF
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancel"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   60
            MouseIcon       =   "LogProspectReminder.frx":28FD
            MousePointer    =   99  'Custom
            Picture         =   "LogProspectReminder.frx":2A4F
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Save Reminder"
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
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      Begin VB.TextBox txtEntity 
         Height          =   345
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1500
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Select Prospect"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3270
         TabIndex        =   25
         Top             =   1500
         Width           =   1185
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
         ItemData        =   "LogProspectReminder.frx":2D9F
         Left            =   2040
         List            =   "LogProspectReminder.frx":2DAC
         TabIndex        =   23
         Text            =   "cboPriority"
         Top             =   240
         Width           =   2235
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
         Left            =   180
         MaxLength       =   60
         TabIndex        =   21
         Top             =   2100
         Width           =   4275
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
         Height          =   1560
         Left            =   60
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2880
         Width           =   4455
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
         ItemData        =   "LogProspectReminder.frx":2DC3
         Left            =   180
         List            =   "LogProspectReminder.frx":2DC5
         TabIndex        =   2
         Top             =   240
         Width           =   1815
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
         Format          =   20643841
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
         Format          =   20643843
         UpDown          =   -1  'True
         CurrentDate     =   39139
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
         Left            =   2040
         TabIndex        =   24
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
         Left            =   180
         TabIndex        =   22
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
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   2580
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
            Underline       =   0   'False
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
      Begin VB.Label lblAss 
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
         Left            =   180
         TabIndex        =   6
         Top             =   1230
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_ProspectReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ENTRY_LOGID                                                       As Long
Dim RS                                                                As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim PROSPECTID                                                        As Variant
Dim WithEvents SEARCHFORM                                             As frmSMIS_Mis_SearchMaster
Attribute SEARCHFORM.VB_VarHelpID = -1

Function GetProspectName(XXX)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("Select Acctname from cris_prospects where prospectid=" & XXX)
    If Not TEMPRS.BOF Or Not TEMPRS.BOF Then
        GetProspectName = Null2String(TEMPRS!AcctName)
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

Sub EditReminder(XXX)
    ENTRY_LOGID = XXX
    AddorEdit = "EDIT"

End Sub

Sub UpdateLog()

End Sub

Sub InitData()
    picDataEntry.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False

    Dim TEMPRS                                                        As ADODB.Recordset
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
    txtEntity = ""
    cboReminder_Type = ""
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    If LOGSAE = "" Then
        RS.Open "SELECT * From CRIS_Reminders Where EntityType='P' Order BY ID desc", gconDMIS, adOpenKeyset, adLockReadOnly
    Else
        RS.Open "SELECT * From CRIS_Reminders Where EntityType='P' AND usercode='" & LOGSAE & "'  Order BY ID desc", gconDMIS, adOpenKeyset, adLockReadOnly
    End If



End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then

        ENTRY_LOGID = RS!ID
        cboReminder_Type = Null2String(RS!REMINDERTYPE)
        txtReminder_Notes = Null2String(RS!ReminderNotes)
        txtReminder_Date.Value = DateValue(RS!DateTimeRemind)
        txtReminder_Time.Value = TimeValue(RS!DateTimeRemind)
        txtReminder_Subject = Null2String(RS!Subject)
        PROSPECTID = Null2String(RS!CSCDE)
        txtEntity = GetProspectName(Null2String(RS!CSCDE))
        labid = RS!ID
        cboPriority = GetPriority(Null2String(RS!Priority))
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub cboPriority_LostFocus()
    cboPriority.ListIndex = SelectCombo(cboPriority, cboPriority.Text)
End Sub

Private Sub cboReminder_Type_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
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
    If Function_Access(LOGID, "Acess_DELETE", "LOG CUSTOMER REMINDERS") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "delete from CRIS_Reminders where ID=" & ENTRY_LOGID
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "X", "LOG PROSPECT VISIT", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
        LogAudit "X", "PROSPECT REMINDER" & " ASSIGNED TO SAE  :" & txtEntity & " PRIORITY" & cboPriority & " DATE REMINDER" & txtReminder_Date & ":" & txtReminder_Time
        rsRefresh
        TIMER_REMIND = ""
        StoreMemVars
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
    Dim CODE

    On Error GoTo ErrorCode:

    cboPriority.ListIndex = SetComboIndex(cboPriority)

    '    If EmployeeOrCustomer = "C" Then
    '        CODE = SetCustomerCode(txtEntity)
    '        If CODE = "" Then
    '            ShowIsRequiredMsg " Assigned To"
    '            On Error Resume Next
    '            txtEntity.SetFocus
    '            Exit Sub
    '        End If
    '    Else
    '        CODE = GetUserID(txtEntity)
    '        If CODE = "" Then
    '            ShowIsRequiredMsg " Assigned To"
    '            On Error Resume Next
    '            txtEntity.SetFocus
    '            Exit Sub
    '        End If
    '    End If

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
        SQL = SQL & " (USERID, CSCDE, ReminderType, DateTimeRemind, ReminderNotes, Subject,Snoozed,NextTime, LOGID,ENTITYTYPE,Priority) "
        SQL = SQL & " VALUES("
        SQL = SQL & N2Str2Null(LOGID) & ","
        SQL = SQL & N2Str2Null(PROSPECTID) & ","
        SQL = SQL & N2Str2Null(cboReminder_Type) & ","
        SQL = SQL & t1 & ","
        SQL = SQL & N2Str2Null(txtReminder_Notes) & ","
        SQL = SQL & N2Str2Null(txtReminder_Subject) & ", 0, "
        SQL = SQL & t1 & "," & N2Str2Null(LOGSAE) & ", 'P' , " & N2Str2Null(SetPriority(cboPriority)) & " )"

        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        NEW_LogAudit "A", "PROSPECT REMINDER", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

        'PROSPECT REMINDER
    Else
        SQL = "Update CRIS_Reminders SET "
        SQL = SQL & " nexttime=" & t1 & ", "
        SQL = SQL & " ReminderType=" & N2Str2Null(cboReminder_Type) & ", "
        SQL = SQL & " Subject=" & N2Str2Null(txtReminder_Subject) & ", "
        SQL = SQL & " ReminderNotes=" & N2Str2Null(txtReminder_Notes) & ","
        SQL = SQL & " Priority =" & N2Str2Null(SetPriority(cboPriority)) & ","
        SQL = SQL & " CSCDE ='" & PROSPECTID & "'"
        SQL = SQL & " WHERE ID=" & ENTRY_LOGID

        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "PROSPECT REMINDER", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

    End If
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

    If FormExist("MainSAE") Then
        MainSAE.ShowData
    End If

    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    If LOGSAE = "" Then
        SEARCHFORM.SearchForProspects ("")
    Else
        SEARCHFORM.SearchForProspects (" USERCODE='" & LOGSAE & "'")
    End If
    SEARCHFORM.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PROSPECT REMINDER)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(PROSPECTID), "PROSPECT REMINDER")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SetComboMaxLength cboReminder_Type, 20
    CenterMe frmMain, Me, 1
    InitData
    InitMemVars
    rsRefresh
    picDataEntry.Enabled = False
    Set SEARCHFORM = New frmSMIS_Mis_SearchMaster
    If AddorEdit <> "ADD" Then
        If ENTRY_LOGID > 0 Then
            cmdEdit.Value = True
            RS.Find ("ID=" & ENTRY_LOGID)
        End If
    Else
        'cmdAdd.Value = True
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
    PROSPECTID = Null2String(oCusRs!PROSPECTID)
    txtEntity = Null2String(oCusRs!AcctName)
    Unload SEARCHFORM
End Sub


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
   LockControls    =   -1  'True
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
         Format          =   57409537
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
         Format          =   57409539
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
Dim rs                                                                As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim PROSPECTID                                                        As Variant
Dim WithEvents SEARCHFORM                                             As frmSMIS_Mis_SearchMaster
Attribute SEARCHFORM.VB_VarHelpID = -1

Sub EditReminder(xxx)
    ENTRY_LOGID = xxx
    AddorEdit = "EDIT"

End Sub

Private Sub cboPriority_LostFocus()
    cboPriority.ListIndex = SelectCombo(cboPriority, cboPriority.Text)
End Sub

Private Sub cboReminder_Type_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

'Upating Code       : AXP-0707200712:37
Private Sub cmdADD_Click()
    On Error GoTo Errorcode:
    AddorEdit = "ADD"
    ENTRY_LOGID = 0
    initMemvars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    On Error Resume Next
    cboReminder_Type.SetFocus
    Exit Sub
Errorcode:
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
Sub UpdateLog()

End Sub
Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "LOG CUSTOMER REMINDERS") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_Reminders where ID=" & ENTRY_LOGID
        rsRefresh
        TIMER_REMIND = ""
        StoreMemVars
    End If
End Sub

'Upating Code       : AXP-0707200712:37
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    AddorEdit = "EDIT"
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    On Error Resume Next
    cboReminder_Type.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdNext_Click()
    rs.MoveNext
    If rs.EOF Then
        rs.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPrevious_Click()
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

'Upating Code       : AXP-0707200712:37
Private Sub cmdSave_Click()
    Dim t1                                                            As String
    Dim sql                                                           As String
    Dim CODE

    On Error GoTo Errorcode:

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
        sql = "INSERT INTO CRIS_Reminders "
        sql = sql & " (USERID, CSCDE, ReminderType, DateTimeRemind, ReminderNotes, Subject,Snoozed,NextTime, LOGID,ENTITYTYPE,Priority) "
        sql = sql & " VALUES("
        sql = sql & N2Str2Null(LOGID) & ","
        sql = sql & N2Str2Null(PROSPECTID) & ","
        sql = sql & N2Str2Null(cboReminder_Type) & ","
        sql = sql & t1 & ","
        sql = sql & N2Str2Null(txtReminder_Notes) & ","
        sql = sql & N2Str2Null(txtReminder_Subject) & ", 0, "
        sql = sql & t1 & "," & N2Str2Null(LOGSAE) & ", 'P' , " & N2Str2Null(SetPriority(cboPriority)) & " )"
    Else
        sql = "Update CRIS_Reminders SET "
        sql = sql & " nexttime=" & t1 & ", "
        sql = sql & " ReminderType=" & N2Str2Null(cboReminder_Type) & ", "
        sql = sql & " Subject=" & N2Str2Null(txtReminder_Subject) & ", "
        sql = sql & " ReminderNotes=" & N2Str2Null(txtReminder_Notes) & ","
        sql = sql & " Priority =" & N2Str2Null(SetPriority(cboPriority)) & ","
        sql = sql & " CSCDE ='" & PROSPECTID & "'"
        sql = sql & " WHERE ID=" & ENTRY_LOGID
    End If

    gconDMIS.Execute (sql)

    If ENTRY_LOGID <= 0 Then
        MessagePop RecSave, "Record Added ", "New Reminder Added", 1000
    Else
        MessagePop RecSaveOk, "RecordSaved", "Reminder Updated", 1000
    End If
    UpdateLog
    rs.Requery
    If ENTRY_LOGID > 0 Then
        rs.Find ("ID=" & ENTRY_LOGID)
    End If

    If FormExist("frmCRIS_Inquiry_TaskList") Then
        frmCRIS_Inquiry_TaskList.FillGrid
    End If

    If FormExist("MainSAE") Then
        MainSAE.ShowData
    End If

    cmdCancel.Value = True





    Exit Sub
Errorcode:
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
Private Sub Form_Load()
    SetComboMaxLength cboReminder_Type, 20
    CenterMe frmMain, Me, 1
    InitData
    initMemvars
    rsRefresh
    picDataEntry.Enabled = False
    Set SEARCHFORM = New frmSMIS_Mis_SearchMaster
    If AddorEdit <> "ADD" Then
        If ENTRY_LOGID > 0 Then
            cmdEdit.Value = True
            rs.Find ("ID=" & ENTRY_LOGID)
        End If
    Else
        'cmdAdd.Value = True
    End If
    StoreMemVars
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ENTRY_LOGID = 0
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

Sub initMemvars()
    txtReminder_Notes = ""
    txtReminder_Date = DateValue(Now)
    txtReminder_Time = TimeValue(Now)
    txtReminder_Subject = ""
    cboPriority = ""
    txtEntity = ""
    cboReminder_Type = ""
End Sub

Sub rsRefresh()
    Set rs = New ADODB.Recordset
    If LOGSAE = "" Then
        rs.Open "SELECT * From CRIS_Reminders Where EntityType='P' Order BY ID desc", gconDMIS, adOpenKeyset, adLockReadOnly
    Else
        rs.Open "SELECT * From CRIS_Reminders Where EntityType='P' AND usercode='" & LOGSAE & "'  Order BY ID desc", gconDMIS, adOpenKeyset, adLockReadOnly
    End If



End Sub

Sub StoreMemVars()
    If Not rs.EOF And Not rs.BOF Then

        ENTRY_LOGID = rs!ID
        cboReminder_Type = Null2String(rs!REMINDERTYPE)
        txtReminder_Notes = Null2String(rs!ReminderNotes)
        txtReminder_Date.Value = DateValue(rs!DATETIMEREMIND)
        txtReminder_Time.Value = TimeValue(rs!DATETIMEREMIND)
        txtReminder_Subject = Null2String(rs!Subject)
        PROSPECTID = Null2String(rs!CSCDE)
        txtEntity = GetProspectName(Null2String(rs!CSCDE))
        labid = rs!ID
        cboPriority = GetPriority(Null2String(rs!Priority))
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Function GetProspectName(xxx)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("Select Acctname from cris_prospects where prospectid=" & xxx)
    If Not TEMPRS.BOF Or Not TEMPRS.BOF Then
        GetProspectName = Null2String(TEMPRS!AcctName)
    End If
End Function


Function GetPriority(xxx)
    If xxx = "N" Then
        GetPriority = "Normal"
    ElseIf xxx = "L" Then
        GetPriority = "Low"
    ElseIf xxx = "H" Then
        GetPriority = "High"
    End If
End Function

Function SetPriority(xxx)
    If xxx = "Normal" Then
        SetPriority = "N"
    ElseIf xxx = "Low" Then
        SetPriority = "L"
    ElseIf xxx = "High" Then
        SetPriority = "H"
    End If

End Function

Private Sub SEARCHFORM_NoSelectionMade()
    Unload SEARCHFORM
End Sub

Private Sub SEARCHFORM_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    PROSPECTID = Null2String(oCusRs!PROSPECTID)
    txtEntity = Null2String(oCusRs!AcctName)
    Unload SEARCHFORM
End Sub

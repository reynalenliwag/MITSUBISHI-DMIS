VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCRIS_Inquiry_TaskList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Task List"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Inquirytask.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11880
   Begin VB.PictureBox picAddEdit 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   2835
      TabIndex        =   7
      Top             =   0
      Width           =   2835
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   675
         Left            =   1950
         MouseIcon       =   "Inquirytask.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "Inquirytask.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   645
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   675
         Left            =   1320
         MouseIcon       =   "Inquirytask.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "Inquirytask.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   645
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   675
         Left            =   690
         MouseIcon       =   "Inquirytask.frx":0EBF
         MousePointer    =   99  'Custom
         Picture         =   "Inquirytask.frx":1011
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   645
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   675
         Left            =   60
         MouseIcon       =   "Inquirytask.frx":136D
         MousePointer    =   99  'Custom
         Picture         =   "Inquirytask.frx":14BF
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   645
      End
   End
   Begin VB.ComboBox cboAssignedTo 
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
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   8
      Width           =   2355
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
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   450
      Width           =   2595
   End
   Begin VB.ComboBox CboStatus 
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
      Left            =   3690
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   8
      Width           =   2595
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6405
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   11298
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Inquirytask.frx":17D2
      NumItems        =   0
   End
   Begin VB.Label lblAssg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Assigned To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8040
      TabIndex        =   6
      Top             =   60
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Priority"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2940
      TabIndex        =   5
      Top             =   480
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2970
      TabIndex        =   4
      Top             =   60
      Width           =   585
   End
End
Attribute VB_Name = "frmCRIS_Inquiry_TaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TaskType                                           As String

Private Sub cboAssignedTo_Click()
    FillGrid
End Sub

Private Sub cboPriority_Click()
    FillGrid
End Sub

Private Sub CboStatus_Click()
    FillGrid
End Sub

'Upating Code       : AXP-0707200712:16
Private Sub cmdAdd_Click()
On Error GoTo Errorcode:

    frmSMIS_Log_CustomerReminder.AddReminder ("C")
    frmSMIS_Log_CustomerReminder.Show
    frmSMIS_Log_CustomerReminder.cmdAdd.Value = True





Exit Sub
Errorcode:
ShowVBError
End Sub

'Upating Code       : AXP-0707200712:16
Private Sub cmdEdit_Click()
On Error GoTo Errorcode:

    If ListView1.SelectedItem Is Nothing Then Exit Sub

    frmSMIS_Log_CustomerReminder.EditReminder "C", ListView1.SelectedItem.ListSubItems(ListView1.ColumnHeaders.Count).Text
    frmSMIS_Log_CustomerReminder.Show





Exit Sub
Errorcode:
ShowVBError
End Sub

Private Sub Form_Load()
CenterMe frmMain, Me, 1
    If TaskType = "C" Then
        AddColumnHeader "Date , Customer , Overdue, ReminderType, Subject, ReminderNotes, Priority, Status", ListView1
    Else
        AddColumnHeader "Date , Employee , Overdue, ReminderType, Subject, ReminderNotes, Priority, Status", ListView1
    End If
    ResizeColumnHeader ListView1, "10,15, 6,14,20,30,10,10"


    Dim SQL                                            As String
    With CboStatus
        .AddItem "Not Started"
        .AddItem "In Progress"
        .AddItem "Completed"
        .AddItem "Waiting"
        .AddItem "Deferred"
        .AddItem "(ANY)"
    End With

    With cboPriority
        .AddItem "Normal"
        .AddItem "High"
        .AddItem "Low"
        .AddItem "(ANY)"
    End With




    Me.Tag = TaskType
    If TaskType = "C" Then
        lblAssg = "Customer Name"
        SQL = "SELECT DISTINCT CUSTOMERNAME FROM CRIS_REMINDERS "
        SQL = SQL & " INNER JOIN CRIS_VW_ALLPROFILE ON CRIS_VW_ALLPROFILE.CUSCDE=CRIS_REMINDERS.CSCDE"
        SQL = SQL & " WHERE ENTITYTYPE='C'"
        Combo_Loadval cboAssignedTo, gconDMIS.Execute(SQL)
        cboAssignedTo.AddItem "(ANY)"
    Else
        lblAssg = "Assigned To"

        SQL = " SELECT DISTINCT USERNAME FROM CRIS_REMINDERS"
        SQL = SQL & " INNER JOIN ALL_Rams_Users ON ALL_Rams_Users.USERID=CRIS_REMINDERS.USERID"
        SQL = SQL & " WHERE ENTITYTYPE='E'"

        Combo_Loadval cboAssignedTo, gconDMIS.Execute(SQL)
        cboAssignedTo.AddItem "(ANY)"
    End If
    FillGrid
End Sub
Sub ShowTaskType(xxx)
    TaskType = xxx

End Sub
Sub FillGrid()
    Dim SQL                                            As String
    Dim Priority                                       As String
    Dim Status                                         As String
    Dim AssignedTo                                     As String

    SQL = " SELECT CONVERT(VARCHAR,DATETIMEREMIND,101) AS DATE,"
    If TaskType = "C" Then
    SQL = SQL & "(Select CustomerName from CRIS_VW_AllProfile Where CUSCDE=CSCDE),"
    Else
    SQL = SQL & "(Select USERNAME from ALL_RAMS_USERS Where USERID=USERID),"
    End If
    SQL = SQL & " CASE WHEN DATETIMEREMIND> GETDATE() THEN '0'"
    SQL = SQL & " ELSE DATEDIFF(DAY, DATETIMEREMIND,GETDATE()) END,"
    SQL = SQL & " REMINDERTYPE,"
    SQL = SQL & " SUBJECT,"
    SQL = SQL & " REMINDERNOTES,"
    SQL = SQL & " Case Priority"
    SQL = SQL & " WHEN 'H' THEN 'HIGH'"
    SQL = SQL & " WHEN 'L' THEN 'LOW'"
    SQL = SQL & " WHEN 'N' THEN 'NORMAL'"
    SQL = SQL & " END AS PRIORITY,"
    SQL = SQL & " Case Status"
    SQL = SQL & " WHEN 'N' THEN 'NOT STARTED'"
    SQL = SQL & " WHEN 'I' THEN 'IN PROGRESS'"
    SQL = SQL & " WHEN 'C' THEN 'COMPLETED'"
    SQL = SQL & " WHEN 'W' THEN 'WAITING'"
    SQL = SQL & " WHEN 'D' THEN 'DEFERRED'"
    SQL = SQL & " END As Status, ID"
    SQL = SQL & " From CRIS_REMINDERS WHERE ENTITYTYPE=" & N2Str2Null(TaskType)


    If UCase(cboPriority) <> "(ANY)" And LTrim(RTrim(cboPriority)) <> "" Then
        SQL = SQL & " AND Priority=" & N2Str2Null(GetPriority(cboPriority))
    End If

    If UCase(CboStatus) <> "(ANY)" And LTrim(RTrim(CboStatus)) <> "" Then
        SQL = SQL & " AND STATUS=" & N2Str2Null(GetPriority(cboPriority))
    End If

    If UCase(cboAssignedTo) <> "(ANY)" And LTrim(RTrim(cboAssignedTo)) <> "" Then
        If TaskType = "C" Then
            SQL = SQL & " AND CSCDE=" & GetAssignedTo(cboAssignedTo)
        Else
            SQL = SQL & " AND USERID=" & GetAssignedTo(cboAssignedTo)
        End If
    End If


    Dim TEMPRS                                         As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute(SQL)
    Listview_Loadval ListView1.ListItems, TEMPRS
End Sub
Function GetPriority(xxx)
    If UCase(cboPriority) = "NORMAL" Then
        GetPriority = "N"
    ElseIf UCase(cboPriority) = "HIGH" Then
        GetPriority = "H"
    ElseIf UCase(cboPriority) = "LOW" Then
        GetPriority = "L"
    End If

End Function


Function GetStatus(xxx)
    With CboStatus
        .AddItem "Not Started"
        .AddItem "In Progress"
        .AddItem "Completed"
        .AddItem "Waiting"
        .AddItem "Deferred"
        .AddItem "(ANY)"
    End With


End Function


Function GetAssignedTo(xxx)
    Dim TEMPRS                                         As ADODB.Recordset

    If TaskType = "C" Then
        Set TEMPRS = gconDMIS.Execute("Select CUSCDE FROM CRIS_VW_ALLPROFILE WHERE CUSTOMERNAME ='" & cboAssignedTo.Text & "'")

    Else
        Set TEMPRS = gconDMIS.Execute("Select USERID FROM ALL_RAMS_USERS WHERE USERNAME ='" & cboAssignedTo.Text & "'")
    End If

    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        GetAssignedTo = N2Str2Null(TEMPRS.Fields(0).Value)
    End If
End Function


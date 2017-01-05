VERSION 5.00
Begin VB.Form frmSMIS_Log_ReminderStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminders & Task"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Log_ReminderStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   5415
   Begin VB.TextBox Text1 
      Height          =   915
      Left            =   30
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2490
      Width           =   5295
   End
   Begin VB.ComboBox CboStatus 
      Height          =   345
      ItemData        =   "Log_ReminderStatus.frx":08CA
      Left            =   840
      List            =   "Log_ReminderStatus.frx":08CC
      TabIndex        =   2
      Text            =   "CboStatus"
      Top             =   3420
      Width           =   2115
   End
   Begin VB.TextBox txtRemind_Notes 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Log_ReminderStatus.frx":08CE
      Top             =   1260
      Width           =   5295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      MouseIcon       =   "Log_ReminderStatus.frx":08D4
      MousePointer    =   99  'Custom
      Picture         =   "Log_ReminderStatus.frx":0A26
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Exit Window"
      Top             =   3420
      Width           =   585
   End
   Begin VB.CommandButton cmdSnooze 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4110
      MouseIcon       =   "Log_ReminderStatus.frx":0D8C
      MousePointer    =   99  'Custom
      Picture         =   "Log_ReminderStatus.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Set Reminders"
      Top             =   3420
      Width           =   585
   End
   Begin VB.Label labid 
      Caption         =   "0"
      Height          =   345
      Left            =   4980
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   90
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Follow Up Notes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   2220
      Width           =   1815
   End
   Begin VB.Label lblReminderType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   930
      Width           =   5295
   End
   Begin VB.Label lblSubject 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Top             =   630
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "Date Time:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "Due By:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label lblDueBy 
      Caption         =   "Due: Saturday , April 30, 2007  12:56 PM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1260
      TabIndex        =   3
      Top             =   360
      Width           =   3435
   End
   Begin VB.Label lblDueFor 
      Caption         =   "Due: Saturday , April 30, 2007  12:56 PM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1260
      TabIndex        =   0
      Top             =   60
      Width           =   3465
   End
End
Attribute VB_Name = "frmSMIS_Log_ReminderStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public REMINDERID

Function GetStatus(XXX)
    If XXX = "N" Then
        GetStatus = "Not Started"
    ElseIf XXX = "I" Then
        GetStatus = "In Progress"
    ElseIf XXX = "C" Then
        GetStatus = "In Progress"
    ElseIf XXX = "W" Then
        GetStatus = "Waiting"
    ElseIf XXX = "D" Then
        GetStatus = "Deferred"
    Else
        GetStatus = "(ANY)"
    End If

End Function

Private Sub cbostatus_LostFocus()
    cbostatus.ListIndex = SelectCombo(cbostatus, cbostatus)
End Sub

Private Sub cmdEdit_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSnooze_Click()
    On Error GoTo ErrorCode:
    gconDMIS.Execute ("UPDATE CRIS_REMINDERS SET followupnotes='" & Repleys(Text1) & "',  status='" & Left(cbostatus.Text, 1) & "' Where ID =" & REMINDERID)
    If FormExist("MainSAE") Then
        MainSAE.ShowData
    End If

    If FormExist("MainForm") Then
        MainForm.ShowData
    End If
    Unload Me
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    With cbostatus
        .AddItem "Not Started"
        .AddItem "In Progress"
        .AddItem "Completed"
        .AddItem "Waiting"
        .AddItem "Deferred"
    End With
    Dim intDays
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select USERID, ReminderType, DateTimeRemind, ReminderNotes, Subject,  Snoozed, ID,followupnotes, NextTime ,status from CRIS_Reminders where id=" & REMINDERID)
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        lblDueFor = Null2String(TEMPRS!DateTimeRemind)
        If IsDate(TEMPRS!Nexttime) = True Then
            intDays = DateDiff("d", TEMPRS!Nexttime, LOGDATE)

            If intDays > 0 Then
                lblDueBy.ForeColor = vbRed
                lblDueBy = " Due By " & intDays & " Day(s)"
            ElseIf intDays = 0 Then
                lblDueBy.ForeColor = &H4000&
                lblDueBy = "Due Today "
            Else
                lblDueBy.ForeColor = vbBlue
                lblDueBy = " Due On  " & Abs(intDays) & " Day(s)"
            End If

        End If
        lblSubject = UCase(Null2String(TEMPRS!Subject))
        txtRemind_Notes = Null2String(TEMPRS!ReminderNotes)
        lblReminderType = UCase(Null2String(TEMPRS!REMINDERTYPE))
        Text1 = Null2String(TEMPRS!followupnotes)
        labid = TEMPRS!ID
        cbostatus = GetStatus(TEMPRS!STATUS)
    End If

End Sub


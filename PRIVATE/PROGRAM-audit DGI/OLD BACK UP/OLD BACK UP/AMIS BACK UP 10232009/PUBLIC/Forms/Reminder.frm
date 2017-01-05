VERSION 5.00
Begin VB.Form frmSMIS_Files_Reminders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminders"
   ClientHeight    =   4080
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
   Icon            =   "Reminder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5415
   Begin VB.ComboBox Combo1 
      Height          =   345
      ItemData        =   "Reminder.frx":08CA
      Left            =   2370
      List            =   "Reminder.frx":08F2
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3390
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Remind Me After"
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   3420
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Completed"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtRemind_Notes 
      Height          =   1605
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Reminder.frx":0925
      Top             =   1380
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
      Height          =   675
      Left            =   4590
      MouseIcon       =   "Reminder.frx":092B
      MousePointer    =   99  'Custom
      Picture         =   "Reminder.frx":0A7D
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit Window"
      Top             =   3240
      Width           =   705
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
      Height          =   675
      Left            =   3900
      MouseIcon       =   "Reminder.frx":0DE3
      MousePointer    =   99  'Custom
      Picture         =   "Reminder.frx":0F35
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Set Reminders"
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label lblReminderType 
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
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label lblSubject 
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
      TabIndex        =   10
      Top             =   750
      Width           =   5295
   End
   Begin VB.Label labID 
      Caption         =   "0"
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
      Height          =   210
      Left            =   2370
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label4 
      Caption         =   "Date Time:"
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
      TabIndex        =   8
      Top             =   90
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Due By:"
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
      TabIndex        =   7
      Top             =   420
      Width           =   1815
   End
   Begin VB.Label lblDueBy 
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
      Left            =   1920
      TabIndex        =   6
      Top             =   420
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "Min"
      Height          =   255
      Left            =   3450
      TabIndex        =   5
      Top             =   3450
      Width           =   225
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
      Left            =   1890
      TabIndex        =   0
      Top             =   90
      Width           =   3465
   End
End
Attribute VB_Name = "frmSMIS_Files_Reminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSnooze_Click()

    On Error GoTo Errorcode:

    If Option2.Value = True Then
        gconDMIS.Execute ("UPDATE CRIS_REMINDERS SET nexttime='" & DateAdd("n", Combo1.Text, Now) & "' Where ID =" & labID)
        ReminderModule DateAdd("n", Combo1.Text, Now)
    Else
        gconDMIS.Execute ("UPDATE CRIS_REMINDERS SET SNOOZED=1 Where ID=" & labID)
        ReminderModule ""

    End If

    Unload Me





    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub CheckOPT()
    If Option2.Value = True Then
        Combo1.Visible = True
        Combo1.ListIndex = 0
        cmdSnooze.Visible = True
        Label1.Visible = True
    Else
        Combo1.Visible = False
        cmdSnooze.Visible = True
        Label1.Visible = False
    End If
End Sub

Private Sub Form_Load()
    frmMain.Timer1.Enabled = False
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim TEMPRS                                         As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select TOP 1 USERID, ReminderType, DateTimeRemind, ReminderNotes, Subject,  Snoozed, ID, NextTime  from CRIS_Reminders where SNOOZED=0 and  MONTH(nexttime)=MONTH(getdate()) and YEAR(nexttime)=YEAR(getdate()) and USERID=" & LOGID & " and nexttime < = getdate()")
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        lblDueFor = Null2String(TEMPRS!DateTimeRemind)
        lblDueBy = DateDiff("n", TEMPRS!DateTimeRemind, LOGTIME) & " Minutes"
        lblSubject = Null2String(TEMPRS!Subject)
        txtRemind_Notes = TEMPRS!ReminderNotes & ""
        lblReminderType = TEMPRS!REMINDERTYPE & ""
        labID = TEMPRS!ID
        Option2.Value = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.Timer1.Enabled = True
End Sub

Private Sub Option1_Click()
    CheckOPT


End Sub

Private Sub Option2_Click()
    CheckOPT
End Sub


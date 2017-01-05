VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_CustomerReminder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customers  Reminders"
   ClientHeight    =   6300
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
   Icon            =   "Log_CustomerReminder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4545
      TabIndex        =   9
      Top             =   5385
      Width           =   4545
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   0
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   14
         Top             =   -45
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   3780
            MouseIcon       =   "Log_CustomerReminder.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "Log_CustomerReminder.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3090
            MouseIcon       =   "Log_CustomerReminder.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "Log_CustomerReminder.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Delete Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   2400
            MouseIcon       =   "Log_CustomerReminder.frx":11FF
            MousePointer    =   99  'Custom
            Picture         =   "Log_CustomerReminder.frx":1351
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   1710
            MouseIcon       =   "Log_CustomerReminder.frx":16AD
            MousePointer    =   99  'Custom
            Picture         =   "Log_CustomerReminder.frx":17FF
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1020
            MouseIcon       =   "Log_CustomerReminder.frx":1B12
            MousePointer    =   99  'Custom
            Picture         =   "Log_CustomerReminder.frx":1C64
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   330
            MouseIcon       =   "Log_CustomerReminder.frx":1FBC
            MousePointer    =   99  'Custom
            Picture         =   "Log_CustomerReminder.frx":210E
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   2970
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   11
         Top             =   -30
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   780
            MouseIcon       =   "Log_CustomerReminder.frx":246D
            MousePointer    =   99  'Custom
            Picture         =   "Log_CustomerReminder.frx":25BF
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancel"
            Top             =   65
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   90
            MouseIcon       =   "Log_CustomerReminder.frx":28FD
            MousePointer    =   99  'Custom
            Picture         =   "Log_CustomerReminder.frx":2A4F
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Save Entry"
            Top             =   65
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
      Height          =   5385
      Left            =   0
      ScaleHeight     =   5385
      ScaleWidth      =   6825
      TabIndex        =   0
      Top             =   0
      Width           =   6825
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
         ItemData        =   "Log_CustomerReminder.frx":2D9F
         Left            =   2040
         List            =   "Log_CustomerReminder.frx":2DAC
         TabIndex        =   24
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
         Left            =   150
         TabIndex        =   22
         Top             =   2100
         Width           =   4035
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
         Height          =   2400
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
         ItemData        =   "Log_CustomerReminder.frx":2DC3
         Left            =   180
         List            =   "Log_CustomerReminder.frx":2DC5
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
         Format          =   50724865
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
         Format          =   50724867
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
         Left            =   2100
         TabIndex        =   25
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
Attribute VB_Name = "frmSMIS_Log_CustomerReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ENTRY_LOGID                         As Long
Dim EmployeeOrCustomer                  As String
Dim RS                                  As ADODB.Recordset
Dim AddorEdit                           As String
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

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    ENTRY_LOGID = 0
    InitMemVars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True

    On Error Resume Next
    cboReminder_Type.SetFocus
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
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_Reminders where ID=" & ENTRY_LOGID
        rsREFRESH
        TIMER_REMIND = ""
        StoreMemVars
    End If
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
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
    Dim t1                              As String
    Dim SQL                             As String
    Dim code

    cboPriority.ListIndex = SetComboIndex(cboPriority)

    If EmployeeOrCustomer = "C" Then
        code = SetCustomerCode(cboReminder_AssignedTo)
        If code = "" Then
            ShowIsRequiredMsg " Assigned To"
            On Error Resume Next
            cboReminder_AssignedTo.SetFocus
            Exit Sub
        End If
    Else
        code = GetUserID(cboReminder_AssignedTo)
        If code = "" Then
            ShowIsRequiredMsg " Assigned To"
            On Error Resume Next
            cboReminder_AssignedTo.SetFocus
            Exit Sub
        End If
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
        If EmployeeOrCustomer = "C" Then
            SQL = SQL & " (CSCDE,  "
        Else
            SQL = SQL & " (USERID,  "
        End If
        SQL = SQL & " ReminderType, DateTimeRemind, ReminderNotes, Subject,Snoozed,NextTime, LOGID,ENTITYTYPE,Priority) "
        SQL = SQL & " VALUES("
        If EmployeeOrCustomer = "C" Then
            SQL = SQL & N2Str2Null(code) & ","
        Else
            SQL = SQL & code & ","
        End If
        SQL = SQL & N2Str2Null(cboReminder_Type) & ","
        SQL = SQL & t1 & ","
        SQL = SQL & N2Str2Null(txtReminder_Notes) & ","
        SQL = SQL & N2Str2Null(txtReminder_Subject) & ", 0, "
        SQL = SQL & t1 & "," & LOGID & ", '" & EmployeeOrCustomer & "' , " & N2Str2Null(SetPriority(cboPriority)) & " )"
    Else
        SQL = "Update CRIS_Reminders SET "
        If EmployeeOrCustomer = "C" Then
            SQL = SQL & " CSCDE=" & N2Str2Null(code) & ", "
        Else
            SQL = SQL & " USERID=" & code & ", "
        End If
        SQL = SQL & " DateTimeRemind=" & t1 & ", "
        SQL = SQL & " ReminderType=" & N2Str2Null(cboReminder_Type) & ", "
        SQL = SQL & " Subject=" & N2Str2Null(txtReminder_Subject) & ", "
        SQL = SQL & " ReminderNotes=" & N2Str2Null(txtReminder_Notes) & ","
        SQL = SQL & " Snoozed =0,"
        SQL = SQL & " Priority =" & N2Str2Null(SetPriority(cboPriority)) & ","
        SQL = SQL & " ENTITYTYPE ='" & EmployeeOrCustomer & "'"
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
    If FormExist("frmCRIS_INQUIRY_TaskList") Then
        frmCRIS_Inquiry_TaskList.FillGrid
    End If

    cmdCancel.Value = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub
Private Sub Form_Load()

    SetComboMaxLength cboReminder_Type, 20
    CenterMe frmMain, Me, 1
    InitData
    InitMemVars
    rsREFRESH
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

Sub InitData()
    picDataEntry.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False

    Dim temprs                          As ADODB.Recordset
    If EmployeeOrCustomer = "E" Then
        Set temprs = gconDMIS.Execute("SELECT USERNAME FROM ALL_RAMS_USERS ORDER BY 1 ")
        lblAss = "Reminder For/Assigned To"
        Me.Caption = "Employees Reminders"
    Else
        Set temprs = gconDMIS.Execute("SELECT CUSTOMERNAME FROM CRIS_VW_ALLProfile ORDER BY 1 ")
        lblAss = "Customer Name"
        Me.Caption = "Customer Reminders"
    End If

    cboReminder_AssignedTo.Clear
    If Not (temprs.EOF Or temprs.BOF) Then
        Combo_Loadval cboReminder_AssignedTo, temprs
    End If

    Set temprs = gconDMIS.Execute("Select Distinct ReminderType from CRIS_Reminders")
    cboReminder_Type.Clear
    If Not (temprs.EOF Or temprs.BOF) Then
        Combo_Loadval cboReminder_Type, temprs
    End If


End Sub
Function SetCustomerCode(XXX)
    Dim temprs                          As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select CUSCDE from CRIS_VW_ALLProfile where CustomerName='" & XXX & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        SetCustomerCode = Null2String(temprs!CUSCDE)
    End If
End Function
Function GetCustomerName(XXX)
    Dim temprs                          As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select CustomerName from CRIS_VW_ALLProfile where CUSCDE='" & XXX & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        GetCustomerName = Null2String(temprs!CUSTOMERNAME)
    End If

End Function
Sub InitMemVars()
    txtReminder_Notes = ""
    txtReminder_Date = DateValue(Now)
    txtReminder_Time = TimeValue(Now)
    txtReminder_Subject = ""
    cboPriority = ""
    cboReminder_AssignedTo = ""
    cboReminder_Type = ""
End Sub

Sub rsREFRESH()
    Set RS = New ADODB.Recordset
    If EmployeeOrCustomer = "E" Then
        RS.Open "SELECT * From CRIS_Reminders Where LOGID =" & LOGID & " AND EntityType='E' Order BY ID desc", gconDMIS, adOpenKeyset, adLockReadOnly
    Else
        RS.Open "SELECT * From CRIS_Reminders Where EntityType='C' Order BY ID desc", gconDMIS, adOpenKeyset, adLockReadOnly
    End If


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

        If EmployeeOrCustomer = "E" Then
            cboReminder_AssignedTo = SetUserID(RS!USERID)
        Else
            cboReminder_AssignedTo = GetCustomerName(RS!CSCDE)
        End If

        labID = RS!ID
        cboPriority = GetPriority(Null2String(RS!Priority))

    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub



Function SetUserID(XXX)
    Dim temprs                          As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT USERNAME FROM ALL_RAMS_USERS where USERID=" & XXX)
    If Not (temprs.EOF Or temprs.BOF) Then
        SetUserID = Null2String(temprs.Collect(0))
    End If
End Function
Function GetUserID(XXX)
    Dim temprs                          As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT USERID FROM ALL_RAMS_USERS where USERNAME='" & XXX & "'")
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


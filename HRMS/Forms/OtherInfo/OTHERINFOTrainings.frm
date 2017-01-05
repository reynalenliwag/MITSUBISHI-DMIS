VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOTrainings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRAINING PLANS"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
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
      Height          =   705
      Left            =   8490
      MouseIcon       =   "OTHERINFOTrainings.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOTrainings.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exit Window"
      Top             =   5010
      Width           =   705
   End
   Begin VB.PictureBox picTrainingPlan 
      Height          =   4470
      Left            =   1440
      ScaleHeight     =   4410
      ScaleWidth      =   6075
      TabIndex        =   6
      Top             =   180
      Width           =   6135
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   5070
         MouseIcon       =   "OTHERINFOTrainings.frx":04B8
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOTrainings.frx":060A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancel Entry"
         Top             =   3690
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   4380
         MouseIcon       =   "OTHERINFOTrainings.frx":0948
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOTrainings.frx":0A9A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Save Entry"
         Top             =   3690
         Width           =   705
      End
      Begin VB.TextBox txtNameOfTraining 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   2250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   990
         Width           =   3795
      End
      Begin VB.TextBox txtDateComp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2250
         TabIndex        =   4
         Top             =   3150
         Width           =   3465
      End
      Begin VB.TextBox txtDesPer 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   2250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   60
         Width           =   3795
      End
      Begin VB.TextBox txtDevType 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2250
         TabIndex        =   3
         Top             =   2745
         Width           =   2925
      End
      Begin VB.TextBox txtDateSched 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2250
         TabIndex        =   2
         Top             =   2295
         Width           =   2925
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Training /Development to fulfill Desired Performance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   60
         TabIndex        =   13
         Top             =   990
         Width           =   2115
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3300
         TabIndex        =   12
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Completed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   3150
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Desired Performance /Competency"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   60
         TabIndex        =   10
         Top             =   90
         Width           =   2115
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Training/Dev. Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   2820
         Width           =   2715
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approx. Date to be Sched."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   60
         TabIndex        =   7
         Top             =   2250
         Width           =   2115
      End
   End
   Begin wizButton.cmd cmdTrainingPlan 
      Height          =   4590
      Left            =   1380
      TabIndex        =   8
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8096
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "OTHERINFOTrainings.frx":0DEA
   End
   Begin MSComctlLib.ListView lstTrainingPlan 
      Height          =   4860
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   8573
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "OTHERINFOTrainings.frx":0E06
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Desired Perf/Comp."
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name Of Training"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date Scheduled"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dev. Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Completed"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ID"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   7800
      MouseIcon       =   "OTHERINFOTrainings.frx":0F68
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOTrainings.frx":10BA
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Delete Selected Record"
      Top             =   5010
      Width           =   705
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   7110
      MouseIcon       =   "OTHERINFOTrainings.frx":13E5
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOTrainings.frx":1537
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Edit Selected Record"
      Top             =   5010
      Width           =   705
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6420
      MouseIcon       =   "OTHERINFOTrainings.frx":1893
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOTrainings.frx":19E5
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Add Record"
      Top             =   5010
      Width           =   705
   End
End
Attribute VB_Name = "frmOTHERINFOTrainings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsTrainingPlan                                                    As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim EMPLIVIL                                                          As String

Sub InitMemVars()
    txtDesPer.Text = ""
    txtNameOfTraining.Text = ""
    txtDateSched.Text = ""
    txtDevType.Text = ""
    txtDateComp.Text = ""
End Sub

Sub StoreEntry(XXX As Variant)
    Set rsTrainingPlan = New ADODB.Recordset
    Set rsTrainingPlan = gconDMIS.Execute("Select * from HRMS_TrainingPlan Where ID = " & XXX)
    If Not rsTrainingPlan.EOF And Not rsTrainingPlan.BOF Then
        labID.Caption = rsTrainingPlan!ID
        txtDesPer.Text = Null2String(rsTrainingPlan!DesPer)
        txtNameOfTraining.Text = Null2String(rsTrainingPlan!NameOfTraining)
        txtDateSched.Text = Null2String(rsTrainingPlan!DateSched)
        txtDevType.Text = Null2String(rsTrainingPlan!DevType)
        txtDateComp.Text = Null2String(rsTrainingPlan!DateComp)
    End If
End Sub

Sub FillGrid()
    lstTrainingPlan.Sorted = False: lstTrainingPlan.ListItems.Clear

    lstTrainingPlan.Enabled = False
    Set rsTrainingPlan = New ADODB.Recordset
    Set rsTrainingPlan = gconDMIS.Execute("select DesPer,NameOfTraining,DateSched,DevType,DateComp,ID from HRMS_TrainingPlan where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsTrainingPlan.EOF And rsTrainingPlan.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstTrainingPlan.ListItems, rsTrainingPlan
        lstTrainingPlan.Refresh
        lstTrainingPlan.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else
        EmptyRecord = True
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

'Upating Code       : AXP-0707200712:02
Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_ADD", "DATA ENTRY") = False Then Exit Sub
    cmdTrainingPlan.ZOrder 0: picTrainingPlan.ZOrder 0
    AddorEdit = "ADD"
    InitMemVars
    On Error Resume Next
    txtNameOfTraining.SetFocus





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    cmdTrainingPlan.ZOrder 1: picTrainingPlan.ZOrder 1
End Sub

'Upating Code       : AXP-0707200712:02
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_DELETE", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstTrainingPlan.SelectedItem.SubItems(5) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_TrainingPlan Where ID = " & lstTrainingPlan.SelectedItem.SubItems(5))

                Call LogAudit("X", "DELETE EMPLOYEE TRAINING PLAN", EMPLOYEE_NO)
                Call ShowDeletedMsg
                Call FillGrid
            End If
        End If
    End If

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_EDIT", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstTrainingPlan.SelectedItem.SubItems(5) <> "" Then
            StoreEntry lstTrainingPlan.SelectedItem.SubItems(5)
            cmdTrainingPlan.ZOrder 0: picTrainingPlan.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200712:02
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    cmdTrainingPlan.ZOrder 1: picTrainingPlan.ZOrder 1
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_TrainingPlan " & _
                         "(EMPLEVEL,EMPNO,DesPer,NameOfTraining,DateSched,DevType,DateComp,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(txtDesPer.Text) & "," & N2Str2Null(txtNameOfTraining.Text) & "," & N2Str2Null(txtDateSched.Text) & "," & N2Str2Null(txtDevType.Text) & "," & N2Str2Null(txtDateComp.Text) & ",'" & LOGCODE & "','" & LOGDATE & "')"

        Call LogAudit("A", "ADD EMPLOYEE TRAINING PLAN", EMPLOYEE_NO)
    Else
        gconDMIS.Execute "update HRMS_TrainingPlan set " & _
                       " DesPer = " & N2Str2Null(txtDesPer.Text) & "," & _
                       " NameOfTraining = " & N2Str2Null(txtNameOfTraining.Text) & "," & _
                       " DateSched = " & N2Str2Null(txtDateSched.Text) & "," & _
                       " DevType = " & N2Str2Null(txtDevType.Text) & "," & _
                       " DateComp = " & N2Str2Null(txtDateComp.Text) & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE TRAINING PLAN", EMPLOYEE_NO)
    End If
    Call FillGrid

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdTrainingPlan.ZOrder 1: picTrainingPlan.ZOrder 1
        Case vbKeyF3
            cmdTrainingPlan.ZOrder 0: picTrainingPlan.ZOrder 0
            AddorEdit = "ADD"
            InitMemVars
            On Error Resume Next
            txtNameOfTraining.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstTrainingPlan.SelectedItem.SubItems(5) <> "" Then
                    StoreEntry lstTrainingPlan.SelectedItem.SubItems(5)
                    cmdTrainingPlan.ZOrder 0: picTrainingPlan.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstTrainingPlan.SelectedItem.SubItems(5) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_TrainingPlan Where ID = " & lstTrainingPlan.SelectedItem.SubItems(5))
                        ShowDeletedMsg
                        FillGrid
                    End If
                End If
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    cmdTrainingPlan.ZOrder 1: picTrainingPlan.ZOrder 1
    FillGrid
End Sub

Private Sub lstTrainingPlan_DblClick()
    If EmptyRecord = False Then
        If lstTrainingPlan.SelectedItem.SubItems(5) <> "" Then
            StoreEntry lstTrainingPlan.SelectedItem.SubItems(5)
            cmdTrainingPlan.ZOrder 0: picTrainingPlan.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub lstTrainingPlan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstTrainingPlan
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub


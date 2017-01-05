VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCSMS_TrainingPlan 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6435
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   90
      ScaleHeight     =   3585
      ScaleWidth      =   6225
      TabIndex        =   15
      Top             =   60
      Width           =   6255
      Begin VB.TextBox txtNameOfTraining 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   2340
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1020
         Width           =   3795
      End
      Begin VB.TextBox txtDateComp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   4
         Top             =   3180
         Width           =   3465
      End
      Begin VB.TextBox txtDevType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2775
         Width           =   2925
      End
      Begin VB.TextBox txtDateSched 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   2
         Top             =   2325
         Width           =   2925
      End
      Begin VB.TextBox txtDesPer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   2340
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   90
         Width           =   3795
      End
      Begin VB.Label lblID 
         BackColor       =   &H000000FF&
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
         Height          =   240
         Left            =   5430
         TabIndex        =   22
         Top             =   2670
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Training /Development to fulfill Desired Performance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   150
         TabIndex        =   21
         Top             =   1020
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
         Left            =   3390
         TabIndex        =   20
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Completed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   3180
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Desired Performance /Competency"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   150
         TabIndex        =   18
         Top             =   120
         Width           =   2115
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Training/Dev. Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   2850
         Width           =   2715
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approx. Date to be Sched."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   150
         TabIndex        =   16
         Top             =   2280
         Width           =   2115
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   90
      ScaleHeight     =   1185
      ScaleWidth      =   6225
      TabIndex        =   14
      Top             =   3750
      Width           =   6255
      Begin MSComctlLib.ListView lsvLIST 
         Height          =   1125
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   1984
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
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Desired Performance"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name Of Training"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date To Take"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Finish"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3270
      ScaleHeight     =   885
      ScaleWidth      =   3105
      TabIndex        =   13
      Top             =   5040
      Width           =   3105
      Begin VB.CommandButton cmdedit 
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
         Height          =   795
         Left            =   720
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmddelete 
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
         Height          =   795
         Left            =   1440
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":04AE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":0600
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdexit 
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
         Height          =   795
         Left            =   2160
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":092B
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":0A7D
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdadd 
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
         Height          =   795
         Left            =   0
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":0DE3
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":0F35
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picture2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4740
      ScaleHeight     =   765
      ScaleWidth      =   1605
      TabIndex        =   12
      Top             =   5010
      Width           =   1605
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
         Left            =   690
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":1248
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":139A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel Entry"
         Top             =   30
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
         Left            =   0
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":16D8
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":182A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCSMS_TrainingPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADD_EDIT As String

Private Sub cmdAdd_Click()
    If frmMainMenu.lblTS.Caption = "SA" Then If Function_Access(LOGID, "ACESS_ADD", "SERVICE ADVISER TRAININGS PLAN") = False Then Exit Sub
    If frmMainMenu.lblTS.Caption = "TECH" Then If Function_Access(LOGID, "ACESS_ADD", "TECHNICIAN TRAININGS PLAN") = False Then Exit Sub
    
    ADD_EDIT = "ADD"
    
    Call initMemvars
    Call DisAblePics(True)
    
    txtDesPer.SetFocus
End Sub

Sub DisAblePics(COND As Boolean)
    picture1.Visible = Not COND
    picture2.Visible = COND
    picList.Enabled = Not COND
    picInfo.Enabled = COND
End Sub

Sub initMemvars()
    txtDateComp.Text = ""
    txtDateSched.Text = ""
    txtDesPer.Text = ""
    txtDevType.Text = ""
    txtNameOfTraining.Text = ""
End Sub

Private Sub cmdCancel_Click()
    If Not lsvLIST.ListItems.Count = 0 Then
        Call lsvLIST_Click
        Call DisAblePics(False)
    Else
        Call initMemvars
        Call DisAblePics(False)
    End If
End Sub

Private Sub cmdDelete_Click()
    If Not txtDesPer.Text = "" Then
        If frmMainMenu.lblTS.Caption = "SA" Then If Function_Access(LOGID, "ACESS_DELETE", "SERVICE TRAINING PLAN") = False Then Exit Sub
        If frmMainMenu.lblTS.Caption = "TECH" Then If Function_Access(LOGID, "ACESS_DELETE", "TECHNICIAN TRAINING PLAN") = False Then Exit Sub

        If MsgBox("Delete " & txtDesPer.Text & " Training Plan", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            gconDMIS.Execute ("Delete From CSMS_TRAINING_PLAN Where Empno = '" & frmCSMSEmpNo.txtEmpNo & _
                "' And ID = " & lblID.Caption & "")
            
            Call FillTheList
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If Not txtDesPer.Text = "" Then
        If frmMainMenu.lblTS.Caption = "SA" Then If Function_Access(LOGID, "ACESS_EDIT", "SERVICE TRAINING PLAN") = False Then Exit Sub
        If frmMainMenu.lblTS.Caption = "TECH" Then If Function_Access(LOGID, "ACESS_EDIT", "TECHNICIAN TRAINING PLAN") = False Then Exit Sub
        
        ADD_EDIT = "EDIT"
        
        Call DisAblePics(True)
        txtDesPer.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim vPER As String
    Dim vTRAIN As String
    Dim vDATE As String
    Dim vTYPE As String
    Dim vDATEC As String
    
    If txtDesPer.Text = "" Then
        MsgBox "Desired Performance Cannot be Blank", vbInformation, "Training Plan"
        txtDesPer.SetFocus
        Exit Sub
    End If
    If txtNameOfTraining.Text = "" Then
        MsgBox "Name of Training Cannot be Blank", vbInformation, "Training Plan"
        txtNameOfTraining.SetFocus
        Exit Sub
    End If
    If txtDateSched.Text = "" Then
        MsgBox "Date Schedule Cannot be Blank", vbInformation, "Training Plan"
        txtDateSched.SetFocus
        Exit Sub
    End If
    If txtDevType.Text = "" Then
        MsgBox "Type Cannot be Blank", vbInformation, "Training Plan"
        txtDevType.SetFocus
        Exit Sub
    End If
    
    vPER = N2Str2Null(txtDesPer.Text)
    vTRAIN = N2Str2Null(txtNameOfTraining.Text)
    vDATE = N2Str2Null(txtDateSched.Text)
    vTYPE = N2Str2Null(txtDevType.Text)
    vDATEC = N2Str2Null(txtDateComp.Text)
    
    If ADD_EDIT = "ADD" Then
        gconDMIS.Execute ("Insert Into CSMS_TRAINING_PLAN (Empno,Desired_Per,NameOFTrain,SchedDate,TrainType,DateComp) VALUES('" & frmCSMSEmpNo.txtEmpNo.Text & _
            "'," & vPER & "," & vTRAIN & "," & vDATE & "," & vTYPE & "," & vDATEC & ")")
    
        If frmMainMenu.lblTS.Caption = "SA" Then Call LogAudit("A", "SERVICE ADVISER TRAINING PLAN", vPER)
        If frmMainMenu.lblTS.Caption = "TECH" Then Call LogAudit("A", "TRAINING TRAINING PLAN", vPER)
        Call ShowSuccessFullyAdded
        
        Call DisAblePics(False)
        Call FillTheList
    Else
        gconDMIS.Execute ("UPDATE CSMS_TRAINING_PLAN SET Desired_Per = " & vPER & _
            ",NameOfTrain = " & vTRAIN & _
            ",SchedDate = " & vDATE & _
            ",TrainType = " & vTYPE & _
            ",DateComp = " & vDATEC & _
            " Where Empno = '" & frmCSMSEmpNo.txtEmpNo.Text & "' And ID = " & lblID.Caption & "")
    
        If frmMainMenu.lblTS.Caption = "SA" Then Call LogAudit("E", "SERVICE ADVISER TRAINING PLAN", vPER & "-" & lblID)
        If frmMainMenu.lblTS.Caption = "TECH" Then Call LogAudit("E", "TRAINING TRAINING PLAN", vPER & "-" & lblID)
        Call ShowSuccessFullyUpdated
        
        Call DisAblePics(False)
        Call FillTheList
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call FillTheList
End Sub

Sub FillTheList()
    Dim rsTmp As New ADODB.Recordset
    Dim ITEM As ListItem
    
    Set rsTmp = gconDMIS.Execute("Select * from CSMS_TRAINING_PLAN Where Empno = '" & frmCSMSEmpNo.txtEmpNo.Text & "' Order by ID ASC")
    lsvLIST.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvLIST.ListItems.Add(, , Null2String(rsTmp!Desired_Per))
            ITEM.SubItems(1) = Null2String(rsTmp!NameOFTrain)
            ITEM.SubItems(2) = Null2String(rsTmp!SchedDate)
            ITEM.SubItems(3) = Null2String(rsTmp!TrainType)
            ITEM.SubItems(4) = Null2String(rsTmp!DateComp)
            ITEM.SubItems(5) = Null2String(rsTmp!ID)
            
            rsTmp.MoveNext
        Loop
    End If
    
    If Not lsvLIST.ListItems.Count = 0 Then Call lsvLIST_Click
    
    Set rsTmp = Nothing
End Sub

Private Sub lsvLIST_Click()
    Dim INDEX As Double
    
    If Not lsvLIST.ListItems.Count = 0 Then
        With lsvLIST
            INDEX = .SelectedItem.INDEX
        
            txtDesPer.Text = .ListItems(INDEX).Text
            txtNameOfTraining.Text = .ListItems(INDEX).SubItems(1)
            txtDateSched.Text = .ListItems(INDEX).SubItems(2)
            txtDevType.Text = .ListItems(INDEX).SubItems(3)
            txtDateComp.Text = .ListItems(INDEX).SubItems(4)
            lblID.Caption = .ListItems(INDEX).SubItems(5)
        End With
    End If
End Sub

Private Sub txtDateComp_LostFocus()
If IsDate(txtDateComp) = False Then txtDateComp = "": Exit Sub
    txtDateComp = FormatDateTime(txtDateComp, vbShortDate)
End Sub

Private Sub txtDateSched_LostFocus()
    If IsDate(txtDateSched) = False Then txtDateSched = "": Exit Sub
    txtDateSched = FormatDateTime(txtDateSched, vbShortDate)
End Sub

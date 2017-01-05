VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCSMS_TRAIN 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5070
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
      Height          =   765
      Left            =   2220
      ScaleHeight     =   765
      ScaleWidth      =   2925
      TabIndex        =   12
      Top             =   4590
      Width           =   2925
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
         Height          =   705
         Left            =   2070
         MouseIcon       =   "frmCSMS_TRAIN.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   705
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
         Height          =   705
         Left            =   1380
         MouseIcon       =   "frmCSMS_TRAIN.frx":04B8
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN.frx":060A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   705
      End
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
         Height          =   705
         Left            =   690
         MouseIcon       =   "frmCSMS_TRAIN.frx":0935
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN.frx":0A87
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   705
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
         Height          =   705
         Left            =   0
         MouseIcon       =   "frmCSMS_TRAIN.frx":0DE3
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN.frx":0F35
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   705
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
      Left            =   3600
      ScaleHeight     =   765
      ScaleWidth      =   1605
      TabIndex        =   20
      Top             =   4560
      Width           =   1605
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
         MouseIcon       =   "frmCSMS_TRAIN.frx":1248
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN.frx":139A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
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
         MouseIcon       =   "frmCSMS_TRAIN.frx":16EA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN.frx":183C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox PicInfo 
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
      Height          =   3195
      Left            =   60
      ScaleHeight     =   3165
      ScaleWidth      =   4905
      TabIndex        =   14
      Top             =   30
      Width           =   4935
      Begin VB.TextBox txtDet 
         Appearance      =   0  'Flat
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
         Height          =   1185
         Left            =   1350
         MaxLength       =   200
         TabIndex        =   4
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtSponsor 
         Appearance      =   0  'Flat
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
         Left            =   1350
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtTraining 
         Appearance      =   0  'Flat
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
         Height          =   675
         Left            =   1350
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   120
         Width           =   3495
      End
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
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
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1200
         Width           =   3480
      End
      Begin VB.TextBox txtMonYear 
         Appearance      =   0  'Flat
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
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Details Of Training"
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
         Height          =   600
         Index           =   1
         Left            =   90
         TabIndex        =   22
         Top             =   1980
         Width           =   1155
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
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   630
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   1590
         TabIndex        =   19
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sponsor"
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
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   1590
         Width           =   795
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Training Title"
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
         Height          =   510
         Left            =   90
         TabIndex        =   17
         Top             =   150
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
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
         Left            =   90
         TabIndex        =   16
         Top             =   1230
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month-Year"
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
         Left            =   90
         TabIndex        =   15
         Top             =   870
         Width           =   1170
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
      Left            =   60
      ScaleHeight     =   1185
      ScaleWidth      =   4905
      TabIndex        =   13
      Top             =   3300
      Width           =   4935
      Begin MSComctlLib.ListView lsvLIST 
         Height          =   1125
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   4845
         _ExtentX        =   8546
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
            Text            =   "Training"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Month-Year"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sponsor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Details"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCSMS_TRAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADD_EDIT As String

Private Sub cmdAdd_Click()
    If frmMainMenu.lblTS.Caption = "SA" Then If Function_Access(LOGID, "ACESS_ADD", "SERVICE TRAINING") = False Then Exit Sub
    If frmMainMenu.lblTS.Caption = "TECH" Then If Function_Access(LOGID, "ACESS_ADD", "TECHNICIAN TRAINING") = False Then Exit Sub
    ADD_EDIT = "ADD"
        
    Call initMemvars
    Call DisAblePics(True)
    
    txtTraining.SetFocus
End Sub

Sub DisAblePics(COND As Boolean)
    picture1.Visible = Not COND
    picture2.Visible = COND
    picList.Enabled = Not COND
    PicInfo.Enabled = COND
End Sub

Sub initMemvars()
    txtTraining.Text = ""
    txtMonYear.Text = ""
    txtPlace.Text = ""
    txtSponsor.Text = ""
    txtDet.Text = ""
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
    If Not txtTraining.Text = "" Then
        If frmMainMenu.lblTS.Caption = "SA" Then If Function_Access(LOGID, "ACESS_DELETE", "SERVICE TRAINING") = False Then Exit Sub
        If frmMainMenu.lblTS.Caption = "TECH" Then If Function_Access(LOGID, "ACESS_DELETE", "TECHNICIAN TRAINING") = False Then Exit Sub
    
        If MsgBox("Delete " & txtTraining.Text & " Training Plan", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            gconDMIS.Execute ("Delete From CSMS_TRAIN Where Empno = '" & frmCSMSEmpNo.txtEmpNo & _
                "' And ID = " & lblID.Caption & "")
            
            Call FillTheList
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If Not txtTraining.Text = "" Then
        If frmMainMenu.lblTS.Caption = "SA" Then If Function_Access(LOGID, "ACESS_EDIT", "SERVICE TRAINING") = False Then Exit Sub
        If frmMainMenu.lblTS.Caption = "TECH" Then If Function_Access(LOGID, "ACESS_EDIT", "TECHNICIAN TRAINING") = False Then Exit Sub
        
        ADD_EDIT = "EDIT"
        
        Call DisAblePics(True)
        txtTraining.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim vTRAIN As String
    Dim vDATE As String
    Dim vPLACE As String
    Dim vSPON As String
    Dim vDET As String
    
    If txtTraining.Text = "" Then
        MsgBox "Training Title Cannot be Blank", vbInformation, "Training Plan"
        txtTraining.SetFocus
        Exit Sub
    End If
    If txtMonYear.Text = "" Then
        MsgBox "Date Taken Cannot be Blank", vbInformation, "Training Plan"
        txtMonYear.SetFocus
        Exit Sub
    End If
    If txtPlace.Text = "" Then
        MsgBox "Place Name Cannot be Blank", vbInformation, "Training Plan"
        txtPlace.SetFocus
        Exit Sub
    End If
    If txtSponsor.Text = "" Then
        MsgBox "Sponsor Cannot be Blank", vbInformation, "Training Plan"
        txtSponsor.SetFocus
        Exit Sub
    End If
    If txtDet.Text = "" Then
        MsgBox "Training Details Cannot be Blank", vbInformation, "Training Plan"
        txtDet.SetFocus
        Exit Sub
    End If
    
    vTRAIN = N2Str2Null(txtTraining.Text)
    vDATE = N2Str2Null(txtDet.Text)
    vPLACE = N2Str2Null(txtPlace.Text)
    vSPON = N2Str2Null(txtSponsor.Text)
    vDET = N2Str2Null(txtDet.Text)
    
    If ADD_EDIT = "ADD" Then
        gconDMIS.Execute ("Insert Into CSMS_TRAIN (Empno,Training,Deyt,Place,Sponsor,Details) VALUES('" & frmCSMSEmpNo.txtEmpNo.Text & _
            "'," & vTRAIN & "," & vDATE & "," & vPLACE & "," & vSPON & "," & vDET & ")")
    
        If frmMainMenu.lblTS.Caption = "SA" Then Call LogAudit("A", "SERVICE ADVISER TRAINING ATTENDED", vTRAIN)
        If frmMainMenu.lblTS.Caption = "TECH" Then Call LogAudit("A", "TRAINING TRAINING ATTENDED", vTRAIN)
        Call ShowSuccessFullyAdded
        
        Call DisAblePics(False)
        Call FillTheList
    Else
        gconDMIS.Execute ("UPDATE CSMS_TRAIN SET Training = " & vTRAIN & _
            ",Deyt = " & vDATE & _
            ",Place = " & vPLACE & _
            ",Sponsor = " & vSPON & _
            ",Details = " & vDET & _
            " Where Empno = '" & frmCSMSEmpNo.txtEmpNo.Text & "' And ID = " & lblID.Caption & "")
    
        If frmMainMenu.lblTS.Caption = "SA" Then Call LogAudit("E", "SERVICE ADVISER TRAINING ATTENDED", vTRAIN & "-" & lblID)
        If frmMainMenu.lblTS.Caption = "TECH" Then Call LogAudit("E", "TRAINING TRAINING ATTENDED", vTRAIN & "-" & lblID)
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
    
    Set rsTmp = gconDMIS.Execute("Select * from CSMS_TRAIN Where Empno = '" & frmCSMSEmpNo.txtEmpNo.Text & "' Order by ID ASC")
    lsvLIST.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvLIST.ListItems.Add(, , Null2String(rsTmp!training))
            ITEM.SubItems(1) = Null2String(rsTmp!Deyt)
            ITEM.SubItems(2) = Null2String(rsTmp!Place)
            ITEM.SubItems(3) = Null2String(rsTmp!Sponsor)
            ITEM.SubItems(4) = Null2String(rsTmp!Details)
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
        
            txtTraining.Text = .ListItems(INDEX).Text
            txtMonYear.Text = .ListItems(INDEX).SubItems(1)
            txtPlace.Text = .ListItems(INDEX).SubItems(2)
            txtSponsor.Text = .ListItems(INDEX).SubItems(3)
            txtDet.Text = .ListItems(INDEX).SubItems(4)
            lblID.Caption = .ListItems(INDEX).SubItems(5)
        End With
    End If
End Sub

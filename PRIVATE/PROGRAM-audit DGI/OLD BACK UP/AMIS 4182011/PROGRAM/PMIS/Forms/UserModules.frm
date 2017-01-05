VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUserModules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " User Modules"
   ClientHeight    =   5415
   ClientLeft      =   450
   ClientTop       =   645
   ClientWidth     =   9960
   ForeColor       =   &H00F5F5F5&
   Icon            =   "UserModules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
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
      Height          =   720
      Left            =   4365
      MouseIcon       =   "UserModules.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "UserModules.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3825
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4365
      MouseIcon       =   "UserModules.frx":12E4
      MousePointer    =   99  'Custom
      Picture         =   "UserModules.frx":1436
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4575
      Width           =   1005
   End
   Begin VB.CommandButton cmdAdd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4365
      MouseIcon       =   "UserModules.frx":1774
      MousePointer    =   99  'Custom
      Picture         =   "UserModules.frx":18C6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   675
      Width           =   1005
   End
   Begin VB.CommandButton cmdRemove 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4365
      MouseIcon       =   "UserModules.frx":1BA8
      MousePointer    =   99  'Custom
      Picture         =   "UserModules.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1005
   End
   Begin MSComctlLib.ListView lvwGranted 
      Height          =   4650
      Left            =   5520
      TabIndex        =   4
      Top             =   705
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8202
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Selected Modules"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwModules 
      Height          =   4650
      Left            =   75
      TabIndex        =   5
      Top             =   705
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8202
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MODULES"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User can access the ff. modules:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   5550
      TabIndex        =   9
      Top             =   375
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   75
      Width           =   975
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1305
      TabIndex        =   7
      Top             =   75
      Width           =   5475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Modules:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   105
      TabIndex        =   6
      Top             =   375
      Width           =   1395
   End
End
Attribute VB_Name = "frmUserModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs                   As ADODB.Recordset
Dim NameStr              As String
Dim QueryStr             As String
Dim itemStr              As String
Dim itemInt              As Integer
Dim mUserID              As Long
Dim i                    As Integer

Public Property Let UserID(sUserID As Long)
    mUserID = sUserID
End Property

Public Property Let Username(sNamestr As String)
    NameStr = sNamestr
End Property

Private Sub cmdAdd_Click()
    If Me.lvwModules.ListItems.Count > 0 Then
        Me.lvwGranted.ListItems.Add , , Me.lvwModules.SelectedItem
        itemInt = Me.lvwGranted.ListItems.Count
        Me.lvwGranted.ListItems(itemInt).SubItems(1) = Me.lvwModules.SelectedItem.SubItems(1)
        Me.lvwModules.ListItems.Remove Me.lvwModules.SelectedItem.Index
        Me.cmdSave.Enabled = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    If Me.lvwGranted.ListItems.Count > 0 Then
        Me.lvwModules.ListItems.Add , , Me.lvwGranted.SelectedItem
        itemInt = Me.lvwModules.ListItems.Count
        Me.lvwModules.ListItems(itemInt).SubItems(1) = Me.lvwGranted.SelectedItem.SubItems(1)
        Me.lvwGranted.ListItems.Remove Me.lvwGranted.SelectedItem.Index
        If Me.lvwGranted.ListItems.Count = 0 Then
            Me.cmdSave.Enabled = False
        Else
            Me.cmdSave.Enabled = True
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    Call Save_Modules
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.lblUsername.Caption = NameStr
    '== Get Modules list ==
    'Set rs = gconDMIS.Execute("select description,id from PMIS_User_Modules order by description")
    Set rs = gconDMIS.Execute("select description,id from PMIS_User_Modules order by id")
    If Not (rs.BOF And rs.EOF) Then
        lvwModules.Enabled = True
        Listview_Loadval Me.lvwModules.ListItems, rs
    Else
        lvwModules.Enabled = False
        Me.lvwModules.ListItems.Clear
    End If
    '== Get Current Modules assigned ==
    'Set rs = gconDMIS.Execute("SELECT PMIS_user_modules.DESCRIPTION, user_modules.ID from PMIS_User_Modules INNER JOIN " _
     '        & "USERACCESS ON user_modules.ID = USERACCESS.MODULE_ID WHERE (USERACCESS.USERID = " & mUserID & ") order by user_modules.DESCRIPTION")
    Set rs = gconDMIS.Execute("SELECT PMIS_user_modules.DESCRIPTION, PMIS_user_modules.ID from PMIS_user_modules INNER JOIN " _
                            & "PMIS_USERACCESS ON PMIS_user_modules.ID = PMIS_USERACCESS.MODULE_ID WHERE (PMIS_USERACCESS.USERID = " & mUserID & ") order by PMIS_user_modules.ID")
                            
    If Not (rs.BOF And rs.EOF) Then
        lvwGranted.Enabled = True
        Listview_Loadval Me.lvwGranted.ListItems, rs
    Else
        lvwModules.Enabled = False
        Me.lvwModules.ListItems.Clear
    End If
    Call Select_Modules
    Set rs = Nothing
    If Me.lblUsername = "" Then
        Me.cmdAdd.Enabled = False
        Me.cmdRemove.Enabled = False
    End If
End Sub

Private Sub lvwGranted_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.cmdRemove.Enabled = True
End Sub

Private Sub Select_Modules()
    Dim x                As Integer
    x = 1
    For i = 1 To Me.lvwGranted.ListItems.Count
        While x <= Me.lvwModules.ListItems.Count
            If Me.lvwGranted.ListItems(i) = Me.lvwModules.ListItems(x) Then
                Me.lvwModules.ListItems.Remove x
                x = x + 1
            End If
            x = x + 1
        Wend
        x = 1
    Next i
End Sub

Private Sub Save_Modules()
    On Error GoTo errZone
    gconDMIS.Execute ("Delete from PMIS_USERACCESS where userid = " & mUserID & " ")
    For i = 1 To Me.lvwGranted.ListItems.Count
        gconDMIS.Execute ("Insert into PMIS_USERACCESS(module_id, userid) values(" & _
                          Me.lvwGranted.ListItems(i).SubItems(1) & ", " & _
                          mUserID & ") ")
    Next i
    MsgBox "Selected modules successfully added!", vbInformation, ""
    Me.cmdSave.Enabled = False
    Exit Sub
errZone:
    MsgBox Err.Description
End Sub

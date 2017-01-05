VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOTraining 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRAININGS AND SEMINARS ATTENDED"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6990
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
      Left            =   6120
      MouseIcon       =   "OTHERINFOTraining.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOTraining.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Exit Window"
      Top             =   4020
      Width           =   705
   End
   Begin VB.PictureBox picTraining 
      Height          =   3015
      Left            =   990
      ScaleHeight     =   2955
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   210
      Width           =   4935
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
         Left            =   4020
         MouseIcon       =   "OTHERINFOTraining.frx":04B8
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOTraining.frx":060A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancel Entry"
         Top             =   2040
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
         Left            =   3330
         MouseIcon       =   "OTHERINFOTraining.frx":0948
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOTraining.frx":0A9A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save Entry"
         Top             =   2040
         Width           =   705
      End
      Begin VB.TextBox txtSponsor 
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
         Left            =   1320
         TabIndex        =   3
         Top             =   1500
         Width           =   3495
      End
      Begin VB.TextBox txtTraining 
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
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   60
         Width           =   3495
      End
      Begin VB.TextBox txtPlace 
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
         Left            =   1320
         TabIndex        =   2
         Top             =   1140
         Width           =   3480
      End
      Begin VB.TextBox txtMonYear 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   780
         Width           =   3495
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
         Left            =   1560
         TabIndex        =   11
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   1530
         Width           =   1215
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
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   1170
         Width           =   2715
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   810
         Width           =   1155
      End
   End
   Begin wizButton.cmd cmdTraining 
      Height          =   3135
      Left            =   930
      TabIndex        =   7
      Top             =   150
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
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
      MICON           =   "OTHERINFOTraining.frx":0DEA
   End
   Begin MSComctlLib.ListView lstTraining 
      Height          =   3900
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6879
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
      MouseIcon       =   "OTHERINFOTraining.frx":0E06
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TRAINING"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MO./YR."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PLACE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SPONSOR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
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
      Left            =   5430
      MouseIcon       =   "OTHERINFOTraining.frx":0F68
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOTraining.frx":10BA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Delete Selected Record"
      Top             =   4020
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
      Left            =   4740
      MouseIcon       =   "OTHERINFOTraining.frx":13E5
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOTraining.frx":1537
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Edit Selected Record"
      Top             =   4020
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
      Left            =   4050
      MouseIcon       =   "OTHERINFOTraining.frx":1893
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOTraining.frx":19E5
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add Record"
      Top             =   4020
      Width           =   705
   End
End
Attribute VB_Name = "frmOTHERINFOTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsTraining                                                        As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim EMPLIVIL                                                          As String

Sub InitMemVars()
    txtTraining.Text = ""
    txtMonYear.Text = ""
    txtPlace.Text = ""
    txtSponsor.Text = ""
End Sub

Sub StoreEntry(XXX As Variant)
    Set rsTraining = New ADODB.Recordset
    Set rsTraining = gconDMIS.Execute("Select * from HRMS_Training Where ID = " & XXX)
    If Not rsTraining.EOF And Not rsTraining.BOF Then
        labID.Caption = rsTraining!ID
        txtTraining.Text = Null2String(rsTraining!Training)
        txtMonYear.Text = Null2String(rsTraining!MonYear)
        txtPlace.Text = Null2String(rsTraining!Place)
        txtSponsor.Text = Null2String(rsTraining!Sponsor)
    End If
End Sub

Sub FillGrid()
    lstTraining.Sorted = False: lstTraining.ListItems.Clear
    lstTraining.Enabled = False
    Set rsTraining = New ADODB.Recordset
    Set rsTraining = gconDMIS.Execute("select Training,MonYear,Place,Sponsor,ID from HRMS_Training where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsTraining.EOF And rsTraining.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstTraining.ListItems, rsTraining
        lstTraining.Refresh
        lstTraining.Enabled = True
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

    cmdTraining.ZOrder 0: picTraining.ZOrder 0
    AddorEdit = "ADD"
    InitMemVars
    On Error Resume Next
    txtTraining.SetFocus

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    cmdTraining.ZOrder 1: picTraining.ZOrder 1
End Sub

'Upating Code       : AXP-0707200712:02
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_DELETE", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstTraining.SelectedItem.SubItems(4) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_Training Where ID = " & lstTraining.SelectedItem.SubItems(4))

                Call LogAudit("X", "DELETE EMPLOYEE TRAINING", EMPLOYEE_NO)
                Call ShowDeletedMsg
                Call FillGrid
            End If
        End If
    End If

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

'Upating Code       : AXP-0707200712:02
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_EDIT", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstTraining.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstTraining.SelectedItem.SubItems(4)
            cmdTraining.ZOrder 0: picTraining.ZOrder 0
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

'Upating Code       : AXP-0707200712:01
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    cmdTraining.ZOrder 1: picTraining.ZOrder 1
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_Training " & _
                         "(EMPLEVEL,EMPNO,Training,MonYear,Place,Sponsor,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(txtTraining.Text) & "," & N2Str2Null(txtMonYear.Text) & "," & N2Str2Null(txtPlace.Text) & "," & N2Str2Null(txtSponsor.Text) & ",'" & LOGCODE & "','" & LOGDATE & "')"

        Call LogAudit("A", "ADD EMPLOYEE TRAINING", EMPLOYEE_NO)
    Else
        gconDMIS.Execute "update HRMS_Training set " & _
                       " Training = " & N2Str2Null(txtTraining.Text) & "," & _
                       " MonYear = " & N2Str2Null(txtMonYear.Text) & "," & _
                       " Place = " & N2Str2Null(txtPlace.Text) & "," & _
                       " Sponsor = " & N2Str2Null(txtSponsor.Text) & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE TRAINING", EMPLOYEE_NO)
    End If
    Call FillGrid

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdTraining.ZOrder 1: picTraining.ZOrder 1
        Case vbKeyF3
            cmdTraining.ZOrder 0: picTraining.ZOrder 0
            AddorEdit = "ADD"
            InitMemVars
            On Error Resume Next
            txtTraining.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstTraining.SelectedItem.SubItems(4) <> "" Then
                    StoreEntry lstTraining.SelectedItem.SubItems(4)
                    cmdTraining.ZOrder 0: picTraining.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstTraining.SelectedItem.SubItems(4) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_Training Where ID = " & lstTraining.SelectedItem.SubItems(4))
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
    cmdTraining.ZOrder 1: picTraining.ZOrder 1
    FillGrid
End Sub

Private Sub lstTraining_DblClick()
    If EmptyRecord = False Then
        If lstTraining.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstTraining.SelectedItem.SubItems(4)
            cmdTraining.ZOrder 0: picTraining.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub lstTraining_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstTraining
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


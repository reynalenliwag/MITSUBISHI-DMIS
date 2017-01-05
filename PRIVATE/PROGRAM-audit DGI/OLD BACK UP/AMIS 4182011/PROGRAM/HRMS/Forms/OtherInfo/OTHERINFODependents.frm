VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFODependents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEPENDENTS"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7035
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
      Left            =   6150
      MouseIcon       =   "OTHERINFODependents.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFODependents.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Exit Window"
      Top             =   2610
      Width           =   705
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
      Left            =   5460
      MouseIcon       =   "OTHERINFODependents.frx":04B8
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFODependents.frx":060A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Delete Selected Record"
      Top             =   2610
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
      Left            =   4770
      MouseIcon       =   "OTHERINFODependents.frx":0935
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFODependents.frx":0A87
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Edit Selected Record"
      Top             =   2610
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
      Left            =   4080
      MouseIcon       =   "OTHERINFODependents.frx":0DE3
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFODependents.frx":0F35
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add Record"
      Top             =   2610
      Width           =   705
   End
   Begin VB.PictureBox picDependents 
      Height          =   1875
      Left            =   1005
      ScaleHeight     =   1815
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   435
      Width           =   4935
      Begin VB.CheckBox chkTaxClaim 
         Caption         =   "Tax Claim"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3510
         TabIndex        =   3
         Top             =   780
         Width           =   1305
      End
      Begin VB.ComboBox cboRelation 
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
         Left            =   1080
         TabIndex        =   2
         Text            =   "cboEmpStatus"
         Top             =   780
         Width           =   2175
      End
      Begin VB.TextBox txtBirthday 
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
         Left            =   1080
         TabIndex        =   1
         Top             =   420
         Width           =   3720
      End
      Begin VB.TextBox txtFullName 
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
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   0
         Top             =   60
         Width           =   3765
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
         Left            =   2370
         MouseIcon       =   "OTHERINFODependents.frx":1248
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFODependents.frx":139A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancel Entry"
         Top             =   1140
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
         Left            =   1680
         MouseIcon       =   "OTHERINFODependents.frx":16D8
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFODependents.frx":182A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save Entry"
         Top             =   1140
         Width           =   705
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
         Left            =   1170
         TabIndex        =   10
         Top             =   90
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
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
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday"
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
         TabIndex        =   7
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Top             =   90
         Width           =   555
      End
   End
   Begin wizButton.cmd cmdDependents 
      Height          =   1995
      Left            =   945
      TabIndex        =   9
      Top             =   375
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3519
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
      MICON           =   "OTHERINFODependents.frx":1B7A
   End
   Begin MSComctlLib.ListView lstDependents 
      Height          =   2505
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4419
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
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
      MouseIcon       =   "OTHERINFODependents.frx":1B96
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FULL NAME"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "BIRTHDAY"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RELATION"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "TAX Claim"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID"
         Object.Width           =   2
      EndProperty
   End
End
Attribute VB_Name = "frmOTHERINFODependents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsDEPENDENTS                                                      As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim EMPLIVIL                                                          As String

Sub InitMemVars()
    txtFullName.Text = ""
    txtBirthday.Text = ""
    cboRelation.Text = ""
    chkTaxClaim.Value = 1
End Sub

Sub StoreEntry(XXX As Variant)
    Set rsDEPENDENTS = New ADODB.Recordset
    Set rsDEPENDENTS = gconDMIS.Execute("Select * from HRMS_Dependents Where ID = " & XXX)
    If Not rsDEPENDENTS.EOF And Not rsDEPENDENTS.BOF Then
        labID.Caption = rsDEPENDENTS!ID
        txtFullName.Text = Null2String(rsDEPENDENTS!FULLNAME)
        txtBirthday.Text = Null2String(rsDEPENDENTS!BIRTHDAY)
        cboRelation.Text = Null2String(rsDEPENDENTS!Relation)
        If Null2String(rsDEPENDENTS!TAXCLAIM) = "Y" Then chkTaxClaim.Value = 1 Else chkTaxClaim.Value = 0
    End If
End Sub

Sub FillGrid()
    lstDependents.Sorted = False: lstDependents.ListItems.Clear
    lstDependents.Enabled = False
    Set rsDEPENDENTS = New ADODB.Recordset
    Set rsDEPENDENTS = gconDMIS.Execute("select FullName,Birthday,Relation,TaxClaim,ID from HRMS_Dependents where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsDEPENDENTS.EOF And rsDEPENDENTS.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstDependents.ListItems, rsDEPENDENTS
        lstDependents.Refresh
        lstDependents.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else
        EmptyRecord = True
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    'If Function_Access(LOGID, "ACESS_ADD", "DATA ENTRY") = False Then Exit Sub
    cmdDependents.ZOrder 0: picDependents.ZOrder 0
    AddorEdit = "ADD"
    InitMemVars
    On Error Resume Next
    txtFullName.SetFocus
End Sub

Private Sub cmdCancel_Click()
    cmdDependents.ZOrder 1: picDependents.ZOrder 1
End Sub

Private Sub cmdDelete_Click()
    'If Function_Access(LOGID, "ACESS_DELETE", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstDependents.SelectedItem.SubItems(4) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_Dependents Where ID = " & lstDependents.SelectedItem.SubItems(4))

                Call LogAudit("X", "DELETE EMPLOYEE DEPENDENTS", EMPLOYEE_NO & "-" & lstDependents.SelectedItem.SubItems(4))
                Call ShowDeletedMsg
                Call FillGrid
            End If
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    'If Function_Access(LOGID, "ACESS_EDIT", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstDependents.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstDependents.SelectedItem.SubItems(4)
            cmdDependents.ZOrder 0: picDependents.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:55
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    cmdDependents.ZOrder 1: picDependents.ZOrder 1
    Dim TAXCLAIM                                                      As String
    If chkTaxClaim.Value = 1 Then TAXCLAIM = "'Y'" Else TAXCLAIM = "'N'"
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_Dependents " & _
                         "(EMPLEVEL,EMPNO,FULLNAME,BIRTHDAY,RELATION,TAXCLAIM,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(txtFullName.Text) & "," & N2Str2Null(txtBirthday.Text) & "," & N2Str2Null(cboRelation.Text) & "," & TAXCLAIM & ",'" & LOGCODE & "','" & LOGDATE & "')"

        Call LogAudit("A", "ADD EMPLOYEE DEPENDENTS", EMPLOYEE_NO)
    Else
        gconDMIS.Execute "update HRMS_Dependents set " & _
                       " FULLNAME = " & N2Str2Null(txtFullName.Text) & "," & _
                       " BIRTHDAY = " & N2Str2Null(txtBirthday.Text) & "," & _
                       " RELATION = " & N2Str2Null(cboRelation.Text) & "," & _
                       " TAXCLAIM = " & TAXCLAIM & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE DEPENDENTS", EMPLOYEE_NO)
    End If
    FillGrid
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdDependents.ZOrder 1: picDependents.ZOrder 1
        Case vbKeyF3
            cmdDependents.ZOrder 0: picDependents.ZOrder 0
            AddorEdit = "ADD"
            InitMemVars
            On Error Resume Next
            txtFullName.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstDependents.SelectedItem.SubItems(4) <> "" Then
                    StoreEntry lstDependents.SelectedItem.SubItems(4)
                    cmdDependents.ZOrder 0: picDependents.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstDependents.SelectedItem.SubItems(4) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_Dependents Where ID = " & lstDependents.SelectedItem.SubItems(4))
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
    cboRelation.Clear
    cboRelation.AddItem "Son"
    cboRelation.AddItem "Daughter"
    cboRelation.AddItem "Father"
    cboRelation.AddItem "Mother"
    cboRelation.AddItem "Brother"
    cboRelation.AddItem "Sister"
    cmdDependents.ZOrder 1: picDependents.ZOrder 1
    FillGrid
End Sub

Private Sub lstDependents_DblClick()
    If EmptyRecord = False Then
        If lstDependents.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstDependents.SelectedItem.SubItems(4)
            cmdDependents.ZOrder 0: picDependents.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub txtBirthday_GotFocus()
    txtBirthday.Text = Format(txtBirthday.Text, "MM/DD/YYYY")
End Sub

Private Sub txtBirthday_LostFocus()
    If IsDate(txtBirthday.Text) = True Then
        txtBirthday.Text = Format(txtBirthday.Text, "DD-MMM-YYYY")
    Else
        txtBirthday.Text = ""
    End If
End Sub

Private Sub lstDependents_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstDependents
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


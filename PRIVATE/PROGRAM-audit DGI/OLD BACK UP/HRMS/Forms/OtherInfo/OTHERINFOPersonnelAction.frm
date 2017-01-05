VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOPersonalAction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERSONNEL ACTION"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8595
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
      Left            =   7770
      MouseIcon       =   "OTHERINFOPersonnelAction.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPersonnelAction.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exit Window"
      Top             =   3960
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
      Left            =   7080
      MouseIcon       =   "OTHERINFOPersonnelAction.frx":04B8
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPersonnelAction.frx":060A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Delete Selected Record"
      Top             =   3960
      Width           =   705
   End
   Begin VB.PictureBox picPersonalAction 
      Height          =   3090
      Left            =   1860
      ScaleHeight     =   3030
      ScaleWidth      =   4875
      TabIndex        =   7
      Top             =   450
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
         MouseIcon       =   "OTHERINFOPersonnelAction.frx":0935
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOPersonnelAction.frx":0A87
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cancel Entry"
         Top             =   2310
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
         MouseIcon       =   "OTHERINFOPersonnelAction.frx":0DC5
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOPersonnelAction.frx":0F17
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save Entry"
         Top             =   2310
         Width           =   705
      End
      Begin VB.TextBox txtAnnBasicSalary 
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
         TabIndex        =   5
         Top             =   1845
         Width           =   1635
      End
      Begin VB.ComboBox cboDepartment 
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
         TabIndex        =   4
         Text            =   "cboEmpStatus"
         Top             =   1500
         Width           =   3525
      End
      Begin VB.ComboBox cboActionDesc 
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
         TabIndex        =   2
         Text            =   "cboEmpStatus"
         Top             =   780
         Width           =   3525
      End
      Begin VB.TextBox txtFrom 
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
         TabIndex        =   0
         Top             =   60
         Width           =   1455
      End
      Begin VB.TextBox txtPosition 
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
         TabIndex        =   3
         Top             =   1140
         Width           =   2985
      End
      Begin VB.TextBox txtTo 
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
         TabIndex        =   1
         Top             =   420
         Width           =   1455
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
         Left            =   1350
         TabIndex        =   15
         Top             =   90
         Width           =   885
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Annual Basic Salary"
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
         Height          =   705
         Left            =   60
         TabIndex        =   14
         Top             =   1890
         Width           =   1185
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         TabIndex        =   13
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         TabIndex        =   12
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         TabIndex        =   11
         Top             =   1170
         Width           =   2715
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Action Desc"
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
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Top             =   450
         Width           =   1215
      End
   End
   Begin wizButton.cmd cmdPersonalAction 
      Height          =   3210
      Left            =   1800
      TabIndex        =   10
      Top             =   375
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5662
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
      MICON           =   "OTHERINFOPersonnelAction.frx":1267
   End
   Begin MSComctlLib.ListView lstPersonalAction 
      Height          =   3840
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   6773
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
      MouseIcon       =   "OTHERINFOPersonnelAction.frx":1283
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "From Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "To Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Action Desc"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Position"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Annual Basic"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ID"
         Object.Width           =   2
      EndProperty
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
      Left            =   6390
      MouseIcon       =   "OTHERINFOPersonnelAction.frx":13E5
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPersonnelAction.frx":1537
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Edit Selected Record"
      Top             =   3960
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
      Left            =   5700
      MouseIcon       =   "OTHERINFOPersonnelAction.frx":1893
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPersonnelAction.frx":19E5
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Add Record"
      Top             =   3960
      Width           =   705
   End
End
Attribute VB_Name = "frmOTHERINFOPersonalAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsPersonalAction                                                  As ADODB.Recordset
Dim rsDepartment                                                      As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim EMPLIVIL                                                          As String

Sub InitMemVars()
    txtFrom.Text = ""
    txtTo.Text = ""
    cboActionDesc.Text = ""
    txtPosition.Text = ""
    cboDepartment.Text = ""
    txtAnnBasicSalary.Text = ""
End Sub

Sub StoreEntry(XXX As Variant)
    Set rsPersonalAction = New ADODB.Recordset
    Set rsPersonalAction = gconDMIS.Execute("Select * from HRMS_PersonalAction Where ID = " & XXX)
    If Not rsPersonalAction.EOF And Not rsPersonalAction.BOF Then
        labID.Caption = rsPersonalAction!ID
        txtFrom.Text = Null2String(rsPersonalAction!From)
        txtTo.Text = Null2String(rsPersonalAction!To)
        cboActionDesc.Text = Null2String(rsPersonalAction!ActionDesc)
        txtPosition.Text = Null2String(rsPersonalAction!Position)
        cboDepartment.Text = Null2String(rsPersonalAction!Department)
        txtAnnBasicSalary.Text = Null2String(rsPersonalAction!AnnBasicSalary)
    End If
End Sub

Sub FillGrid()
    lstPersonalAction.Sorted = False: lstPersonalAction.ListItems.Clear
    lstPersonalAction.Enabled = False
    Set rsPersonalAction = New ADODB.Recordset
    Set rsPersonalAction = gconDMIS.Execute("select [From],[To],ActionDesc,[Position],Department,AnnBasicSalary,ID from HRMS_PersonalAction where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsPersonalAction.EOF And rsPersonalAction.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstPersonalAction.ListItems, rsPersonalAction
        lstPersonalAction.Refresh
        lstPersonalAction.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else
        EmptyRecord = True
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

'Upating Code       : AXP-0707200712:01
Private Sub cmdAdd_Click()

    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "ACESS_ADD", "DATA ENTRY") = False Then Exit Sub
    cmdPersonalAction.ZOrder 0: picPersonalAction.ZOrder 0
    AddorEdit = "ADD"
    InitMemVars
    On Error Resume Next
    txtFrom.SetFocus





    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200712:00
Private Sub cmdCancel_Click()
    On Error GoTo Errorcode:

    cmdPersonalAction.ZOrder 1: picPersonalAction.ZOrder 1





    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200712:01
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_DELETE", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstPersonalAction.SelectedItem.SubItems(6) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_PersonalAction Where ID = " & lstPersonalAction.SelectedItem.SubItems(6))

                Call LogAudit("X", "DELETE EMPLOYEE PERSONAL ACTION", EMPLOYEE_NO)
                Call ShowDeletedMsg
                Call FillGrid
            End If
        End If
    End If

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

'Upating Code       : AXP-0707200712:01
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_EDIT", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstPersonalAction.SelectedItem.SubItems(6) <> "" Then
            StoreEntry lstPersonalAction.SelectedItem.SubItems(6)
            cmdPersonalAction.ZOrder 0: picPersonalAction.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If





    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200712:01
Private Sub cmdExit_Click()
    On Error GoTo Errorcode:

    Unload Me





    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200712:00
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    cmdPersonalAction.ZOrder 1: picPersonalAction.ZOrder 1
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_PersonalAction " & _
                         "(EMPLEVEL,EMPNO,ACTIONDESC,[Position],[From],[To],Department,AnnBasicSalary,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(cboActionDesc.Text) & "," & N2Str2Null(txtPosition.Text) & "," & N2Str2Null(txtFrom.Text) & "," & N2Str2Null(txtTo.Text) & "," & N2Str2Null(cboDepartment.Text) & "," & N2Str2Null(txtAnnBasicSalary.Text) & ",'" & LOGCODE & "','" & LOGDATE & "')"

        Call LogAudit("A", "ADD EMPLOYEE PERSONAL ACTION", EMPLOYEE_NO)
    Else
        gconDMIS.Execute "update HRMS_PersonalAction set " & _
                       " ACTIONDESC = " & N2Str2Null(cboActionDesc.Text) & "," & _
                       " [Position] = " & N2Str2Null(txtPosition.Text) & "," & _
                       " [From] = " & N2Str2Null(txtFrom.Text) & "," & _
                       " [To] = " & N2Str2Null(txtTo.Text) & "," & _
                       " Department = " & N2Str2Null(cboDepartment.Text) & "," & _
                       " AnnBasicSalary = " & N2Str2Null(txtAnnBasicSalary.Text) & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE PERSONAL ACTION", EMPLOYEE_NO)
    End If
    Call FillGrid

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdPersonalAction.ZOrder 1: picPersonalAction.ZOrder 1
        Case vbKeyF3
            cmdPersonalAction.ZOrder 0: picPersonalAction.ZOrder 0
            AddorEdit = "ADD"
            InitMemVars
            On Error Resume Next
            txtFrom.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstPersonalAction.SelectedItem.SubItems(6) <> "" Then
                    StoreEntry lstPersonalAction.SelectedItem.SubItems(6)
                    cmdPersonalAction.ZOrder 0: picPersonalAction.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstPersonalAction.SelectedItem.SubItems(6) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_PersonalAction Where ID = " & lstPersonalAction.SelectedItem.SubItems(6))
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
    cmdPersonalAction.ZOrder 1: picPersonalAction.ZOrder 1
    cboActionDesc.Clear
    cboActionDesc.AddItem "PROBATIONARY"
    cboActionDesc.AddItem "REGULAR"
    cboActionDesc.AddItem "PROMOTION"
    cboActionDesc.AddItem "PERMANENT"
    cboActionDesc.AddItem "TRANSFER"
    cboActionDesc.AddItem "DEMOTION"
    cboActionDesc.AddItem "TEMPORARY"
    cboActionDesc.AddItem "DISMISSAL"
    cboActionDesc.AddItem "UPGRADING"
    cboActionDesc.AddItem "REEMPLOYMENT"
    cboActionDesc.AddItem "STEP INCRE"
    cboActionDesc.AddItem "INITIAL APPT"
    cboActionDesc.AddItem "CASUAL"
    cboActionDesc.AddItem "DEVOLVED"
    cboActionDesc.AddItem "RESIGNED"
    cboActionDesc.AddItem "DEVOLVED"
    cboActionDesc.AddItem "RA 6758"
    cboActionDesc.AddItem "AnnBasicSalary ADJ."
    cboActionDesc.AddItem "RETROACTIVE"
    cboActionDesc.AddItem "AUDIT REPORT"
    cboActionDesc.AddItem "PCPP IMPL."
    cboActionDesc.AddItem "CONTRACTUAL"
    cboActionDesc.AddItem "PROBATIONAL"
    cboActionDesc.AddItem "C. OF POS."
    cboActionDesc.AddItem "APPOINTIVE"
    cboActionDesc.AddItem "PROM (TEMP)"
    cboActionDesc.AddItem "PROM (PERM)"
    cboActionDesc.AddItem "CSC-BUL-5"
    cboActionDesc.AddItem "EMERGENCY"
    cboActionDesc.AddItem "ELECTIVE"
    cboActionDesc.AddItem "RENEWAL AS P"
    cboActionDesc.AddItem "SECONDMENT"
    cboActionDesc.AddItem "NCC 16"
    cboActionDesc.AddItem "NCC 27"
    cboActionDesc.AddItem "NCC 35"
    cboActionDesc.AddItem "NCC 47"
    cboActionDesc.AddItem "NCC 51"
    cboActionDesc.AddItem "EO 116"
    cboActionDesc.AddItem "NCC 56"
    cboActionDesc.AddItem "DANGAL NG B."
    cboActionDesc.AddItem "MERIT INC."
    cboActionDesc.AddItem "NCC 41"
    cboActionDesc.AddItem "TEM./STIPEND"
    cboActionDesc.AddItem "ELGD.TCH."
    cboActionDesc.AddItem "SUBSTITUTE"
    cboActionDesc.AddItem "BC # 3"
    cboActionDesc.AddItem "SERV. CONTR."
    Set rsDepartment = New ADODB.Recordset
    Set rsDepartment = gconDMIS.Execute("Select * from HRMS_Department Order by DeptName asc")
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        rsDepartment.MoveFirst
        Do While Not rsDepartment.EOF
            cboDepartment.AddItem Null2String(rsDepartment!DeptName)
            rsDepartment.MoveNext
        Loop
    End If
    FillGrid
End Sub

Private Sub lstPersonalAction_DblClick()
    If EmptyRecord = False Then
        If lstPersonalAction.SelectedItem.SubItems(6) <> "" Then
            StoreEntry lstPersonalAction.SelectedItem.SubItems(6)
            cmdPersonalAction.ZOrder 0: picPersonalAction.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub txtFrom_GotFocus()
    txtFrom.Text = Format(txtFrom.Text, "MM/DD/YYYY")
End Sub

Private Sub txtFrom_LostFocus()
    If IsDate(txtFrom.Text) = True Then
        txtFrom.Text = Format(txtFrom.Text, "DD-MMM-YYYY")
    Else
        txtFrom.Text = ""
    End If
End Sub

Private Sub txtTo_GotFocus()
    txtTo.Text = Format(txtTo.Text, "MM/DD/YYYY")
End Sub

Private Sub txtTo_LostFocus()
    If IsDate(txtTo.Text) = True Then
        txtTo.Text = Format(txtTo.Text, "DD-MMM-YYYY")
    Else
        txtTo.Text = ""
    End If
End Sub

Private Sub lstPersonalAction_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPersonalAction
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


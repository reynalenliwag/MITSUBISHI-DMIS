VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOPastEmployment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAST EMPLOYMENT"
   ClientHeight    =   4020
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   6975
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6975
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
      MouseIcon       =   "OTHERINFOPastEmployment.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPastEmployment.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exit Window"
      Top             =   3240
      Width           =   705
   End
   Begin VB.PictureBox picPastEmployment 
      Height          =   2700
      Left            =   990
      ScaleHeight     =   2640
      ScaleWidth      =   4875
      TabIndex        =   6
      Top             =   180
      Width           =   4935
      Begin VB.TextBox txtSalary 
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
         TabIndex        =   4
         Top             =   1500
         Width           =   2625
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
         Left            =   1320
         TabIndex        =   0
         Top             =   60
         Width           =   1455
      End
      Begin VB.TextBox txtAgency 
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
         Top             =   1140
         Width           =   2625
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
         Left            =   1320
         TabIndex        =   2
         Top             =   780
         Width           =   2625
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
         Left            =   1320
         TabIndex        =   1
         Top             =   420
         Width           =   1455
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
         Left            =   4080
         MouseIcon       =   "OTHERINFOPastEmployment.frx":04B8
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOPastEmployment.frx":060A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancel Entry"
         Top             =   1890
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
         Left            =   3390
         MouseIcon       =   "OTHERINFOPastEmployment.frx":0948
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOPastEmployment.frx":0A9A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Save Entry"
         Top             =   1890
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
         Left            =   1410
         TabIndex        =   13
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Salary"
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
         TabIndex        =   11
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Agency"
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
         Top             =   1170
         Width           =   2715
      End
      Begin VB.Label Label3 
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
         TabIndex        =   8
         Top             =   810
         Width           =   975
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
         TabIndex        =   7
         Top             =   450
         Width           =   1215
      End
   End
   Begin wizButton.cmd cmdPastEmployment 
      Height          =   2820
      Left            =   930
      TabIndex        =   9
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4974
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
      MICON           =   "OTHERINFOPastEmployment.frx":0DEA
   End
   Begin MSComctlLib.ListView lstPastEmployment 
      Height          =   3105
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5477
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
      MouseIcon       =   "OTHERINFOPastEmployment.frx":0E06
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FROM DATE"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TO DATE"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "POSITION"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "AGENCY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SALARY"
         Object.Width           =   1587
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
      Left            =   5430
      MouseIcon       =   "OTHERINFOPastEmployment.frx":0F68
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPastEmployment.frx":10BA
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Delete Selected Record"
      Top             =   3240
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
      MouseIcon       =   "OTHERINFOPastEmployment.frx":13E5
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPastEmployment.frx":1537
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Edit Selected Record"
      Top             =   3240
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
      MouseIcon       =   "OTHERINFOPastEmployment.frx":1893
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPastEmployment.frx":19E5
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Add Record"
      Top             =   3240
      Width           =   705
   End
End
Attribute VB_Name = "frmOTHERINFOPastEmployment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsPastEmployment                                                  As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim EMPLIVIL                                                          As String

Sub InitMemVars()
    txtFrom.Text = ""
    txtTo.Text = ""
    txtPosition.Text = ""
    txtAgency.Text = ""
    txtSalary.Text = ""
End Sub

Sub StoreEntry(XXX As Variant)
    Set rsPastEmployment = New ADODB.Recordset
    Set rsPastEmployment = gconDMIS.Execute("Select * from HRMS_PastEmployment Where ID = " & XXX)
    If Not rsPastEmployment.EOF And Not rsPastEmployment.BOF Then
        labID.Caption = rsPastEmployment!ID
        txtFrom.Text = Null2String(rsPastEmployment!From)
        txtTo.Text = Null2String(rsPastEmployment!To)
        txtPosition.Text = Null2String(rsPastEmployment!Position)
        txtAgency.Text = Null2String(rsPastEmployment!Agency)
        txtSalary.Text = Null2String(rsPastEmployment!Salary)
    End If
End Sub

Sub FillGrid()
    lstPastEmployment.Sorted = False: lstPastEmployment.ListItems.Clear
    lstPastEmployment.Enabled = False
    Set rsPastEmployment = New ADODB.Recordset
    Set rsPastEmployment = gconDMIS.Execute("select [From],[To],[Position],Agency,Salary,ID from HRMS_PastEmployment where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsPastEmployment.EOF And rsPastEmployment.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstPastEmployment.ListItems, rsPastEmployment
        lstPastEmployment.Refresh
        lstPastEmployment.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else
        EmptyRecord = True
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

'Upating Code       : AXP-0707200711:59
Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Add", "DATA ENTRY") = False Then Exit Sub
    cmdPastEmployment.ZOrder 0: picPastEmployment.ZOrder 0
    AddorEdit = "ADD"
    InitMemVars
    On Error Resume Next
    txtFrom.SetFocus

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    cmdPastEmployment.ZOrder 1: picPastEmployment.ZOrder 1
End Sub

'Upating Code       : AXP-0707200711:59
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Delete", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstPastEmployment.SelectedItem.SubItems(5) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_PastEmployment Where ID = " & lstPastEmployment.SelectedItem.SubItems(5))

                Call LogAudit("X", "DELETE EMPLOYEE PAST EMPLOYMENT INFORMATION", EMPLOYEE_NO)
                Call ShowDeletedMsg
                Call FillGrid
            End If
        End If
    End If

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

'Upating Code       : AXP-0707200711:59
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Edit", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstPastEmployment.SelectedItem.SubItems(5) <> "" Then
            StoreEntry lstPastEmployment.SelectedItem.SubItems(5)
            cmdPastEmployment.ZOrder 0: picPastEmployment.ZOrder 0
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

Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    cmdPastEmployment.ZOrder 1: picPastEmployment.ZOrder 1
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_PastEmployment " & _
                         "(EMPLEVEL,EMPNO,[Position],[From],[To],Agency,Salary,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(txtPosition.Text) & "," & N2Str2Null(txtFrom.Text) & "," & N2Str2Null(txtTo.Text) & "," & N2Str2Null(txtAgency.Text) & "," & N2Str2Null(txtSalary.Text) & ",'" & LOGCODE & "','" & LOGDATE & "')"

        Call LogAudit("A", "ADD EMPLOYEE PAST EMPLOYEMENT", EMPLOYEE_NO)
    Else
        gconDMIS.Execute "update HRMS_PastEmployment set " & _
                       " [Position] = " & N2Str2Null(txtPosition.Text) & "," & _
                       " [From] = " & N2Str2Null(txtFrom.Text) & "," & _
                       " [To] = " & N2Str2Null(txtTo.Text) & "," & _
                       " Agency = " & N2Str2Null(txtAgency.Text) & "," & _
                       " Salary = " & N2Str2Null(txtSalary.Text) & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE PAST EMPLOYMENT INFORMATION", EMPLOYEE_NO)
    End If
    Call FillGrid

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdPastEmployment.ZOrder 1: picPastEmployment.ZOrder 1
        Case vbKeyF3
            cmdPastEmployment.ZOrder 0: picPastEmployment.ZOrder 0
            AddorEdit = "ADD"
            InitMemVars
            On Error Resume Next
            txtFrom.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstPastEmployment.SelectedItem.SubItems(5) <> "" Then
                    StoreEntry lstPastEmployment.SelectedItem.SubItems(5)
                    cmdPastEmployment.ZOrder 0: picPastEmployment.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstPastEmployment.SelectedItem.SubItems(5) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_PastEmployment Where ID = " & lstPastEmployment.SelectedItem.SubItems(5))
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
    cmdPastEmployment.ZOrder 1: picPastEmployment.ZOrder 1
    FillGrid
End Sub

Private Sub lstPastEmployment_DblClick()
    If EmptyRecord = False Then
        If lstPastEmployment.SelectedItem.SubItems(5) <> "" Then
            StoreEntry lstPastEmployment.SelectedItem.SubItems(5)
            cmdPastEmployment.ZOrder 0: picPastEmployment.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub lstPastEmployment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPastEmployment
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


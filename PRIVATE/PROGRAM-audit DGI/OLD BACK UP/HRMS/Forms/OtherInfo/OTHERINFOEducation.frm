VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOEducation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDUCATION PROFILE"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picEducation 
      Height          =   2265
      Left            =   1050
      ScaleHeight     =   2205
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   210
      Width           =   4935
      Begin VB.TextBox txtYearLastAttend 
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
         Left            =   3555
         TabIndex        =   3
         Top             =   1170
         Width           =   1155
      End
      Begin VB.TextBox txtSchool 
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
         Top             =   780
         Width           =   3645
      End
      Begin VB.ComboBox cboDegree 
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
         TabIndex        =   0
         Text            =   "cboEmpStatus"
         Top             =   60
         Width           =   3615
      End
      Begin VB.TextBox txtMajor 
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
         Width           =   3645
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
         Left            =   2310
         MouseIcon       =   "OTHERINFOEducation.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOEducation.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cancel Entry"
         Top             =   1530
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
         Left            =   1620
         MouseIcon       =   "OTHERINFOEducation.frx":0490
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOEducation.frx":05E2
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Save Entry"
         Top             =   1530
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
         Left            =   1140
         TabIndex        =   11
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year of Last Attendance"
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
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Degree"
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
         Width           =   1125
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "School"
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
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Major"
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
         Top             =   450
         Width           =   1215
      End
   End
   Begin wizButton.cmd cmdEducation 
      Height          =   2385
      Left            =   990
      TabIndex        =   8
      Top             =   150
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4207
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
      MICON           =   "OTHERINFOEducation.frx":0932
   End
   Begin MSComctlLib.ListView lstEducation 
      Height          =   2565
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   4524
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
      MouseIcon       =   "OTHERINFOEducation.frx":094E
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "DEGREE"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MAJOR"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SCHOOL"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "YEAR"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID"
         Object.Width           =   2
      EndProperty
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
      Height          =   705
      Left            =   6090
      MouseIcon       =   "OTHERINFOEducation.frx":0AB0
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOEducation.frx":0C02
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit Window"
      Top             =   2670
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
      Left            =   5400
      MouseIcon       =   "OTHERINFOEducation.frx":0F68
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOEducation.frx":10BA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Delete Selected Record"
      Top             =   2670
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
      Left            =   4710
      MouseIcon       =   "OTHERINFOEducation.frx":13E5
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOEducation.frx":1537
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Edit Selected Record"
      Top             =   2670
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
      Left            =   4020
      MouseIcon       =   "OTHERINFOEducation.frx":1893
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOEducation.frx":19E5
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Add Record"
      Top             =   2670
      Width           =   705
   End
End
Attribute VB_Name = "frmOTHERINFOEducation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsEducation                                                       As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim EMPLIVIL                                                          As String

Sub InitMemVars()
    cboDEGREE.Text = ""
    txtMajor.Text = ""
    txtSchool.Text = ""
    txtYearLastAttend.Text = ""
End Sub

Sub StoreEntry(XXX As Variant)
    Set rsEducation = New ADODB.Recordset
    Set rsEducation = gconDMIS.Execute("Select * from HRMS_Education Where ID = " & XXX)
    If Not rsEducation.EOF And Not rsEducation.BOF Then
        labID.Caption = rsEducation!ID
        cboDEGREE.Text = Null2String(rsEducation!DEGREE)
        txtMajor.Text = Null2String(rsEducation!Major)
        txtSchool.Text = Null2String(rsEducation!SCHOOL)
        txtYearLastAttend.Text = Null2String(rsEducation!YEARLASTATTEND)
    End If
End Sub

Sub FillGrid()
    lstEducation.Sorted = False: lstEducation.ListItems.Clear
    Set rsEducation = New ADODB.Recordset
    lstEducation.Enabled = False
    Set rsEducation = gconDMIS.Execute("select DEGREE,MAJOR,SCHOOL,YEARLASTATTEND,ID from HRMS_Education where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsEducation.EOF And rsEducation.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstEducation.ListItems, rsEducation
        lstEducation.Refresh
        lstEducation.Enabled = True
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
    cmdEducation.ZOrder 0: picEducation.ZOrder 0
    AddorEdit = "ADD"
    InitMemVars
    On Error Resume Next
    cboDEGREE.SetFocus
End Sub

Private Sub cmdCancel_Click()
    cmdEducation.ZOrder 1: picEducation.ZOrder 1
End Sub

Private Sub cmdDelete_Click()
    'If Function_Access(LOGID, "ACESS_DELETE", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstEducation.SelectedItem.SubItems(4) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_Education Where ID = " & lstEducation.SelectedItem.SubItems(4))

                Call LogAudit("X", "DELETE EMPLOYEE EDUCATION INFORMATION", EMPLOYEE_NO)
                Call ShowDeletedMsg
                Call FillGrid
            End If
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    'If Function_Access(LOGID, "ACESS_EDIT", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstEducation.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstEducation.SelectedItem.SubItems(4)
            cmdEducation.ZOrder 0: picEducation.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:56
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    cmdEducation.ZOrder 1: picEducation.ZOrder 1
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_Education " & _
                         "(EMPLEVEL,EMPNO,DEGREE,MAJOR,SCHOOL,YEARLASTATTEND,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(cboDEGREE.Text) & "," & N2Str2Null(txtMajor.Text) & "," & N2Str2Null(txtSchool.Text) & "," & N2Str2Null(txtYearLastAttend.Text) & ",'" & LOGCODE & "','" & LOGDATE & "')"

        Call LogAudit("A", "ADD EMPLOYEE EDUCATION INFORMATION", EMPLOYEE_NO)
    Else
        gconDMIS.Execute "update HRMS_Education set " & _
                       " DEGREE = " & N2Str2Null(cboDEGREE.Text) & "," & _
                       " MAJOR = " & N2Str2Null(txtMajor.Text) & "," & _
                       " SCHOOL = " & N2Str2Null(txtSchool.Text) & "," & _
                       " YEARLASTATTEND = " & N2Str2Null(txtYearLastAttend.Text) & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE EDUCTION INFORMATION", EMPLOYEE_NO)
    End If
    FillGrid





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdEducation.ZOrder 1: picEducation.ZOrder 1
        Case vbKeyF3
            cmdEducation.ZOrder 0: picEducation.ZOrder 0
            AddorEdit = "ADD"
            InitMemVars
            On Error Resume Next
            cboDEGREE.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstEducation.SelectedItem.SubItems(4) <> "" Then
                    StoreEntry lstEducation.SelectedItem.SubItems(4)
                    cmdEducation.ZOrder 0: picEducation.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstEducation.SelectedItem.SubItems(4) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_Education Where ID = " & lstEducation.SelectedItem.SubItems(4))
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
    cboDEGREE.Clear
    cboDEGREE.AddItem "Ph/D"
    cboDEGREE.AddItem "MS/MA"
    cboDEGREE.AddItem "BS/AB"
    cboDEGREE.AddItem "Vocational"
    cboDEGREE.AddItem "HS"
    cboDEGREE.AddItem "Elem"
    cboDEGREE.AddItem "Others"
    cmdEducation.ZOrder 1: picEducation.ZOrder 1
    FillGrid
End Sub

Private Sub lstEducation_DblClick()
    If EmptyRecord = False Then
        If lstEducation.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstEducation.SelectedItem.SubItems(4)
            cmdEducation.ZOrder 0: picEducation.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub lstEducation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstEducation
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


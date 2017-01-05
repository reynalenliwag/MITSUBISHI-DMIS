VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOExamsPassed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXAMINATION TAKEN"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picExamsPassed 
      Height          =   2400
      Left            =   990
      ScaleHeight     =   2340
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
         Height          =   660
         Left            =   3570
         MouseIcon       =   "OTHERINFOExamsPassed.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOExamsPassed.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancel Entry"
         Top             =   1575
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
         Height          =   660
         Left            =   2880
         MouseIcon       =   "OTHERINFOExamsPassed.frx":0490
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOExamsPassed.frx":05E2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Save Entry"
         Top             =   1575
         Width           =   705
      End
      Begin VB.TextBox txtExamination 
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
         Width           =   3405
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
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1140
         Width           =   2985
      End
      Begin VB.TextBox txtRating 
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
         MaxLength       =   20
         TabIndex        =   2
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtDateTaken 
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
         Left            =   1530
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Examination"
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
         TabIndex        =   9
         Top             =   1170
         Width           =   2715
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Rating"
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
         Caption         =   "Date Taken"
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
   Begin wizButton.cmd cmdExamsPassed 
      Height          =   2520
      Left            =   930
      TabIndex        =   8
      Top             =   150
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4445
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
      MICON           =   "OTHERINFOExamsPassed.frx":0932
   End
   Begin MSComctlLib.ListView lstExamsPassed 
      Height          =   2865
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5054
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
      MouseIcon       =   "OTHERINFOExamsPassed.frx":094E
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EXAMINATION"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DATE TAKEN"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RATING"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PLACE"
         Object.Width           =   2999
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
      Left            =   6180
      MouseIcon       =   "OTHERINFOExamsPassed.frx":0AB0
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOExamsPassed.frx":0C02
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit Window"
      Top             =   3030
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
      Left            =   5490
      MouseIcon       =   "OTHERINFOExamsPassed.frx":0F68
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOExamsPassed.frx":10BA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Delete Selected Record"
      Top             =   3030
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
      Left            =   4800
      MouseIcon       =   "OTHERINFOExamsPassed.frx":13E5
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOExamsPassed.frx":1537
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Edit Selected Record"
      Top             =   3030
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
      Left            =   4110
      MouseIcon       =   "OTHERINFOExamsPassed.frx":1893
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOExamsPassed.frx":19E5
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Add Record"
      Top             =   3030
      Width           =   705
   End
End
Attribute VB_Name = "frmOTHERINFOExamsPassed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsEXAMSPASSED                                                     As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim EMPLIVIL                                                          As String

Sub InitMemVars()
    txtExamination.Text = ""
    txtDateTaken.Text = ""
    txtRating.Text = ""
    txtPlace.Text = ""
End Sub

Sub StoreEntry(XXX As Variant)
    Set rsEXAMSPASSED = New ADODB.Recordset
    Set rsEXAMSPASSED = gconDMIS.Execute("Select * from HRMS_ExamsPassed Where ID = " & XXX)
    If Not rsEXAMSPASSED.EOF And Not rsEXAMSPASSED.BOF Then
        labID.Caption = rsEXAMSPASSED!ID
        txtExamination.Text = Null2String(rsEXAMSPASSED!Examination)
        txtDateTaken.Text = Null2String(rsEXAMSPASSED!DateTaken)
        txtRating.Text = Null2String(rsEXAMSPASSED!Rating)
        txtPlace.Text = Null2String(rsEXAMSPASSED!Place)
    End If
End Sub

Sub FillGrid()
    lstExamsPassed.Sorted = False: lstExamsPassed.ListItems.Clear
    lstExamsPassed.Enabled = False
    Set rsEXAMSPASSED = New ADODB.Recordset
    Set rsEXAMSPASSED = gconDMIS.Execute("select Examination,DateTaken,Rating,Place,ID from HRMS_ExamsPassed where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsEXAMSPASSED.EOF And rsEXAMSPASSED.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstExamsPassed.ListItems, rsEXAMSPASSED
        lstExamsPassed.Refresh
        lstExamsPassed.Enabled = True
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
    cmdExamsPassed.ZOrder 0: picExamsPassed.ZOrder 0
    AddorEdit = "ADD"
    InitMemVars
    On Error Resume Next
    txtExamination.SetFocus
End Sub

Private Sub cmdCancel_Click()
    cmdExamsPassed.ZOrder 1: picExamsPassed.ZOrder 1
End Sub

'Upating Code       : AXP-0707200711:57
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_DELETE", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstExamsPassed.SelectedItem.SubItems(4) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_ExamsPassed Where ID = " & lstExamsPassed.SelectedItem.SubItems(4))

                Call LogAudit("X", "DELETE EMPLOYEE EXAM TAKEN", EMPLOYEE_NO)
                Call ShowDeletedMsg
                Call FillGrid
            End If
        End If
    End If

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

'Upating Code       : AXP-0707200711:56
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_EDIT", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstExamsPassed.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstExamsPassed.SelectedItem.SubItems(4)
            cmdExamsPassed.ZOrder 0: picExamsPassed.ZOrder 0
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

'Upating Code       : AXP-0707200711:56
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    cmdExamsPassed.ZOrder 1: picExamsPassed.ZOrder 1
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_ExamsPassed " & _
                         "(EMPLEVEL,EMPNO,Examination,DateTaken,Rating,Place,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(txtExamination.Text) & "," & N2Str2Null(txtDateTaken.Text) & "," & N2Str2Null(txtRating.Text) & "," & N2Str2Null(txtPlace.Text) & ",'" & LOGCODE & "','" & LOGDATE & "')"

        Call LogAudit("A", "ADD EMPLOYEE EXAM TAKEN", EMPLOYEE_NO)
    Else
        gconDMIS.Execute "update HRMS_ExamsPassed set " & _
                       " Examination = " & N2Str2Null(txtExamination.Text) & "," & _
                       " DateTaken = " & N2Str2Null(txtDateTaken.Text) & "," & _
                       " Rating = " & N2Str2Null(txtRating.Text) & "," & _
                       " Place = " & N2Str2Null(txtPlace.Text) & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE EXAM TAKEN", EMPLOYEE_NO)
    End If
    Call FillGrid

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdExamsPassed.ZOrder 1: picExamsPassed.ZOrder 1
        Case vbKeyF3
            cmdExamsPassed.ZOrder 0: picExamsPassed.ZOrder 0
            AddorEdit = "ADD"
            InitMemVars
            On Error Resume Next
            txtExamination.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstExamsPassed.SelectedItem.SubItems(4) <> "" Then
                    StoreEntry lstExamsPassed.SelectedItem.SubItems(4)
                    cmdExamsPassed.ZOrder 0: picExamsPassed.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstExamsPassed.SelectedItem.SubItems(4) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_ExamsPassed Where ID = " & lstExamsPassed.SelectedItem.SubItems(4))
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
    cmdExamsPassed.ZOrder 1: picExamsPassed.ZOrder 1
    FillGrid
End Sub

Private Sub lstExamsPassed_DblClick()
    If EmptyRecord = False Then
        If lstExamsPassed.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstExamsPassed.SelectedItem.SubItems(4)
            cmdExamsPassed.ZOrder 0: picExamsPassed.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub txtDateTaken_GotFocus()
    txtDateTaken.Text = Format(txtDateTaken.Text, "MM/DD/YYYY")
End Sub

Private Sub txtDateTaken_LostFocus()
    If IsDate(txtDateTaken.Text) = True Then
        txtDateTaken.Text = Format(txtDateTaken.Text, "DD-MMM-YYYY")
    Else
        txtDateTaken.Text = ""
    End If
End Sub

Private Sub lstExamsPassed_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstExamsPassed
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


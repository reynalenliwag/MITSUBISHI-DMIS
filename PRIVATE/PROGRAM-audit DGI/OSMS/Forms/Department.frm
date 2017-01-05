VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOSMSFilesDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "Department.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   5655
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   60
      TabIndex        =   5
      Top             =   1080
      Width           =   5505
      Begin VB.OptionButton optName 
         Caption         =   "Department &Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   150
         Width           =   1845
      End
      Begin VB.OptionButton optCode 
         Caption         =   "Department &Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         Top             =   180
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   570
         Width           =   5295
      End
      Begin MSComctlLib.ListView lstDept 
         Height          =   2175
         Left            =   60
         TabIndex        =   10
         Top             =   960
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   3836
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Department.frx":030A
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Department Data Entry"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5535
      Begin VB.TextBox txtDeptName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   660
         Width           =   3945
      End
      Begin VB.TextBox txtDeptCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "XXXXXXXXXX"
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   330
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   90
      ScaleHeight     =   900
      ScaleWidth      =   9705
      TabIndex        =   11
      Top             =   4410
      Width           =   9705
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
         Height          =   855
         Left            =   4680
         MouseIcon       =   "Department.frx":046C
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":05BE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   735
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
         Height          =   855
         Left            =   3960
         MouseIcon       =   "Department.frx":0924
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   735
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
         Height          =   855
         Left            =   3240
         MouseIcon       =   "Department.frx":0DA1
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":0EF3
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   735
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
         Height          =   855
         Left            =   2400
         MouseIcon       =   "Department.frx":124F
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":13A1
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1680
         MouseIcon       =   "Department.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   960
         MouseIcon       =   "Department.frx":1B00
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":1C52
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         MouseIcon       =   "Department.frx":1FAA
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":20FC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4185
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   19
      Top             =   4410
      Width           =   2580
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
         Height          =   795
         Left            =   720
         MouseIcon       =   "Department.frx":245B
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":25AD
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   60
         Width           =   675
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
         Height          =   795
         Left            =   60
         MouseIcon       =   "Department.frx":28EB
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":2A3D
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmOSMSFilesDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDepartment As ADODB.Recordset
Dim AddorEdit As String
Dim PrevDCODE As String

Private Sub cmdAdd_Click()
    Frame1.Caption = "Add A Record"
    AddorEdit = "ADD"
    Picture1.Visible = False
    Picture2.Visible = True
    Frame1.Enabled = True
    initMemvars
    On Error Resume Next
    txtDeptCode.SetFocus
End Sub

Sub initMemvars()
    txtDeptCode.Text = ""
    txtDeptName.Text = ""
End Sub

Private Sub cmdCancel_Click()
    Frame1.Caption = "Department Data Entry"
    AddorEdit = ""
    Picture1.Visible = True
    Picture2.Visible = False
    Frame1.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "delete  from  OSMS_department where Department_Code = '" & txtDeptCode.Text & "'"
        rsRefresh
        StoreMemVars
    End If
End Sub

Private Sub cmdEdit_Click()
    Frame1.Caption = "Edit Record"
    Frame1.Enabled = True
    AddorEdit = "EDIT"
    PrevDCODE = txtDeptCode.Text
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtDeptCode.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    txtSearch.SetFocus
End Sub

Function RecordFound(AAA As Variant) As Boolean
    Dim rsRecordFound As ADODB.Recordset
    Set rsRecordFound = New ADODB.Recordset
    Set rsRecordFound = rsDepartment.Clone
    rsRecordFound.Find "Dept_Description = '" & AAA & "'"
    If Not rsRecordFound.EOF Then
        rsDepartment.Bookmark = rsRecordFound.Bookmark
        RecordFound = True
    Else
        Set rsRecordFound = New ADODB.Recordset
        Set rsRecordFound = rsDepartment.Clone
        rsRecordFound.Find "Department_Code = '" & AAA & "'"
        If Not rsRecordFound.EOF Then
            rsDepartment.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            RecordFound = False
        End If
    End If
End Function

Private Sub cmdNext_Click()
    On Error Resume Next
    rsDepartment.MoveNext
    If rsDepartment.EOF Then
        MsgBoxXP "Last of Record!", "Last Record", XP_OKOnly, msg_Information
        rsDepartment.MoveLast
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsDepartment.MovePrevious
    If rsDepartment.BOF Then
        ShowFirstRecordMsg
        rsDepartment.MoveFirst
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    Screen.MousePointer = 11
    If txtDeptCode.Text = "" Then
        Screen.MousePointer = 0
        MsgBoxXP "Department Code must not be empty!", "Invalid Department Code", XP_OKOnly, msg_Exclamation
        On Error Resume Next
        txtDeptCode.SetFocus
        Exit Sub
    End If

    If txtDeptName.Text = "" Then
        Screen.MousePointer = 0
        MsgBoxXP "Department Name must not be empty!", "Invalid Department Name", XP_OKOnly, msg_Exclamation
        On Error Resume Next
        txtDeptCode.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        rsDepartment.Find " Department_Code = '" & txtDeptCode & "'"
        If Not rsDepartment.EOF Then
            Screen.MousePointer = 0
            MsgBoxXP "Department Code already exists!", "Invalid Department Code", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtDeptCode.SetFocus
            Exit Sub
        End If
        gconDMIS.Execute "insert into  OSMS_department " & _
                         "(department_code, dept_description)" & _
                       " values ('" & txtDeptCode.Text & "','" & txtDeptName.Text & "')"
    Else
        If UCase(PrevDCODE) <> UCase(txtDeptCode.Text) Then
            rsDepartment.Find " Department_Code = '" & txtDeptCode & "'"
            If Not rsDepartment.EOF Then
                Screen.MousePointer = 0
                MsgBoxXP "Department Code already exists!", "Invalid Department Code", XP_OKOnly, msg_Exclamation
                On Error Resume Next
                txtDeptCode.SetFocus
                Exit Sub
            End If
        End If
        gconDMIS.Execute "update OSMS_department set " & _
                         "department_code = '" & txtDeptCode.Text & "'," & _
                         "dept_description = '" & txtDeptName.Text & "'" & _
                         "where department_code = '" & PrevDCODE & "'"
    End If
    rsRefresh
    cmdCancel.Value = True
    Screen.MousePointer = 0
    Exit Sub

ErrorHandler:
    Screen.MousePointer = 0
    'MsgBoxXP "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description, "Error Encountered", XP_OKOnly, msg_Critical
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    rsRefresh
    txtSearch.Text = ""
    StoreMemVars

    Call AddColumnHeader("DepartmentCode, DepartmentName", lstDept)
    Call ResizeColumnHeader(lstDept, "30,63")

End Sub

Sub rsRefresh()
    Set rsDepartment = New ADODB.Recordset
    rsDepartment.Open "select * from  OSMS_department order by department_code asc", gconDMIS
End Sub

Sub StoreMemVars()
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        txtDeptCode.Text = rsDepartment!DEPARTMENT_CODE
        txtDeptName.Text = rsDepartment!dept_description
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Exclamation
        cmdAdd.Value = True
    End If
End Sub



Private Sub lstDept_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsDepartment.Bookmark = rsFind(rsDepartment.Clone, "department_Code", lstDept.SelectedItem.Text).Bookmark
    StoreMemVars
End Sub

Private Sub lstDept_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstDept
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstDept_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtSearch_Change()

    FillSearchGrid (txtSearch.Text)

    'If optName.Value = True Then Else End If       FillSearchGrid (txtSearch.Text)

End Sub


Sub FillSearchGrid(XXX As String)
    Dim rsDepartment2 As ADODB.Recordset
    lstDept.Sorted = False
    lstDept.ListItems.Clear
    
    Set rsDepartment2 = New ADODB.Recordset

    lstDept.Enabled = False

    If optName.Value = True Then
        Set rsDepartment2 = gconDMIS.Execute("select department_code,  DEPT_DESCRIPTION from  OSMS_department where DEPT_DESCRIPTION like'" & XXX & "%' order by DEPT_DESCRIPTION asc")
    Else
        Set rsDepartment2 = gconDMIS.Execute("select department_code,  DEPT_DESCRIPTION from  OSMS_department where department_code like'" & XXX & "%' order by department_code asc")
    End If

    If Not (rsDepartment2.EOF And rsDepartment2.BOF) Then
        Listview_Loadval Me.lstDept.ListItems, rsDepartment2
        lstDept.Refresh
        lstDept.Enabled = True
    End If
End Sub

Private Sub optCode_Click()
    FillSearchGrid (txtSearch.Text)
End Sub
Private Sub optName_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

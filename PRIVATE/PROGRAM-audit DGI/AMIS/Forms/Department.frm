VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmAMISFILESDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department Data Entry"
   ClientHeight    =   5490
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   5790
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Department.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   5790
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   5940
      TabIndex        =   9
      Top             =   4650
      Width           =   5940
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
         Height          =   795
         Left            =   4860
         MouseIcon       =   "Department.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
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
         Left            =   4170
         MouseIcon       =   "Department.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Print this Record"
         Top             =   30
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
         Height          =   795
         Left            =   3480
         MouseIcon       =   "Department.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
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
         Height          =   795
         Left            =   2790
         MouseIcon       =   "Department.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Height          =   795
         Left            =   2100
         MouseIcon       =   "Department.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
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
         Height          =   795
         Left            =   1410
         MouseIcon       =   "Department.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
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
         Height          =   795
         Left            =   720
         MouseIcon       =   "Department.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
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
         Height          =   795
         Left            =   30
         MouseIcon       =   "Department.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4275
      ScaleHeight     =   885
      ScaleWidth      =   1620
      TabIndex        =   18
      Top             =   4650
      Width           =   1620
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
         MouseIcon       =   "Department.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancel"
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
         Height          =   795
         Left            =   30
         MouseIcon       =   "Department.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      Begin Crystal.CrystalReport rptDepartment 
         Left            =   5130
         Top             =   90
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtDeptName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   780
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   570
         Width           =   4785
      End
      Begin VB.TextBox txtDeptCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   780
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   30
         TabIndex        =   5
         Top             =   630
         Width           =   675
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         TabIndex        =   3
         Top             =   600
         Width           =   465
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4590
         TabIndex        =   4
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3705
      Left            =   60
      TabIndex        =   7
      Top             =   930
      Width           =   5625
      Begin MSComctlLib.ListView lstDepartment 
         Height          =   3525
         Left            =   30
         TabIndex        =   8
         Top             =   150
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6218
         View            =   3
         LabelEdit       =   1
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
         MouseIcon       =   "Department.frx":36A3
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DEPT CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DEPARTMENT NAME"
            Object.Width           =   6702
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAMISFILESDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDepartment                                            As ADODB.Recordset
Dim AddorEdit                                               As String

Sub rsRefresh()
    Set rsDepartment = New ADODB.Recordset
    rsDepartment.Open "select * from AMIS_Department order by DeptCode asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtDeptCode.Text = ""
    txtDeptName.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        Frame1.Enabled = False
        labID.Caption = rsDepartment!ID
        txtDeptCode.Text = Null2String(rsDepartment!DeptCode)
        txtDeptName.Text = Null2String(rsDepartment!DeptName)
    Else
        lstDepartment.ListItems.Clear
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsDepartment2                                       As ADODB.Recordset
    Set rsDepartment2 = New ADODB.Recordset
    rsDepartment2.Open "select * from AMIS_Department where ID = " & XXX, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsDepartment2.EOF And Not rsDepartment2.BOF Then
        fraDetails.Enabled = False
        lstDepartment.Enabled = False
        labID.Caption = rsDepartment2!ID
        txtDeptCode.Text = Null2String(rsDepartment2!DeptCode)
        txtDeptName.Text = Null2String(rsDepartment2!DeptName)
    End If
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "DEPARTMENT CODES") = False Then Exit Sub
    AddorEdit = "ADD"
    initMemvars
    Picture1.Visible = False
    Picture2.Visible = True
    lstDepartment.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstDepartment.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
    FillGrid
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "DEPARTMENT CODES") = False Then Exit Sub
    If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
        gconDMIS.Execute "delete from AMIS_Department where ID = " & lstDepartment.SelectedItem.SubItems(2)
        LogAudit "X", "DEPARTMENT MASTER FILE", txtDeptName
    End If
    rsRefresh
    StoreMemVars
    FillGrid
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "DEPARTMENT CODES") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    StoreEntry (lstDepartment.SelectedItem.SubItems(2))
    lstDepartment.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                                             As String
    findStr = InputBox("Please Input Department ...", "Find")
    If findStr <> "" Then
        On Error Resume Next
        rsDepartment.Bookmark = rsFind(rsDepartment.Clone, "DeptCode", findStr).Bookmark
        If Err.Number = 3021 Then
            On Error GoTo ErrorCode
            rsDepartment.Bookmark = rsFind(rsDepartment.Clone, "DeptName", findStr).Bookmark
        End If
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.Number = 3021 Then
        MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
        Resume Next
    End If
End Sub

Private Sub cmdNext_Click()
    rsDepartment.MoveNext
    If rsDepartment.EOF Then
        rsDepartment.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsDepartment.MovePrevious
    If rsDepartment.BOF Then
        rsDepartment.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200713:05
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "DEPARTMENT CODES") = False Then Exit Sub
    Screen.MousePointer = 11

    rptDepartment.Reset
    rptDepartment.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptDepartment.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptDepartment.ReportTitle = "Department Code"


    PrintSQLReport rptDepartment, AMIS_REPORT_PATH & "AccountFiles\Department.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "DEPARTMENT MASTER FILE", txtDeptName
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:05
Private Sub cmdSave_Click()
    Dim VtxtDeptCode, VtxtDeptName                          As String

    On Error GoTo ErrorCode

    VtxtDeptCode = N2Str2Null(txtDeptCode.Text)
    VtxtDeptName = N2Str2Null(txtDeptName.Text)

    If AddorEdit = "ADD" Then
        Dim rsDepartmentDup                                 As ADODB.Recordset
        Set rsDepartmentDup = New ADODB.Recordset
        rsDepartmentDup.Open "select DeptCode from AMIS_Department where DeptCode = " & VtxtDeptCode, gconDMIS
        If Not rsDepartmentDup.EOF And Not rsDepartmentDup.BOF Then
            MsgBox "Department Code Already Exist!", vbCritical, "Duplicate Bank Code Not Allowed"
            Exit Sub
        End If
        gconDMIS.Execute "Insert into AMIS_Department " & _
                         "(DeptCode,DeptName) " & _
                         " values (" & VtxtDeptCode & _
                         ", " & VtxtDeptName & ")"
        LogAudit "A", "DEPARTMENT MASTER FILE", txtDeptName
    Else
        gconDMIS.Execute "update AMIS_Department set" & _
                         " DeptCode = " & VtxtDeptCode & ", " & _
                         " DeptName = " & VtxtDeptName & _
                         " where ID = " & labID.Caption
        LogAudit "E", "DEPARTMENT MASTER FILE", txtDeptName
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsDepartment.Find "ID = " & labID.Caption
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    initMemvars
    StoreMemVars
    FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub FillGrid()
    Dim rsDepartment2                                       As ADODB.Recordset
    lstDepartment.Enabled = False
    lstDepartment.Sorted = False: lstDepartment.ListItems.Clear
    Set rsDepartment2 = New ADODB.Recordset
    Set rsDepartment2 = gconDMIS.Execute("select DeptCode,DeptName,ID from AMIS_Department ORDER BY DEPTCODE ASC")
    If Not (rsDepartment2.EOF And rsDepartment2.BOF) Then
        lstDepartment.Enabled = True
        Listview_Loadval Me.lstDepartment.ListItems, rsDepartment2
        lstDepartment.Refresh
        lstDepartment.Enabled = True
    Else
        lstDepartment.Enabled = False
    End If

End Sub

Private Sub lstDepartment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstDepartment
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

Private Sub lstDepartment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstDepartment_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsDepartment.Bookmark = rsFind(rsDepartment.Clone, "deptcode", Me.lstDepartment.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub txtDeptCode_LostFocus()
    txtDeptCode.Text = UCase(txtDeptCode.Text)
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department Codes"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7935
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Department.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   7935
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2250
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   12
      Top             =   3825
      Width           =   5580
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
         MouseIcon       =   "Department.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":0594
         Style           =   1  'Graphical
         TabIndex        =   20
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
         MouseIcon       =   "Department.frx":08FA
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":0A4C
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "Department.frx":0DB2
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":0F04
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "Department.frx":122F
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":1381
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "Department.frx":16DD
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":182F
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "Department.frx":1B42
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":1C94
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "Department.frx":1F8E
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":20E0
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "Department.frx":2438
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":258A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   4560
      Left            =   30
      ScaleHeight     =   4500
      ScaleWidth      =   1845
      TabIndex        =   11
      Top             =   135
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6960
         Left            =   0
         Picture         =   "Department.frx":28E9
         Top             =   0
         Width           =   9915
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2010
      ScaleHeight     =   2295
      ScaleWidth      =   5865
      TabIndex        =   9
      Top             =   1485
      Width           =   5865
      Begin MSComctlLib.ListView lstDepartment 
         Height          =   2175
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
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
         MouseIcon       =   "Department.frx":16646
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2117
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
   Begin VB.PictureBox picDepartment 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   2010
      ScaleHeight     =   1305
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   135
      Width           =   5865
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   750
         TabIndex        =   5
         Top             =   60
         Width           =   1155
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   750
         TabIndex        =   4
         Top             =   840
         Width           =   5025
      End
      Begin VB.TextBox txtInitials 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   750
         TabIndex        =   3
         Top             =   450
         Width           =   1155
      End
      Begin Crystal.CrystalReport rptDepartment 
         Left            =   5310
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
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
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   675
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
         Left            =   -30
         TabIndex        =   7
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Initials"
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
         Left            =   0
         TabIndex        =   6
         Top             =   510
         Width           =   675
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
         Left            =   4320
         TabIndex        =   2
         Top             =   900
         Width           =   225
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
         Left            =   3780
         TabIndex        =   1
         Top             =   900
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6390
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   21
      Top             =   3825
      Width           =   1440
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
         MouseIcon       =   "Department.frx":167A8
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":168FA
         Style           =   1  'Graphical
         TabIndex        =   23
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
         MouseIcon       =   "Department.frx":16C38
         MousePointer    =   99  'Custom
         Picture         =   "Department.frx":16D8A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDepartment                                                      As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub rsrefresh()
    Set rsDepartment = New ADODB.Recordset
    rsDepartment.Open "select * from HRMS_Department order by deptcode", gconDMIS, adOpenKeyset
End Sub

Sub InitMemvars()
    picDepartment.Enabled = True
    txtDeptCode.Text = ""
    txtDeptName.Text = ""
    txtInitials.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        picDepartment.Enabled = False
        labID.Caption = rsDepartment!ID
        txtDeptCode.Text = Null2String(rsDepartment!Deptcode)
        txtDeptName.Text = Null2String(rsDepartment!DeptName)
        txtInitials.Text = Null2String(rsDepartment!initials)
    Else
        picDepartment.Enabled = False
        ShowNoRecord
    End If
End Sub

Sub FillGrid()
    Dim rsDepartment2                                                 As ADODB.Recordset
    lstDepartment.Enabled = False
    lstDepartment.Sorted = False: lstDepartment.ListItems.Clear
    Set rsDepartment2 = New ADODB.Recordset
    Set rsDepartment2 = gconDMIS.Execute("select DeptCode,DeptName,ID from HRMS_Department")
    If Not (rsDepartment2.EOF And rsDepartment2.BOF) Then
        Listview_Loadval Me.lstDepartment.ListItems, rsDepartment2
        lstDepartment.Refresh
        lstDepartment.Enabled = True
    End If

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "FILES DEPARTMENT") = False Then Exit Sub
    AddorEdit = "ADD"
    InitMemvars
    lstDepartment.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdCancel_Click()
    picDepartment.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstDepartment.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "FILES DEPARTMENT") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from HRMS_Department where id = " & labID.Caption
        LogAudit "X", "DELETE DEPARTMENT RECORD", labID.Caption
        ShowDeletedMsg
    End If
    rsrefresh
    StoreMemVars
    FillGrid
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "FILES DEPARTMENT") = False Then Exit Sub
    AddorEdit = "EDIT"
    picDepartment.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstDepartment.Enabled = False
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    MsgBox "Use the List view to find...", vbInformation, "Find"
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

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "FILES DEPARTMENT") = False Then Exit Sub

    LogAudit "V", "PRINT DEPARTMENT RECORD", ""
    Screen.MousePointer = 11
    rptDepartment.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptDepartment.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptDepartment.Formulas(2) = "COMPANYTIN = '" & COMPANY_TIN & "'"
    rptDepartment.Formulas(3) = "PRINTBY = '" & LOGNAME & "'"

    PrintSQLReport rptDepartment, HRMS_REPORT_PATH & "Department list.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    txtDeptCode.Text = N2Str2Null(txtDeptCode.Text)
    txtDeptName.Text = N2Str2Null(txtDeptName.Text)
    txtInitials.Text = N2Str2Null(txtInitials.Text)
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_Department " & _
                         "(Deptcode,deptname,initials) " & _
                       " values (" & txtDeptCode.Text & ", " & _
                         "" & txtDeptName.Text & ", " & txtInitials.Text & ")"

        LogAudit "A", "ADD DEPARTMENT RECORD", txtDeptCode.Text
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "update HRMS_Department set" & _
                       " deptcode = " & txtDeptCode.Text & "," & _
                       " initials = " & txtInitials.Text & "," & _
                       " deptname = " & txtDeptName.Text & _
                       " where id = " & labID.Caption

        LogAudit "E", "UPDATE DEPARTMENT RECORD", txtDeptCode.Text
        ShowSuccessFullyUpdated
    End If

    rsrefresh
    FillGrid
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsrefresh
    StoreMemVars
    FillGrid
    'DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub lstDepartment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstDepartment
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

Private Sub lstDepartment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstDepartment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsDepartment.Bookmark = rsFind(rsDepartment.Clone, "deptcode", Me.lstDepartment.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub txtDeptCode_LostFocus()
    txtDeptCode.Text = UCase(txtDeptCode.Text)
End Sub


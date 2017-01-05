VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOSMSFilesEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5205
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "frmEmployee.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5205
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   30
      TabIndex        =   11
      Top             =   2130
      Width           =   5115
      Begin VB.OptionButton optID 
         Caption         =   "Employee &ID"
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
         Left            =   1380
         TabIndex        =   12
         Top             =   180
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton optName 
         Caption         =   "&Name"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   150
         Width           =   915
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
         Height          =   390
         Left            =   120
         MaxLength       =   35
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   540
         Width           =   4845
      End
      Begin MSComctlLib.ListView lstEmployee 
         Height          =   2205
         Left            =   90
         TabIndex        =   16
         Top             =   960
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   3889
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
         MouseIcon       =   "frmEmployee.frx":030A
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Full Name"
            Object.Width           =   8943
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label6 
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
         Left            =   150
         TabIndex        =   14
         Top             =   180
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Data Entry"
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
      Height          =   2145
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      Begin VB.ComboBox cboDepartment 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1740
         Width           =   3645
      End
      Begin VB.TextBox txtEmpMI 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "X"
         Top             =   1380
         Width           =   225
      End
      Begin VB.TextBox txtEmpFirstName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   1020
         Width           =   3615
      End
      Begin VB.TextBox txtEmpLastName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         TabIndex        =   3
         Top             =   660
         Width           =   3615
      End
      Begin VB.TextBox txtEmpID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "XXXXXXXXX"
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "M.I."
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
         Left            =   240
         TabIndex        =   8
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Firstname"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lastname"
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
         Left            =   240
         TabIndex        =   4
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Emplyee ID"
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
         Left            =   240
         TabIndex        =   2
         Top             =   330
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   150
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   17
      Top             =   5460
      Width           =   9225
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
         Left            =   4320
         MouseIcon       =   "frmEmployee.frx":046C
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":05BE
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   60
         Width           =   675
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
         Left            =   3660
         MouseIcon       =   "frmEmployee.frx":0924
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   675
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
         Left            =   3000
         MouseIcon       =   "frmEmployee.frx":0DA1
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":0EF3
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   60
         Width           =   675
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
         Left            =   2340
         MouseIcon       =   "frmEmployee.frx":124F
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":13A1
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   675
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
         Left            =   1680
         MouseIcon       =   "frmEmployee.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   60
         Width           =   675
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
         Left            =   1020
         MouseIcon       =   "frmEmployee.frx":1B00
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":1C52
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   60
         Width           =   675
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
         Left            =   360
         MouseIcon       =   "frmEmployee.frx":1FAA
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":20FC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   3750
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   25
      Top             =   5460
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
         MouseIcon       =   "frmEmployee.frx":245B
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":25AD
         Style           =   1  'Graphical
         TabIndex        =   26
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
         MouseIcon       =   "frmEmployee.frx":28EB
         MousePointer    =   99  'Custom
         Picture         =   "frmEmployee.frx":2A3D
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   60
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmOSMSFilesEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsemployee As ADODB.Recordset
Dim rsDepartment As ADODB.Recordset
Dim AddorEdit As String
Dim PrevEmpID As String

Private Sub cmdAdd_Click()
    
    Frame1.Caption = "Add A Record"
    Frame1.Enabled = True
    AddorEdit = "ADD"
    Picture1.Visible = False
    initMemvars
    lstEmployee.Enabled = False
    txtSearch.Enabled = False
    On Error Resume Next
    txtEmpID.SetFocus
End Sub

Sub initMemvars()
    txtEmpID.Text = ""
    txtEmpLastName = ""
    txtEmpFirstName = ""
    txtEmpMI = ""
    INITCBODEPT
End Sub

Private Sub cmdCancel_Click()
    Frame1.Caption = "Employee Data Entry"
    AddorEdit = ""
    Picture1.Visible = True
    Frame1.Enabled = False
    lstEmployee.Enabled = True
    txtSearch.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "delete from OSMS_EMPLOYEE  where Employee_ID = '" & txtEmpID.Text & "'"
        rsRefresh
        StoreMemVars
    End If
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    txtSearch.SetFocus
End Sub

Function RecordFound(AAA As Variant) As Boolean
    If AAA <> "" Then
        Dim rsRecordFound As ADODB.Recordset
        Set rsRecordFound = New Recordset
        rsRecordFound.Open "Select lastname + ', ' + firstname + ' ' AS NAME from OSMS_EMPLOYEE  order by Employee_ID asc", gconDMIS
        rsRecordFound.Find "Name like '" & AAA & "%'"
        If Not rsRecordFound.EOF Then
            rsemployee.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            Set rsRecordFound = New Recordset
            rsRecordFound.Open "Select * from OSMS_EMPLOYEE  order by Employee_ID asc", gconDMIS
            rsRecordFound.Find "FirstName like '" & AAA & "%'"
            If Not rsRecordFound.EOF Then
                rsemployee.Bookmark = rsRecordFound.Bookmark
                RecordFound = True
            Else
                Set rsRecordFound = New Recordset
                rsRecordFound.Open "Select * from OSMS_EMPLOYEE  order by Employee_ID asc", gconDMIS
                rsRecordFound.Find "Employee_ID = '" & AAA & "'"
                If Not rsRecordFound.EOF Then
                    rsemployee.Bookmark = rsRecordFound.Bookmark
                    RecordFound = True
                Else
                    RecordFound = False
                End If
            End If
        End If
    End If
End Function

Private Sub cmdEdit_Click()
    Frame1.Caption = "Edit Record"
    Frame1.Enabled = True
    AddorEdit = "EDIT"
    PrevEmpID = txtEmpID.Text
    Picture1.Visible = False
    On Error Resume Next
    txtEmpID.SetFocus
    lstEmployee.Enabled = False
    fraDetails.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsemployee.MoveNext
    If rsemployee.EOF Then
        MsgBoxXP "Last of Record!", "Last Record", XP_OKOnly, msg_Information
        rsemployee.MoveLast
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsemployee.MovePrevious
    If rsemployee.BOF Then
        ShowFirstRecordMsg
        rsemployee.MoveFirst
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    Screen.MousePointer = 11
    If txtEmpID.Text = "" Then
        MsgBox "Employee ID must not be empty!"
        On Error Resume Next
        txtEmpID.SetFocus
        Exit Sub
    End If

    If txtEmpLastName.Text = "" Then
        MsgBox "Employee Last Name must not be empty!"
        On Error Resume Next
        txtEmpLastName.SetFocus
        Exit Sub
    End If

    If txtEmpFirstName.Text = "" Then
        MsgBox "Employee First Name must not be empty!"
        On Error Resume Next
        txtEmpFirstName.SetFocus
        Exit Sub
    End If

    If txtEmpMI.Text = "" Then
        MsgBox "Employee's Middle Initial must not be empty!"
        On Error Resume Next
        txtEmpMI.SetFocus
        Exit Sub
    End If

    If cboDepartment.Text = "" Then
        Screen.MousePointer = 0
        MsgBoxXP "Department Code must not be empty!", "Invalid "
        On Error Resume Next
        cboDepartment.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        rsemployee.Find " Employee_ID = '" & txtEmpID & "'"
        If Not rsemployee.EOF Then
            Screen.MousePointer = 0
            MsgBoxXP "Employee ID already exists!", "Invalid Employee ID", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtEmpID.SetFocus
            Exit Sub
        End If
        gconDMIS.Execute "INSERT OSMS_employee " & _
                         "(Employee_ID, Lastname, firstname, MI, department_Code) values ('" & txtEmpID.Text & "','" & txtEmpLastName.Text & "','" & txtEmpFirstName.Text & "','" & txtEmpMI.Text & "','" & cboDepartment.Text & "')"

    Else
        If PrevEmpID <> txtEmpID.Text Then
            rsemployee.Find " Employee_ID = '" & txtEmpID & "'"
            If Not rsemployee.EOF Then
                Screen.MousePointer = 0
                MsgBoxXP "Employee ID already exists!", "Invalid Employee ID", XP_OKOnly, msg_Exclamation
                On Error Resume Next
                txtEmpID.SetFocus
                Exit Sub
            End If
        End If
        gconDMIS.Execute "update OSMS_employee set " & _
                         "Employee_ID = '" & txtEmpID.Text & "'," & _
                         "Lastname = '" & txtEmpLastName.Text & "'," & _
                         "Firstname = '" & txtEmpFirstName.Text & "'," & _
                         "MI = '" & txtEmpMI.Text & "'" & _
                         "where Employee_ID = '" & PrevEmpID & "'"
    End If
    rsRefresh
    cmdCancel.Value = True
    Screen.MousePointer = 0
    Exit Sub

ErrorHandler:
    Screen.MousePointer = 0
    'MsgBoxXP "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description
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
    Call AddColumnHeader("EMPID,EMPLOYEENAME", lstEmployee)
    Call ResizeColumnHeader(lstEmployee, "25,68")
End Sub

Sub rsRefresh()
    Set rsemployee = New ADODB.Recordset
    rsemployee.Open "select * from OSMS_EMPLOYEE  order by Employee_ID asc", gconDMIS
End Sub

Sub StoreMemVars()
    If Not rsemployee.EOF And Not rsemployee.BOF Then
        txtEmpID.Text = Null2String(rsemployee!EMPLOYEE_ID)
        txtEmpLastName.Text = Null2String(rsemployee!lastname)
        txtEmpFirstName.Text = Null2String(rsemployee!FirstName)
        txtEmpMI.Text = Null2String(rsemployee!MI)
        cboDepartment.Text = SETCBODEPT2(Null2String(rsemployee!DEPARTMENT_CODE))
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
        cmdAdd.Value = True
    End If
End Sub

Sub INITCBODEPT()
    Set rsDepartment = New Recordset
    rsDepartment.Open "Select Dept_description from  OSMS_department order by Dept_description asc", gconDMIS
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        rsDepartment.MoveFirst
        cboDepartment.Clear
        Do While Not rsDepartment.EOF
            cboDepartment.AddItem Null2String(rsDepartment!dept_description)
            rsDepartment.MoveNext
        Loop
    End If
End Sub

Function SETCBODEPT(XXX As Variant) As String
    Set rsDepartment = New Recordset
    rsDepartment.Open "Dept_Description, DEPARTMENT_CODE  from  OSMS_department WHERE Dept_Description = '" & XXX & "'", gconDMIS
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        SETCBODEPT = Null2String(rsDepartment!DEPARTMENT_CODE)
    End If
End Function

Function SETCBODEPT2(XXX As Variant) As String
    Set rsDepartment = New Recordset
    rsDepartment.Open "Select Dept_Description,Department_CODE from  OSMS_department WHERE Department_CODE = '" & XXX & "'", gconDMIS
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        SETCBODEPT2 = Null2String(rsDepartment!dept_description)
    End If
End Function


Private Sub lstEmployee_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsemployee.Bookmark = rsFind(rsemployee.Clone, "Employee_ID", lstEmployee.SelectedItem.Text).Bookmark
    StoreMemVars
End Sub

Private Sub lstEmployee_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstEmployee
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

Private Sub lstEmployee_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtSearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub



Sub FillSearchGrid(XXX As String)
    Dim rsemployee As ADODB.Recordset
    lstEmployee.Enabled = False
    lstEmployee.Sorted = False
    lstEmployee.ListItems.Clear
    XXX = Repleys(XXX)
    Set rsemployee = New ADODB.Recordset
    If optID.Value = True Then
        Set rsemployee = gconDMIS.Execute("select Employee_ID, LastName + ',' + FirstName + '.' + MI from OSMS_EMPLOYEE  where Employee_ID like'" & XXX & "%' order by Employee_ID asc")
    Else
        Set rsemployee = gconDMIS.Execute("select Employee_ID, LastName + ',' + FirstName + '.' + MI from OSMS_EMPLOYEE  where LastName like'" & XXX & "%' order by Employee_ID asc")
    End If


    If Not (rsemployee.EOF And rsemployee.BOF) Then
        Listview_Loadval Me.lstEmployee.ListItems, rsemployee
        lstEmployee.Refresh
        lstEmployee.Enabled = True
    End If
     
End Sub


Private Sub optID_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub
Private Sub optName_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub


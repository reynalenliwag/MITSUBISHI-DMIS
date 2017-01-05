VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOSMSFilesSignatories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Signatories"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   ForeColor       =   &H8000000F&
   Icon            =   "frmSignatories.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   5280
   Begin VB.Frame Frame1 
      Caption         =   "Employee Data Entry"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   5115
      Begin VB.ComboBox cboDepartment 
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
         Left            =   1350
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1740
         Width           =   3645
      End
      Begin VB.TextBox txtEmpMI 
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
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "X"
         Top             =   1380
         Width           =   225
      End
      Begin VB.TextBox txtEmpFirstName 
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
         Left            =   1350
         TabIndex        =   2
         Top             =   1020
         Width           =   3615
      End
      Begin VB.TextBox txtEmpLastName 
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
         Left            =   1350
         TabIndex        =   1
         Top             =   660
         Width           =   3615
      End
      Begin VB.TextBox txtEmpID 
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
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   0
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
         Left            =   210
         TabIndex        =   9
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
         Left            =   210
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
         Left            =   210
         TabIndex        =   7
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
         Left            =   210
         TabIndex        =   6
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
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
         Left            =   210
         TabIndex        =   5
         Top             =   330
         Width           =   1215
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   60
      TabIndex        =   11
      Top             =   2160
      Width           =   5115
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
         Height          =   375
         Left            =   120
         MaxLength       =   35
         TabIndex        =   14
         Top             =   510
         Width           =   4845
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
      Begin MSComctlLib.ListView lstSignatories 
         Height          =   2115
         Left            =   90
         TabIndex        =   15
         Top             =   930
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   3731
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
         MouseIcon       =   "frmSignatories.frx":030A
         NumItems        =   0
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
         TabIndex        =   16
         Top             =   180
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   180
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   20
      Top             =   5310
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
         MouseIcon       =   "frmSignatories.frx":046C
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":05BE
         Style           =   1  'Graphical
         TabIndex        =   21
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
         MouseIcon       =   "frmSignatories.frx":0924
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   22
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
         MouseIcon       =   "frmSignatories.frx":0DA1
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":0EF3
         Style           =   1  'Graphical
         TabIndex        =   23
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
         MouseIcon       =   "frmSignatories.frx":124F
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":13A1
         Style           =   1  'Graphical
         TabIndex        =   24
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
         MouseIcon       =   "frmSignatories.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   25
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
         MouseIcon       =   "frmSignatories.frx":1B00
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":1C52
         Style           =   1  'Graphical
         TabIndex        =   26
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
         MouseIcon       =   "frmSignatories.frx":1FAA
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":20FC
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   3795
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   17
      Top             =   5310
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
         MouseIcon       =   "frmSignatories.frx":245B
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":25AD
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "frmSignatories.frx":28EB
         MousePointer    =   99  'Custom
         Picture         =   "frmSignatories.frx":2A3D
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   60
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmOSMSFilesSignatories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSignatories As ADODB.Recordset
Dim rsDepartment As ADODB.Recordset
Dim AddorEdit As String
Dim PrevEmpID As String

Private Sub cmdAdd_Click()
    Frame1.Caption = "Add A Record"
    Frame1.Enabled = True
    AddorEdit = "ADD"
    Picture1.Visible = False
    initMemvars
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
    Frame1.Caption = "Signatories Data Entry"
    AddorEdit = ""
    Picture1.Visible = True
    Frame1.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "delete from OSMS_Signatories  where Signatory_ID = '" & txtEmpID.Text & "'"
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
        rsRecordFound.Open "Select lastname + ', ' + firstname + ' ' AS NAME from OSMS_Signatories  order by Signatory_ID asc", gconDMIS
        rsRecordFound.Find "Name like '" & AAA & "%'"
        If Not rsRecordFound.EOF Then
            rsSignatories.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            Set rsRecordFound = New Recordset
            rsRecordFound.Open "Select * from OSMS_Signatories  order by Signatory_ID asc", gconDMIS
            rsRecordFound.Find "FirstName like '" & AAA & "%'"
            If Not rsRecordFound.EOF Then
                rsSignatories.Bookmark = rsRecordFound.Bookmark
                RecordFound = True
            Else
                Set rsRecordFound = New Recordset
                rsRecordFound.Open "Select * from OSMS_Signatories  order by Signatory_ID asc", gconDMIS
                rsRecordFound.Find "Signatory_ID = '" & AAA & "'"
                If Not rsRecordFound.EOF Then
                    rsSignatories.Bookmark = rsRecordFound.Bookmark
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
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    rsSignatories.MoveNext
    If rsSignatories.EOF Then
        MsgBoxXP "Last of Record!", "Last Record", XP_OKOnly, msg_Information
        rsSignatories.MoveLast
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSignatories.MovePrevious
    If rsSignatories.BOF Then
        ShowFirstRecordMsg
        rsSignatories.MoveFirst
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    Screen.MousePointer = 11
    If txtEmpID.Text = "" Then
        Screen.MousePointer = 0
        MsgBox "Signatories ID must not be empty!"
        On Error Resume Next
        txtEmpID.SetFocus
        Exit Sub
    End If

    If txtEmpLastName.Text = "" Then
        Screen.MousePointer = 0
        MsgBox "Signatories Last Name must not be empty!"
        On Error Resume Next
        txtEmpLastName.SetFocus
        Exit Sub
    End If

    If txtEmpFirstName.Text = "" Then
        Screen.MousePointer = 0
        MsgBox "Signatories First Name must not be empty!"
        On Error Resume Next
        txtEmpFirstName.SetFocus
        Exit Sub
    End If

    If txtEmpMI.Text = "" Then
        Screen.MousePointer = 0
        MsgBox "Signatories's Middle Initial must not be empty!"
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
        rsSignatories.Find " Signatory_ID = '" & txtEmpID & "'"
        If Not rsSignatories.EOF Then
            Screen.MousePointer = 0
            MsgBoxXP "Signatories ID already exists!", "Invalid Signatories ID", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtEmpID.SetFocus
            Exit Sub
        End If
        gconDMIS.Execute "insert into OSMS_Signatories " & _
                         "(Signatory_ID, Lastname, firstname, MI, department_Code) values ('" & txtEmpID.Text & "','" & txtEmpLastName.Text & "','" & txtEmpFirstName.Text & "','" & txtEmpMI.Text & "','" & SETCBODEPT(cboDepartment.Text) & "')"

    Else
        If PrevEmpID <> txtEmpID.Text Then
            rsSignatories.Find " Signatory_ID = '" & txtEmpID & "'"
            If Not rsSignatories.EOF Then
                Screen.MousePointer = 0
                MsgBoxXP "Signatories ID already exists!", "Invalid Signatories ID", XP_OKOnly, msg_Exclamation
                On Error Resume Next
                txtEmpID.SetFocus
                Exit Sub
            End If
        End If
        gconDMIS.Execute "update OSMS_Signatories set " & _
                         "Signatory_ID = '" & txtEmpID.Text & "'," & _
                         "Lastname = '" & txtEmpLastName.Text & "'," & _
                         "Firstname = '" & txtEmpFirstName.Text & "'," & _
                         "MI = '" & txtEmpMI.Text & "'" & _
                         "where Signatory_ID = '" & PrevEmpID & "'"
    End If
    rsRefresh
    cmdCancel.Value = True
    Screen.MousePointer = 0
    Exit Sub

ErrorHandler:
    Screen.MousePointer = 0
    MsgBoxXP "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    rsRefresh
    txtSearch.Text = ""
    StoreMemVars
    Call AddColumnHeader("EmployeeID, Fullname", lstSignatories)
    Call ResizeColumnHeader(lstSignatories, "30,68")
    FillSearchGrid ""
End Sub

Sub rsRefresh()
    Set rsSignatories = New ADODB.Recordset
    rsSignatories.Open "select * from OSMS_Signatories  order by Signatory_ID asc", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        txtEmpID.Text = Null2String(rsSignatories!SIGNATORY_ID)
        txtEmpLastName.Text = Null2String(rsSignatories!lastname)
        txtEmpFirstName.Text = Null2String(rsSignatories!FirstName)
        txtEmpMI.Text = Null2String(rsSignatories!MI)
        cboDepartment.Text = SETCBODEPT2(Null2String(rsSignatories!DEPARTMENT_CODE))
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
        cmdAdd.Value = True
    End If
End Sub

Sub INITCBODEPT()
    Set rsDepartment = New Recordset
    rsDepartment.Open "Select  * from  OSMS_department order by Dept_description asc", gconDMIS
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
    rsDepartment.Open "select Dept_Description, DEPARTMENT_CODE  from  OSMS_department WHERE Dept_Description = '" & XXX & "'", gconDMIS
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        SETCBODEPT = Null2String(rsDepartment!DEPARTMENT_CODE)
    End If
End Function

Function SETCBODEPT2(XXX As Variant) As String
    Set rsDepartment = New Recordset
    rsDepartment.Open "Select Dept_Description,Department_CODE  from  OSMS_department WHERE Department_CODE = '" & XXX & "'", gconDMIS
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        SETCBODEPT2 = Null2String(rsDepartment!dept_description)
    End If
End Function




Private Sub lstSignatories_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Call rsSignatories.Find("SIGNATORY_ID='" & lstSignatories.SelectedItem.Text & "'")
    rsSignatories.Bookmark = rsFind(rsSignatories.Clone, "SIGNATORY_ID", lstSignatories.SelectedItem.Text).Bookmark
    StoreMemVars
End Sub

Private Sub lstSignatories_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstSignatories
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

Private Sub lstSignatories_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtSearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsSignatories As ADODB.Recordset
    lstSignatories.Enabled = False
    lstSignatories.Sorted = False
    lstSignatories.ListItems.Clear

    Set rsSignatories = New ADODB.Recordset
    If optID.Value = True Then
        Set rsSignatories = gconDMIS.Execute("select SIGNATORY_ID, LastName + ',' + FirstName + '.'  + MI from OSMS_Signatories  where SIGNATORY_ID like'" & XXX & "%' order by SIGNATORY_ID asc")
    Else
        Set rsSignatories = gconDMIS.Execute("select SIGNATORY_ID, LastName + ',' + FirstName + '.'  + MI from OSMS_Signatories  where LastName like'" & XXX & "%' order by LastName asc")
    End If


    If Not (rsSignatories.EOF And rsSignatories.BOF) Then
        Listview_Loadval Me.lstSignatories.ListItems, rsSignatories
        lstSignatories.Refresh
        lstSignatories.Enabled = True
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



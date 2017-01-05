VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOSMSFilesUnit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "frmUnit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5220
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   120
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   11
      Top             =   3930
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
         MouseIcon       =   "frmUnit.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   12
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
         MouseIcon       =   "frmUnit.frx":07C2
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   13
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
         MouseIcon       =   "frmUnit.frx":0C3F
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":0D91
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "frmUnit.frx":10ED
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":123F
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "frmUnit.frx":1552
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":16A4
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "frmUnit.frx":199E
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":1AF0
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "frmUnit.frx":1E48
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":1F9A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Unit Data Entry"
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
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   180
      TabIndex        =   2
      Top             =   30
      Width           =   4965
      Begin VB.TextBox txtUnitCode 
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
         Left            =   1380
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtUnitDescription 
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
         Left            =   1380
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   630
         Width           =   1455
      End
   End
   Begin VB.Frame Trans_No 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   180
      TabIndex        =   5
      Top             =   990
      Width           =   4935
      Begin VB.OptionButton optCode 
         Caption         =   "Unit &Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   750
         TabIndex        =   8
         Top             =   420
         Value           =   -1  'True
         Width           =   1575
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
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   720
         Width           =   4755
      End
      Begin MSComctlLib.ListView lstUnit 
         Height          =   1695
         Left            =   60
         TabIndex        =   9
         Top             =   1110
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   2990
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
         MouseIcon       =   "frmUnit.frx":22F9
         NumItems        =   0
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Unit &Description"
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
         Left            =   2370
         TabIndex        =   7
         Top             =   390
         Width           =   1845
      End
      Begin VB.Label Label4 
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
         TabIndex        =   10
         Top             =   180
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   3690
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   19
      Top             =   3930
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
         MouseIcon       =   "frmUnit.frx":245B
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":25AD
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
         MouseIcon       =   "frmUnit.frx":28EB
         MousePointer    =   99  'Custom
         Picture         =   "frmUnit.frx":2A3D
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmOSMSFilesUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUnit As ADODB.Recordset
Dim AddorEdit As String
Dim PrevUCODE As String

Private Sub cmdAdd_Click()
    Frame1.Caption = "Add A Record"
    AddorEdit = "ADD"
    Picture1.Visible = False
    Picture2.Visible = True
    Frame1.Enabled = True
    initMemvars
    On Error Resume Next
    txtUnitCode.SetFocus
End Sub

Sub initMemvars()
    txtUnitCode.Text = ""
    txtUnitDescription.Text = ""
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    txtSearch.SetFocus
End Sub

Function RecordFound(AAA As Variant) As Boolean
    If AAA <> "" Then
        Dim rsRecordFound As ADODB.Recordset
        Set rsRecordFound = New Recordset
        rsRecordFound.Open "Select Unit_Description from OSMS_UNIT order by Unit_Code asc", gconDMIS
        rsRecordFound.Find "Unit_Description like '" & AAA & "%'"
        If Not rsRecordFound.EOF Then
            rsUnit.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            Set rsRecordFound = New Recordset
            rsRecordFound.Open "Select * from OSMS_UNIT order by Unit_Code asc", gconDMIS
            rsRecordFound.Find "Unit_Code = '" & AAA & "'"
            If Not rsRecordFound.EOF Then
                rsUnit.Bookmark = rsRecordFound.Bookmark
                RecordFound = True
            Else
                RecordFound = False
            End If
        End If
    End If
End Function

Private Sub cmdCancel_Click()
    Frame1.Caption = "Unit Data Entry"
    AddorEdit = ""
    Picture1.Visible = True
    Picture2.Visible = False
    Frame1.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "delete from OSMS_UNIT where Unit_Code = '" & txtUnitCode.Text & "'"
        rsRefresh
        StoreMemVars
    End If
End Sub

Private Sub cmdEdit_Click()
    Frame1.Caption = "Edit Record"
    AddorEdit = "EDIT"
    PrevUCODE = txtUnitCode.Text
    Frame1.Enabled = True
    On Error Resume Next
    txtUnitCode.SetFocus
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    rsUnit.MoveNext
    If rsUnit.EOF Then
        ShowLastRecordMsg
        rsUnit.MoveLast
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsUnit.MovePrevious
    If rsUnit.BOF Then
        ShowFirstRecordMsg
        rsUnit.MoveFirst
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
'    On Error GoTo ErrorHandler
    Screen.MousePointer = 11
    If txtUnitCode.Text = "" Then
        MsgBoxXP "Unit Code must not be empty!", "Input Unit Code", XP_OKOnly, msg_Information
        On Error Resume Next
        txtUnitCode.SetFocus
        Exit Sub
    End If

    If txtUnitDescription.Text = "" Then
        MsgBoxXP "Unit Description must not be empty!", "Input Unit Description", XP_OKOnly, msg_Information
        On Error Resume Next
        txtUnitDescription.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        rsUnit.Find " Unit_Code = '" & txtUnitCode & "'"
        If Not rsUnit.EOF Then
            Screen.MousePointer = 0
            MsgBox "Unit Code already exists!", "Unit Code already exists", XP_OKOnly, msg_Information
            On Error Resume Next
            txtUnitCode.SetFocus
            Exit Sub
        End If
        gconDMIS.Execute "insert into OSMS_Unit  " & _
                         "(Unit_code, Unit_description) values ('" & txtUnitCode.Text & "','" & txtUnitDescription.Text & "')"

    Else
        If PrevUCODE <> txtUnitCode.Text Then
            rsUnit.Find " Unit_Code = '" & txtUnitCode & "'"
            If Not rsUnit.EOF Then
                Screen.MousePointer = 0
                MsgBox "Unit Code already exists! Unit Code alredy exists!", vbCritical
                On Error Resume Next
                txtUnitCode.SetFocus
                Exit Sub
            End If
        End If
        gconDMIS.Execute "update OSMS_Unit set " & _
                         "Unit_code = '" & txtUnitCode.Text & "'," & _
                         "Unit_description = '" & txtUnitDescription.Text & "'" & _
                         "where Unit_code = '" & PrevUCODE & "'"
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
    AddColumnHeader "UNIT CODE, UNIT DESCRIPTION", lstUnit
    ResizeColumnHeader lstUnit, "28,65"
    
End Sub

Sub rsRefresh()
    Set rsUnit = New ADODB.Recordset
    rsUnit.Open "select * from OSMS_UNIT order by Unit_code asc", gconDMIS
End Sub

Sub StoreMemVars()
    If Not rsUnit.EOF And Not rsUnit.BOF Then
        txtUnitCode.Text = rsUnit!Unit_Code
        txtUnitDescription.Text = rsUnit!Unit_description
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
        cmdAdd.Value = True
    End If
End Sub


Private Sub lstUnit_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsUnit.Bookmark = rsFind(rsUnit.Clone, "Unit_Code", lstUnit.SelectedItem.Text).Bookmark
    StoreMemVars
End Sub

Private Sub lstUnit_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstUnit
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

Private Sub lstUnit_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtSearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub



Sub FillSearchGrid(XXX As String)
    Dim rsUnit As ADODB.Recordset
        lstUnit.Sorted = False
        lstUnit.Enabled = False
        lstUnit.ListItems.Clear
    Set rsUnit = New ADODB.Recordset
    If optCode.Value = True Then
        Set rsUnit = gconDMIS.Execute("SELECT UNIT_CODE, UNIT_DESCRIPTION FROM OSMS_UNIT WHERE UNIT_CODE LIKE'" & XXX & "%' ORDER BY UNIT_CODE ASC")
    Else
        Set rsUnit = gconDMIS.Execute("SELECT UNIT_CODE, UNIT_DESCRIPTION FROM OSMS_UNIT WHERE UNIT_DESCRIPTION LIKE'" & XXX & "%' ORDER BY UNIT_DESCRIPTION  ASC")
    End If
    
    If Not (rsUnit.EOF And rsUnit.BOF) Then
        Listview_Loadval Me.lstUnit.ListItems, rsUnit
        lstUnit.Refresh
         lstUnit.Enabled = True
    End If
   
End Sub



Private Sub optCode_Click()
   FillSearchGrid (txtSearch.Text)
   On Error Resume Next
    txtSearch.SetFocus
End Sub
Private Sub optDesc_Click()
  FillSearchGrid (txtSearch.Text)
  On Error Resume Next
    txtSearch.SetFocus
End Sub



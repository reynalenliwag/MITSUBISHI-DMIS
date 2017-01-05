VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISFILESAccType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Type"
   ClientHeight    =   5130
   ClientLeft      =   1665
   ClientTop       =   1170
   ClientWidth     =   6045
   ForeColor       =   &H00FFFFFF&
   Icon            =   "AccType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   6045
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      ScaleHeight     =   855
      ScaleWidth      =   5940
      TabIndex        =   10
      Top             =   4140
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
         MouseIcon       =   "AccType.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "AccType.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "AccType.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "AccType.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "AccType.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "AccType.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   13
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
         MouseIcon       =   "AccType.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   12
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
         MouseIcon       =   "AccType.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4500
      ScaleHeight     =   885
      ScaleWidth      =   1620
      TabIndex        =   19
      Top             =   4140
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
         MouseIcon       =   "AccType.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   20
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
         MouseIcon       =   "AccType.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "AccType.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   5865
      Begin VB.TextBox txtDescription 
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
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "XX"
         Top             =   210
         Width           =   405
      End
      Begin Crystal.CrystalReport rptAccType 
         Left            =   5100
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Account Headers"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
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
         Left            =   30
         TabIndex        =   1
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Top             =   600
         Width           =   1125
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
         Left            =   3870
         TabIndex        =   3
         Top             =   570
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
         Left            =   4350
         TabIndex        =   4
         Top             =   570
         Width           =   225
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3225
      Left            =   60
      TabIndex        =   7
      Top             =   900
      Width           =   5865
      Begin VB.TextBox txtSearch 
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
         Left            =   90
         MaxLength       =   35
         TabIndex        =   8
         Top             =   150
         Width           =   5685
      End
      Begin MSComctlLib.ListView lstAccType 
         Height          =   2625
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   4630
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
         MouseIcon       =   "AccType.frx":36A3
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ACCOUNT TYPE"
            Object.Width           =   7761
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAMISFILESAccType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAccType                                          As ADODB.Recordset
Dim AddorEdit                                          As String
Dim PrevCode                                           As String

Sub FillSearchGrid(XXX As Variant)
    Dim rsAccountType                                  As ADODB.Recordset
    lstAccType.Enabled = False
    lstAccType.Sorted = False: lstAccType.ListItems.Clear
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsAccountType = New ADODB.Recordset
    Set rsAccountType = gconDMIS.Execute("select code,description from AMIS_Acctype where Description like '" & XXX & "%'")
    If Not (rsAccountType.EOF And rsAccountType.BOF) Then
        Listview_Loadval Me.lstAccType.ListItems, rsAccountType
        lstAccType.Refresh
        lstAccType.Enabled = True
        lstAccType.Enabled = True
    Else
        lstAccType.Enabled = False
    End If

End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtCode.Text = ""
    txtDescription.Text = ""
End Sub

Sub rsRefresh()
    Set rsAccType = New ADODB.Recordset
    rsAccType.Open "select * from AMIS_Acctype order by code asc", gconDMIS, adOpenKeyset
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsAccountType                                  As ADODB.Recordset
    Set rsAccountType = New ADODB.Recordset
    Set rsAccountType = gconDMIS.Execute("select * from AMIS_Acctype where code = '" & XXX & "'")
    If Not rsAccountType.EOF And Not rsAccountType.BOF Then
        fraDetails.Enabled = False
        lstAccType.Enabled = False
        txtCode.Text = Null2String(rsAccountType!Code)
        txtDescription.Text = Null2String(rsAccountType!DESCRIPTION)
    End If
End Sub

Sub StoreMemVars()
    If Not rsAccType.EOF And Not rsAccType.BOF Then
        Frame1.Enabled = False
        txtCode.Text = Null2String(rsAccType!Code)
        txtDescription.Text = Null2String(rsAccType!DESCRIPTION)
    Else
        MessagePop RecNotFound, "Not Found", "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

'Upating Code       : AXP-0707200713:02
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "ACCOUNT TYPES") = False Then Exit Sub
    AddorEdit = "ADD": initMemvars: Picture1.Visible = False: Picture2.Visible = True
    On Error Resume Next
    txtCode.SetFocus
    lstAccType.Enabled = False
    txtSearch.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False: txtCode.Enabled = True: StoreMemVars: fraDetails.Enabled = True: lstAccType.Enabled = True
    lstAccType.Enabled = True
    txtSearch.Enabled = True
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "ACCOUNT TYPES") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from AMIS_Acctype where code = " & N2Str2Null((lstAccType.SelectedItem))
        rsRefresh
        StoreMemVars
        LogAudit "X", "ACCOUNT TYPE MASTER FILE", txtDescription
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:02
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "ACCOUNT TYPES") = False Then Exit Sub
    AddorEdit = "EDIT": Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True: txtCode.Enabled = False
    StoreEntry (lstAccType.SelectedItem)
    PrevCode = txtCode.Text
    On Error Resume Next
    txtCode.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSearch.SetFocus
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsAccType.MoveNext
    If rsAccType.EOF Then
        rsAccType.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsAccType.MovePrevious
    If rsAccType.BOF Then
        rsAccType.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "ACCOUNT TYPES") = False Then Exit Sub
    Screen.MousePointer = 11

    rptAccType.Reset
    rptAccType.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAccType.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptAccType.ReportTitle = " Account Type"
    PrintSQLReport rptAccType, AMIS_REPORT_PATH & "AccountFiles\AccType.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "ACCOUNT TYPE MASTER FILE", txtDescription
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdSave_Click()
    Dim VtxtCode, vtxtDescription                      As String
    On Error GoTo ErrorCode:

    VtxtCode = N2Str2Null(txtCode.Text)
    vtxtDescription = N2Str2Null(txtDescription.Text)
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into AMIS_Acctype " & _
                         "(code,Description) " & _
                         " values (" & VtxtCode & "," & vtxtDescription & ")"
        LogAudit "A", "ACCOUNT TYPE MASTER FILE", txtDescription
    Else
        gconDMIS.Execute "update AMIS_Acctype set" & _
                         " code = " & VtxtCode & "," & _
                         " Description = " & vtxtDescription & _
                         " where code = '" & PrevCode & "'"
        LogAudit "E", "ACCOUNT TYPE MASTER FILE", txtDescription
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsAccType.Find "code = " & VtxtCode
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub FillGrid()
    Dim rsAccountType                                  As ADODB.Recordset
    lstAccType.Enabled = False
    lstAccType.Sorted = False: lstAccType.ListItems.Clear
    Set rsAccountType = New ADODB.Recordset
    Set rsAccountType = gconDMIS.Execute("select code,description from AMIS_Acctype order by code asc")
    If Not (rsAccountType.EOF And rsAccountType.BOF) Then
        Listview_Loadval Me.lstAccType.ListItems, rsAccountType
        lstAccType.Refresh
        lstAccType.Enabled = True
        lstAccType.Enabled = True
    Else
        lstAccType.Enabled = False
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    initMemvars
    rsRefresh
    StoreMemVars
    FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub lstAccType_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstAccType
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstAccType_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstAccType_ItemClick(ByVal Item As MSComctlLib.ListItem)
'rsAccType.Bookmark = rsFind(rsAccType.Clone, "code", Me.lstAccType.SelectedItem).Bookmark
    rsRefresh
    rsAccType.Find "Code = '" & Item & "'"
    StoreMemVars
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub


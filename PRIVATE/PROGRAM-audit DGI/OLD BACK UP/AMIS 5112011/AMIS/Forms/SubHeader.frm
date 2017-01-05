VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISFILESSubHeader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extended Classification"
   ClientHeight    =   5085
   ClientLeft      =   1665
   ClientTop       =   1275
   ClientWidth     =   5745
   ForeColor       =   &H00FFFFFF&
   Icon            =   "SubHeader.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   5745
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   10
      Top             =   4185
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
         MouseIcon       =   "SubHeader.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "SubHeader.frx":0A1C
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
         MouseIcon       =   "SubHeader.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "SubHeader.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
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
         Height          =   795
         Left            =   3480
         MouseIcon       =   "SubHeader.frx":123A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
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
         Height          =   795
         Left            =   2790
         MouseIcon       =   "SubHeader.frx":138C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
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
         Height          =   795
         Left            =   2100
         MouseIcon       =   "SubHeader.frx":14DE
         MousePointer    =   99  'Custom
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
         MouseIcon       =   "SubHeader.frx":1630
         MousePointer    =   99  'Custom
         Picture         =   "SubHeader.frx":1782
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
         MouseIcon       =   "SubHeader.frx":1A7C
         MousePointer    =   99  'Custom
         Picture         =   "SubHeader.frx":1BCE
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
         MouseIcon       =   "SubHeader.frx":1F26
         MousePointer    =   99  'Custom
         Picture         =   "SubHeader.frx":2078
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
      Left            =   4230
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   19
      Top             =   4185
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
         MouseIcon       =   "SubHeader.frx":23D7
         MousePointer    =   99  'Custom
         Picture         =   "SubHeader.frx":2529
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
         MouseIcon       =   "SubHeader.frx":2867
         MousePointer    =   99  'Custom
         Picture         =   "SubHeader.frx":29B9
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
      Top             =   0
      Width           =   5625
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
         Left            =   1590
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "XX"
         Top             =   180
         Width           =   405
      End
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
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   570
         Width           =   3975
      End
      Begin Crystal.CrystalReport rptSubHeader 
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
         Caption         =   "Code Series"
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
         Left            =   270
         TabIndex        =   1
         Top             =   210
         Width           =   1245
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
         Left            =   270
         TabIndex        =   5
         Top             =   600
         Width           =   1245
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
      Top             =   930
      Width           =   5625
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
         Width           =   5445
      End
      Begin MSComctlLib.ListView lstSubHeader 
         Height          =   2625
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   5505
         _ExtentX        =   9710
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
         MouseIcon       =   "SubHeader.frx":2D09
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
Attribute VB_Name = "frmAMISFILESSubHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSubHeader                                        As ADODB.Recordset
Dim AddorEdit                                          As String
Dim PrevCode                                           As String

Sub rsRefresh()
    Set rsSubHeader = New ADODB.Recordset
    Set rsSubHeader = gconDMIS.Execute("select code,description from AMIS_SubHeader order by code asc")
End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtCode.Text = ""
    txtDescription.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsSubHeader.EOF And Not rsSubHeader.BOF Then
        Frame1.Enabled = False
        txtCode.Text = Null2String(rsSubHeader!Code)
        txtDescription.Text = Null2String(rsSubHeader!DESCRIPTION)
    Else
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsSubHeader2                                   As ADODB.Recordset
    Set rsSubHeader2 = New ADODB.Recordset
    Set rsSubHeader2 = gconDMIS.Execute("select * from AMIS_SubHeader where code = '" & XXX & "'")
    If Not rsSubHeader2.EOF And Not rsSubHeader2.BOF Then
        fraDetails.Enabled = False
        lstSubHeader.Enabled = False
        txtCode.Text = Null2String(rsSubHeader2!Code)
        txtDescription.Text = Null2String(rsSubHeader2!DESCRIPTION)
    End If
End Sub

Sub FillSearchGrid(XXX As Variant)
    Dim rsSubHeader2                                   As ADODB.Recordset
    lstSubHeader.Enabled = False
    lstSubHeader.Sorted = False: lstSubHeader.ListItems.Clear
    Set rsSubHeader2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsSubHeader2 = gconDMIS.Execute("select code,description from AMIS_SubHeader where Description like '" & XXX & "%'")
    If Not (rsSubHeader2.EOF And rsSubHeader2.BOF) Then
        Listview_Loadval Me.lstSubHeader.ListItems, rsSubHeader2
        lstSubHeader.Refresh
        lstSubHeader.Enabled = True
        lstSubHeader.Enabled = True
    Else
        lstSubHeader.Enabled = False
    End If

End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "EXTENDED CLASSIFICATION") = False Then Exit Sub
    AddorEdit = "ADD": initMemvars: Picture1.Visible = False: Picture2.Visible = True
    On Error Resume Next
    txtCode.SetFocus
    lstSubHeader.Enabled = False
    txtSearch.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False: StoreMemVars: fraDetails.Enabled = True: lstSubHeader.Enabled = True: FillGrid
    lstSubHeader.FindItem(txtCode.Text).EnsureVisible
    lstSubHeader.Enabled = True
    txtSearch.Enabled = True
End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "EXTENDED CLASSIFICATION") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from AMIS_SubHeader where code = " & N2Str2Null((lstSubHeader.SelectedItem))
        rsRefresh
        StoreMemVars
        LogAudit "X", "EXTENDED CLASSIFICATION MASTER FILE", txtCode & " - " & txtDescription
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "EXTENDED CLASSIFICATION") = False Then Exit Sub
    AddorEdit = "EDIT": Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True:
    StoreEntry (lstSubHeader.SelectedItem)
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

Private Sub cmdNext_Click()
    rsSubHeader.MoveNext
    If rsSubHeader.EOF Then
        rsSubHeader.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSubHeader.MovePrevious
    If rsSubHeader.BOF Then
        rsSubHeader.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "EXTENDED CLASSIFICATION") = False Then Exit Sub
    Screen.MousePointer = 11
    rptSubHeader.Reset
    rptSubHeader.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptSubHeader.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptSubHeader.ReportTitle = "EXTENDED CLASSIFICATION"
    PrintSQLReport rptSubHeader, AMIS_REPORT_PATH & "AccountFiles\SubHeader.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "EXTENDED CLASSIFICATION MASTER FILE", txtCode & " - " & txtDescription
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdSave_Click()

    Dim VtxtHeaderCode, VtxtSubHeaderCode, VtxtCode, vtxtDescription As String
    On Error GoTo ErrorCode:

    VtxtHeaderCode = N2Str2Null(Left(txtCode.Text, 1))
    VtxtSubHeaderCode = N2Str2Null(Mid(txtCode.Text, 2, 1))
    VtxtCode = N2Str2Null(txtCode.Text)
    vtxtDescription = N2Str2Null(txtDescription.Text)
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into AMIS_SubHeader " & _
                         "(HeaderCode,SubHeaderCode,code,Description) " & _
                         " values (" & VtxtHeaderCode & "," & VtxtSubHeaderCode & "," & VtxtCode & "," & vtxtDescription & ")"
        LogAudit "A", "EXTENDED CLASSIFICATION MASTER FILE", txtCode & " - " & txtDescription
    Else
        gconDMIS.Execute "update AMIS_SubHeader set" & _
                         " Headercode = " & VtxtHeaderCode & "," & _
                         " SubHeadercode = " & VtxtSubHeaderCode & "," & _
                         " code = " & VtxtCode & "," & _
                         " Description = " & vtxtDescription & _
                         " where code = '" & PrevCode & "'"
        LogAudit "E", "EXTENDED CLASSIFICATION MASTER FILE", txtCode & " - " & txtDescription
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsSubHeader.Find "code = " & VtxtCode
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
    initMemvars
    rsRefresh
    StoreMemVars
    FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub FillGrid()
    Dim rsSubHeader2                                   As ADODB.Recordset
    lstSubHeader.Enabled = False
    lstSubHeader.Sorted = False: lstSubHeader.ListItems.Clear
    Set rsSubHeader2 = New ADODB.Recordset
    Set rsSubHeader2 = gconDMIS.Execute("select code,description from AMIS_SubHeader order by code asc")
    If Not (rsSubHeader2.EOF And rsSubHeader2.BOF) Then
        Listview_Loadval Me.lstSubHeader.ListItems, rsSubHeader2
        lstSubHeader.Refresh
        lstSubHeader.Enabled = True
        lstSubHeader.Enabled = True
    Else
        lstSubHeader.Enabled = False
    End If

End Sub

Private Sub lstSubHeader_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstSubHeader
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstSubHeader_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstSubHeader_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsSubHeader.Bookmark = rsFind(rsSubHeader.Clone, "code", STR(lstSubHeader.SelectedItem)).Bookmark
    StoreMemVars
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub


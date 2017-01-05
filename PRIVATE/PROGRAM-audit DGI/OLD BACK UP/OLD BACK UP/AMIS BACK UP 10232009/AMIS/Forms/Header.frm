VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAMISFILESHeader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Classification"
   ClientHeight    =   5130
   ClientLeft      =   1665
   ClientTop       =   1170
   ClientWidth     =   5865
   ForeColor       =   &H00F5F5F5&
   Icon            =   "Header.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   5865
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   12
      Top             =   4200
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
         MouseIcon       =   "Header.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Header.frx":0A1C
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
         MouseIcon       =   "Header.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Header.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "Header.frx":123A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "Header.frx":138C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "Header.frx":14DE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "Header.frx":1630
         MousePointer    =   99  'Custom
         Picture         =   "Header.frx":1782
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
         MouseIcon       =   "Header.frx":1A7C
         MousePointer    =   99  'Custom
         Picture         =   "Header.frx":1BCE
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
         MouseIcon       =   "Header.frx":1F26
         MousePointer    =   99  'Custom
         Picture         =   "Header.frx":2078
         Style           =   1  'Graphical
         TabIndex        =   13
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
      ScaleWidth      =   1485
      TabIndex        =   21
      Top             =   4215
      Width           =   1485
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
         MouseIcon       =   "Header.frx":23D7
         MousePointer    =   99  'Custom
         Picture         =   "Header.frx":2529
         Style           =   1  'Graphical
         TabIndex        =   22
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
         MouseIcon       =   "Header.frx":2867
         MousePointer    =   99  'Custom
         Picture         =   "Header.frx":29B9
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00973640&
         Height          =   360
         Left            =   1590
         TabIndex        =   8
         Text            =   "cboType"
         Top             =   960
         Width           =   1935
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
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   570
         Width           =   3975
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
         Left            =   1590
         MaxLength       =   1
         TabIndex        =   2
         Text            =   "X"
         Top             =   180
         Width           =   255
      End
      Begin Crystal.CrystalReport rptHeader 
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
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
         TabIndex        =   7
         Top             =   990
         Width           =   1485
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
         TabIndex        =   6
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
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   570
         Width           =   225
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2745
      Left            =   60
      TabIndex        =   9
      Top             =   1380
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
         TabIndex        =   10
         Top             =   150
         Width           =   5445
      End
      Begin MSComctlLib.ListView lstHeader 
         Height          =   2145
         Left            =   60
         TabIndex        =   11
         Top             =   540
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   3784
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
         MouseIcon       =   "Header.frx":2D09
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
Attribute VB_Name = "frmAMISFILESHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAccType                                                         As ADODB.Recordset
Dim rsHeader                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim PrevCode                                                          As String

Function SetAccType(Acc As String) As String
    Set rsAccType = New ADODB.Recordset
    Set rsAccType = gconDMIS.Execute("select * from AMIS_Acctype where code = " & N2Str2Null(Acc))
    If Not rsAccType.EOF And Not rsAccType.BOF Then
        SetAccType = Null2String(rsAccType!Description)
    End If
End Function

Function SetAccCode(Acc As String) As String
    Set rsAccType = New ADODB.Recordset
    Set rsAccType = gconDMIS.Execute("select * from AMIS_Acctype where description = " & N2Str2Null(Acc))
    If Not rsAccType.EOF And Not rsAccType.BOF Then
        SetAccCode = Null2String(rsAccType!code)
    End If
End Function

Sub rsRefresh()
    Set rsHeader = New ADODB.Recordset
    Set rsHeader = gconDMIS.Execute("select code,description,accttype from AMIS_Header order by code asc")
End Sub

Sub InitMemVars()
    Frame1.Enabled = True
    txtCode.Text = ""
    txtDescription.Text = ""
    Set rsAccType = New ADODB.Recordset
    Set rsAccType = gconDMIS.Execute("select Description from AMIS_Acctype order by code asc")
    If Not rsAccType.EOF And Not rsAccType.BOF Then
        Combo_Loadval cboType, rsAccType
    End If
End Sub

Sub StoreMemvars()
    If Not rsHeader.EOF And Not rsHeader.BOF Then
        Frame1.Enabled = False
        txtCode.Text = Null2String(rsHeader!code)
        txtDescription.Text = Null2String(rsHeader!Description)
        cboType.Text = SetAccType(Null2String(rsHeader!AcctType))
    Else
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsHeader2                                                     As ADODB.Recordset
    Set rsHeader2 = New ADODB.Recordset
    Set rsHeader2 = gconDMIS.Execute("select * from AMIS_Header where code = '" & XXX & "'")
    If Not rsHeader2.EOF And Not rsHeader2.BOF Then
        fraDetails.Enabled = False
        lstHeader.Enabled = False
        txtCode.Text = Null2String(rsHeader2!code)
        txtDescription.Text = Null2String(rsHeader2!Description)
        cboType.Text = SetAccType(Null2String(rsHeader2!AcctType))
    End If
End Sub

Sub FillSearchGrid(XXX As Variant)
    Dim rsHeader2                                                     As ADODB.Recordset
    lstHeader.Enabled = False
    lstHeader.Sorted = False: lstHeader.ListItems.Clear
    Set rsHeader2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsHeader2 = gconDMIS.Execute("select code,description from AMIS_Header where Description like '" & XXX & "%'")
    If Not (rsHeader2.EOF And rsHeader2.BOF) Then
        Listview_Loadval Me.lstHeader.ListItems, rsHeader2
        lstHeader.Refresh
        lstHeader.Enabled = True
        lstHeader.Enabled = True
    Else
        lstHeader.Enabled = False
    End If

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "ACCOUNT CLASSIFICATION") = False Then Exit Sub
    AddorEdit = "ADD": InitMemVars: Picture1.Visible = False: Picture2.Visible = True
    On Error Resume Next
    txtCode.SetFocus
    lstHeader.Enabled = False
    txtSEARCH.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False: StoreMemvars: fraDetails.Enabled = True: lstHeader.Enabled = True
    lstHeader.FindItem(txtCode.Text).EnsureVisible
    lstHeader.Enabled = True
    txtSEARCH.Enabled = True
End Sub

'Upating Code       : AXP-0707200713:07
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Delete", "ACCOUNT CLASSIFICATION") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from AMIS_Header where code = " & N2Str2Null((lstHeader.SelectedItem))
        rsRefresh
        StoreMemvars
        LogAudit "X", "ACCOUNT CLASSIFICATION MASTER FILE", cboType & " - " & txtDescription
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "ACCOUNT CLASSIFICATION") = False Then Exit Sub
    AddorEdit = "EDIT": Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True:
    StoreEntry (lstHeader.SelectedItem)
    PrevCode = txtCode.Text
    On Error Resume Next
    txtCode.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                                                       As String
    findStr = InputBox("Please Input Header ...", "Find")
    If findStr <> "" Then
        On Error GoTo Errorcode
        rsHeader.Bookmark = rsFind(rsHeader.Clone, "Description", findStr).Bookmark
    End If
    StoreMemvars
    Exit Sub

Errorcode:
    If Err.Number = 3021 Then
        MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
        Resume Next
    End If
End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    rsHeader.MoveNext
    If rsHeader.EOF Then
        rsHeader.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rsHeader.MovePrevious
    If rsHeader.BOF Then
        rsHeader.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Print", "ACCOUNT CLASSIFICATION") = False Then Exit Sub
    Screen.MousePointer = 11
    rptHeader.Reset
    rptHeader.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptHeader.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptHeader.ReportTitle = "ACCOUNT CLASSIFICATION"
    PrintSQLReport rptHeader, AMIS_REPORT_PATH & "ACCOUNTFILES\ACCOUNT CLASSIFICATION.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "ACCOUNT CLASSIFICATION MASTER FILE", cboType & " - " & txtDescription
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:08
Private Sub cmdSave_Click()
    Dim VtxtCode, vtxtDescription, VcboType                           As String
    On Error GoTo Errorcode:

    VtxtCode = N2Str2Null(txtCode.Text)
    vtxtDescription = N2Str2Null(txtDescription.Text)
    VcboType = N2Str2Null(SetAccCode(cboType.Text))
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into AMIS_Header " & _
                         "(code,Description,AcctType) " & _
                       " values (" & VtxtCode & "," & vtxtDescription & "," & VcboType & ")"
        LogAudit "A", "ACCOUNT CLASSIFICATION MASTER FILE", cboType & " - " & txtDescription
    Else
        gconDMIS.Execute "update AMIS_Header set" & _
                       " code = " & VtxtCode & "," & _
                       " Description = " & vtxtDescription & "," & _
                       " AcctType = " & VcboType & _
                       " where code = '" & PrevCode & "'"
        LogAudit "E", "ACCOUNT CLASSIFICATION MASTER FILE", cboType & " - " & txtDescription
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsHeader.Find "code = " & VtxtCode
    cmdCancel.Value = True
    Exit Sub
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemVars
    rsRefresh
    StoreMemvars
    FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub FillGrid()
    Dim rsHeader2                                                     As ADODB.Recordset
    lstHeader.Enabled = False
    lstHeader.Sorted = False: lstHeader.ListItems.Clear
    Set rsHeader2 = New ADODB.Recordset
    Set rsHeader2 = gconDMIS.Execute("select code,description from AMIS_Header")
    If Not (rsHeader2.EOF And rsHeader2.BOF) Then
        Listview_Loadval Me.lstHeader.ListItems, rsHeader2
        lstHeader.Refresh
        lstHeader.Enabled = True
        lstHeader.Enabled = True
    Else
        lstHeader.Enabled = False
    End If

End Sub

Private Sub lstHeader_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstHeader
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstHeader_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstHeader_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsHeader.Bookmark = rsFind(rsHeader.Clone, "code", STR(lstHeader.SelectedItem)).Bookmark
    StoreMemvars
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSEARCH_Change()
    If Trim(txtSEARCH.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSEARCH.Text)
    End If
End Sub


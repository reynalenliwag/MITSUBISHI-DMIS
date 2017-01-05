VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISFILESTitleCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Sub-Totals"
   ClientHeight    =   5265
   ClientLeft      =   1665
   ClientTop       =   1275
   ClientWidth     =   5745
   ForeColor       =   &H00FFFFFF&
   Icon            =   "TitleCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   5745
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   11
      Top             =   4350
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
         MouseIcon       =   "TitleCode.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "TitleCode.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "TitleCode.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "TitleCode.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":1809
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
         MouseIcon       =   "TitleCode.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":1CB7
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
         MouseIcon       =   "TitleCode.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "TitleCode.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   13
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
         MouseIcon       =   "TitleCode.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   12
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
      TabIndex        =   20
      Top             =   4365
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
         MouseIcon       =   "TitleCode.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   21
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
         MouseIcon       =   "TitleCode.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "TitleCode.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   30
      TabIndex        =   3
      Top             =   -30
      Width           =   5625
      Begin VB.ComboBox cboCashFlowCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1770
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1620
         Width           =   3765
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
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   570
         Width           =   4305
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
         Left            =   1230
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "XXXX"
         Top             =   180
         Width           =   765
      End
      Begin Crystal.CrystalReport rptTitleCode 
         Left            =   5100
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Account Titles"
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
         Caption         =   "Cash Flow Code"
         Enabled         =   0   'False
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
         Left            =   -630
         TabIndex        =   23
         Top             =   1650
         Width           =   2295
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
         Left            =   -60
         TabIndex        =   4
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
         Left            =   -60
         TabIndex        =   7
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
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   570
         Width           =   225
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3225
      Left            =   30
      TabIndex        =   8
      Top             =   1050
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
         TabIndex        =   9
         Top             =   150
         Width           =   5445
      End
      Begin MSComctlLib.ListView lstTitleCode 
         Height          =   2625
         Left            =   60
         TabIndex        =   10
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
         MouseIcon       =   "TitleCode.frx":36A3
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
Attribute VB_Name = "frmAMISFILESTitleCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTitleCode                                        As ADODB.Recordset
Dim AddorEdit                                          As String
Dim PrevCode                                           As String

Sub rsRefresh()
    Set rsTitleCode = New ADODB.Recordset
    Set rsTitleCode = gconDMIS.Execute("select code,description,CashFlowCode from AMIS_TitleCode order by code asc")
End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtCode.Text = ""
    txtDescription.Text = ""
    cboCashFlowCode.ListIndex = -1
    cboCashFlowCode.AddItem "OPERATING ACTIVITIES"
    cboCashFlowCode.AddItem "INVESTING ACTIVITIES"
    cboCashFlowCode.AddItem "FINANCING ACTIVITIES"
    cboCashFlowCode.AddItem "CASH EQUIVALENTS"
End Sub

Sub StoreMemVars()
    If Not rsTitleCode.EOF And Not rsTitleCode.BOF Then
        Frame1.Enabled = False
        txtCode.Text = Null2String(rsTitleCode!Code)
        txtDescription.Text = Null2String(rsTitleCode!DESCRIPTION)
        If Null2String(rsTitleCode!CashFlowCode) = "OA" Then cboCashFlowCode.Text = "OPERATING ACTIVITIES"
        If Null2String(rsTitleCode!CashFlowCode) = "IA" Then cboCashFlowCode.Text = "INVESTING ACTIVITIES"
        If Null2String(rsTitleCode!CashFlowCode) = "FA" Then cboCashFlowCode.Text = "FINANCING ACTIVITIES"
        If Null2String(rsTitleCode!CashFlowCode) = "CE" Then cboCashFlowCode.Text = "CASH EQUIVALENTS"
        If Null2String(rsTitleCode!CashFlowCode) = "" Then cboCashFlowCode.ListIndex = -1
    Else
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsTitleCode2                                   As ADODB.Recordset
    Set rsTitleCode2 = New ADODB.Recordset
    Set rsTitleCode2 = gconDMIS.Execute("select * from AMIS_TitleCode where code = '" & XXX & "'")
    If Not rsTitleCode2.EOF And Not rsTitleCode2.BOF Then
        fraDetails.Enabled = False
        lstTitleCode.Enabled = False
        txtCode.Text = Null2String(rsTitleCode2!Code)
        txtDescription.Text = Null2String(rsTitleCode2!DESCRIPTION)
        If Null2String(rsTitleCode2!CashFlowCode) = "OA" Then cboCashFlowCode.Text = "OPERATING ACTIVITIES"
        If Null2String(rsTitleCode2!CashFlowCode) = "IA" Then cboCashFlowCode.Text = "INVESTING ACTIVITIES"
        If Null2String(rsTitleCode2!CashFlowCode) = "FA" Then cboCashFlowCode.Text = "FINANCING ACTIVITIES"
        If Null2String(rsTitleCode2!CashFlowCode) = "CE" Then cboCashFlowCode.Text = "CASH EQUIVALENTS"
        If Null2String(rsTitleCode2!CashFlowCode) = "" Then cboCashFlowCode.ListIndex = -1
    End If
End Sub

Sub FillSearchGrid(XXX As Variant)
    Dim rsTitleCode2                                   As ADODB.Recordset
    lstTitleCode.Enabled = False
    lstTitleCode.Sorted = False: lstTitleCode.ListItems.Clear
    Set rsTitleCode2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsTitleCode2 = gconDMIS.Execute("select code,description from AMIS_TitleCode where Description like '" & XXX & "%'")
    If Not (rsTitleCode2.EOF And rsTitleCode2.BOF) Then
        Listview_Loadval Me.lstTitleCode.ListItems, rsTitleCode2
        lstTitleCode.Refresh
        lstTitleCode.Enabled = True
        lstTitleCode.Enabled = True
    Else
        lstTitleCode.Enabled = False
    End If

End Sub

Private Sub cboCashFlowCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboCashFlowCode.ListIndex = -1
End Sub

'Upating Code       : AXP-0707200713:09
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "ACCOUNT SUB TOTALS") = False Then Exit Sub
    AddorEdit = "ADD": initMemvars: Picture1.Visible = False: Picture2.Visible = True
    On Error Resume Next
    txtCode.SetFocus
    lstTitleCode.Enabled = False
    txtSearch.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False: StoreMemVars: fraDetails.Enabled = True: lstTitleCode.Enabled = True: FillGrid
    lstTitleCode.FindItem(txtCode.Text).EnsureVisible
    lstTitleCode.Enabled = True
    txtSearch.Enabled = True
End Sub

'Upating Code       : AXP-0707200713:09
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "ACCOUNT SUB TOTALS") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from AMIS_TitleCode where code = " & N2Str2Null((lstTitleCode.SelectedItem))
        rsRefresh
        StoreMemVars
        FillGrid
        LogAudit "X", "ACCOUNT SUB-TOTALS", txtCode & "-" & txtDescription
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:09
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "ACCOUNT SUB TOTALS") = False Then Exit Sub

    AddorEdit = "EDIT": Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True:
    StoreEntry (lstTitleCode.SelectedItem)
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

'Upating Code       : AXP-0707200713:09
Private Sub cmdFind_Click()
    On Error GoTo ErrorCode:

    txtSearch.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:09
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsTitleCode.MoveNext
    If rsTitleCode.EOF Then
        rsTitleCode.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:09
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsTitleCode.MovePrevious
    If rsTitleCode.BOF Then
        rsTitleCode.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:09
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "ACCOUNT SUB TOTALS") = False Then Exit Sub

    Screen.MousePointer = 11
    rptTitleCode.Reset
    rptTitleCode.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptTitleCode.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptTitleCode.ReportTitle = "Accounts Sub Totals"
    PrintSQLReport rptTitleCode, AMIS_REPORT_PATH & "ACCOUNTFILES\TitleCode.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "ACCOUNT SUB-TOTALS", txtCode & "-" & txtDescription
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:09
Private Sub cmdSave_Click()

    Dim VtxtCode, vtxtDescription                      As String
    Dim VHeaderCode, VSubHeaderCode, VSubTitleCode, vCashFlowCode As String
    On Error GoTo ErrorCode:

    VtxtCode = N2Str2Null(txtCode.Text)
    vtxtDescription = N2Str2Null(txtDescription.Text)
    VHeaderCode = N2Str2Null(Left(txtCode.Text, 1))
    VSubHeaderCode = N2Str2Null(Mid(txtCode.Text, 1, 2))
    VSubTitleCode = N2Str2Null(Right(txtCode.Text, 2))
    If cboCashFlowCode.Text = "OPERATING ACTIVITIES" Then
        vCashFlowCode = "'OA'"
    ElseIf cboCashFlowCode.Text = "INVESTING ACTIVITIES" Then
        vCashFlowCode = "'IA'"
    ElseIf cboCashFlowCode.Text = "FINANCING ACTIVITIES" Then
        vCashFlowCode = "'FA'"
    ElseIf cboCashFlowCode.Text = "CASH EQUIVALENTS" Then
        vCashFlowCode = "'CE'"
    Else
        vCashFlowCode = "NULL"
    End If

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into AMIS_TitleCode " & _
                         "(HeaderCode,SubHeaderCode,SubTitleCode,Code,Description,CashFlowCode) " & _
                         " values (" & VHeaderCode & "," & VSubHeaderCode & "," & VSubTitleCode & "," & VtxtCode & "," & vtxtDescription & "," & vCashFlowCode & ")"
        LogAudit "A", "ACCOUNT SUB-TOTALS", txtCode & "-" & txtDescription
    Else
        If txtCode.Text <> PrevCode Then
            Dim rsCheckCode                            As ADODB.Recordset
            Set rsCheckCode = New ADODB.Recordset
            Set rsCheckCode = gconDMIS.Execute("Select * from AMIS_TitleCode Where Code = '" & txtCode.Text & "'")
            If Not rsCheckCode.EOF And Not rsCheckCode.BOF Then
                MsgBox "Code Already Exist!", vbExclamation, "Warning"
                Exit Sub
            End If
        End If
        gconDMIS.Execute "update AMIS_TitleCode set" & _
                         " HeaderCode = " & VHeaderCode & "," & _
                         " SubHeaderCode = " & VSubHeaderCode & "," & _
                         " SubTitleCode = " & VSubTitleCode & "," & _
                         " code = " & VtxtCode & "," & _
                         " CashFlowcode = " & vCashFlowCode & "," & _
                         " Description = " & vtxtDescription & _
                         " where code = '" & PrevCode & "'"
        LogAudit "E", "ACCOUNT SUB-TOTALS", txtCode & "-" & txtDescription
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsTitleCode.Find "code = " & VtxtCode
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
    Dim rsTitleCode2                                   As ADODB.Recordset
    lstTitleCode.Enabled = False
    lstTitleCode.Sorted = False: lstTitleCode.ListItems.Clear
    Set rsTitleCode2 = New ADODB.Recordset
    Set rsTitleCode2 = gconDMIS.Execute("select code,description from AMIS_TitleCode order by code asc")
    If Not (rsTitleCode2.EOF And rsTitleCode2.BOF) Then
        Listview_Loadval Me.lstTitleCode.ListItems, rsTitleCode2
        lstTitleCode.Refresh
        lstTitleCode.Enabled = True
        lstTitleCode.Enabled = True
    Else
        lstTitleCode.Enabled = False
    End If

End Sub

Private Sub lstTitleCode_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstTitleCode
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstTitleCode_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstTitleCode_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsTitleCode.Bookmark = rsFind(rsTitleCode.Clone, "code", STR(lstTitleCode.SelectedItem)).Bookmark
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


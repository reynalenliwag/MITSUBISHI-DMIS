VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISMASTERFILEPayTerm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Term"
   ClientHeight    =   5025
   ClientLeft      =   855
   ClientTop       =   750
   ClientWidth     =   5715
   ForeColor       =   &H00FFFFFF&
   Icon            =   "PayTerm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5715
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   5625
      TabIndex        =   11
      Top             =   4095
      Width           =   5625
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
         MouseIcon       =   "PayTerm.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":0A1C
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
         MouseIcon       =   "PayTerm.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":0ED4
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
         MouseIcon       =   "PayTerm.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":138C
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
         MouseIcon       =   "PayTerm.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":1809
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
         MouseIcon       =   "PayTerm.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":1CB7
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
         MouseIcon       =   "PayTerm.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":211C
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
         MouseIcon       =   "PayTerm.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":2568
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
         MouseIcon       =   "PayTerm.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":2A12
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
      Left            =   4185
      ScaleHeight     =   885
      ScaleWidth      =   1485
      TabIndex        =   20
      Top             =   4095
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
         MouseIcon       =   "PayTerm.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":2EC3
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
         MouseIcon       =   "PayTerm.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "PayTerm.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin Crystal.CrystalReport rptPayTerm 
      Left            =   2610
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Payment Terms"
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
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   5625
      Begin VB.TextBox txtPay_Desc 
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
         Left            =   1290
         MaxLength       =   25
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   570
         Width           =   4245
      End
      Begin VB.TextBox txtNo_Days 
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
         Left            =   1290
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   465
      End
      Begin VB.TextBox txtPay_Code 
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
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label5 
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
         Left            =   -900
         TabIndex        =   3
         Top             =   630
         Width           =   2115
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Days"
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
         Left            =   -900
         TabIndex        =   7
         Top             =   1020
         Width           =   2115
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
         Left            =   4680
         TabIndex        =   6
         Top             =   630
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
         Left            =   3840
         TabIndex        =   5
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label3 
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
         Left            =   -1020
         TabIndex        =   1
         Top             =   240
         Width           =   2235
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2775
      Left            =   30
      TabIndex        =   9
      Top             =   1260
      Width           =   5625
      Begin MSComctlLib.ListView lstPayTerm 
         Height          =   2565
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   4524
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
         MouseIcon       =   "PayTerm.frx":36A3
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "NUM. OF DAYS"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAMISMASTERFILEPayTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPayTerm                                     As ADODB.Recordset
Dim AddorEdit, PrevPayCode                        As String
Attribute PrevPayCode.VB_VarUserMemId = 1073938433

Sub rsRefresh()
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "select * from ALL_PayTerm order by Pay_Code asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtPay_Code.Text = ""
    txtPay_Desc.Text = ""
    txtNo_Days.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        Frame1.Enabled = False
        labID.Caption = rsPayTerm!ID
        txtPay_Code.Text = Null2String(rsPayTerm!pay_Code)
        txtPay_Desc.Text = Null2String(rsPayTerm!pay_desc)
        txtNo_Days.Text = N2Str2Zero(rsPayTerm!no_Days)
    Else
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsPayTerm2                                As ADODB.Recordset
    Set rsPayTerm2 = New ADODB.Recordset
    rsPayTerm2.Open "select * from ALL_PayTerm where id = " & XXX, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm2.BOF Then
        labID.Caption = rsPayTerm2!ID
        fraDetails.Enabled = False
        lstPayTerm.Enabled = False
        txtPay_Code.Text = Null2String(rsPayTerm2!pay_Code)
        txtPay_Desc.Text = Null2String(rsPayTerm2!pay_desc)
        txtNo_Days.Text = N2Str2Zero(rsPayTerm2!no_Days)
    End If
End Sub

'Upating Code       : AXP-0713200713:54
Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Add", "TERMS OF PAYMENT") = False Then Exit Sub

    AddorEdit = "ADD"
    initMemvars
    Picture1.Visible = False
    Picture2.Visible = True
    lstPayTerm.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstPayTerm.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
    FillGrid
End Sub

'Upating Code       : AXP-0707200713:07
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Delete", "TERMS OF PAYMENT") = False Then Exit Sub

    If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
        SQL_STATEMENT = "delete from ALL_PayTerm where ID = " & lstPayTerm.SelectedItem.SubItems(3)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "X", "PAYMENT TERM", SQL_STATEMENT, labID.Caption, "", txtPay_Code, "", ""
    End If
    rsRefresh
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:54
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Edit", "TERMS OF PAYMENT") = False Then Exit Sub

    AddorEdit = "EDIT"
    PrevPayCode = txtPay_Code.Text
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    StoreEntry (lstPayTerm.SelectedItem.SubItems(3))
    lstPayTerm.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                                   As String
    findStr = InputBox("Please Input Payment Term ...", "Find")
    If findStr <> "" Then
        On Error Resume Next
        rsPayTerm.Bookmark = rsFind(rsPayTerm.Clone, "Pay_Code", findStr).Bookmark
        If Err.Number = 3021 Then
            On Error GoTo Errorcode
            rsPayTerm.Bookmark = rsFind(rsPayTerm.Clone, "Pay_Desc", findStr).Bookmark
        End If
    End If
    StoreMemVars
    Exit Sub

Errorcode:
    If Err.Number = 3021 Then
        MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
        Resume Next
    End If
End Sub

'Upating Code       : AXP-0713200713:54
Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    rsPayTerm.MoveNext
    If rsPayTerm.EOF Then
        rsPayTerm.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:54
Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rsPayTerm.MovePrevious
    If rsPayTerm.BOF Then
        rsPayTerm.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:54
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Print", "TERMS OF PAYMENT") = False Then Exit Sub

    Screen.MousePointer = 11
    rptPayTerm.Reset
    rptPayTerm.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPayTerm.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptPayTerm.ReportTitle = "TERMS OF PAYMENT"
    PrintSQLReport rptPayTerm, AMIS_REPORT_PATH & "files\PayTerm.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    NEW_LogAudit "V", "PAYMENT TERMS", "", labID.Caption, "", txtPay_Code, "", ""
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:54
Private Sub cmdSave_Click()

    Dim VtxtPay_Code, VtxtPay_Desc                As String
    Dim VtxtNo_Days                               As Double
    On Error GoTo Errorcode:

    VtxtPay_Code = N2Str2Null(txtPay_Code.Text)
    VtxtPay_Desc = N2Str2Null(txtPay_Desc.Text)
    VtxtNo_Days = NumericVal(txtNo_Days.Text)

    If AddorEdit = "ADD" Then
        Dim rsPayTermDup                          As ADODB.Recordset
        Set rsPayTermDup = New ADODB.Recordset
        rsPayTermDup.Open "select Pay_Code from ALL_PayTerm where Pay_Code = " & VtxtPay_Code, gconDMIS
        If Not rsPayTermDup.EOF And Not rsPayTermDup.BOF Then
            MsgBox "Payment Code Already Exist!", vbCritical, "Duplicate Payment CodeNot Allowed"
            Exit Sub
        End If
        SQL_STATEMENT = "Insert into ALL_PayTerm " & _
                        "(No_Days,Pay_Code,Pay_Desc) " & _
                        " values (" & VtxtNo_Days & ", " & VtxtPay_Code & ", " & VtxtPay_Desc & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "PAYMENT TERM", SQL_STATEMENT, labID.Caption, "", txtPay_Code, "", ""
    Else
        SQL_STATEMENT = "update ALL_PayTerm set" & _
                        " No_Days = " & VtxtNo_Days & "," & _
                        " Pay_Code = " & VtxtPay_Code & "," & _
                        " Pay_Desc = " & VtxtPay_Desc & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "PAYMENT TERM", SQL_STATEMENT, labID.Caption, "", txtPay_Code, "", ""
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsPayTerm.Find "Pay_Code = " & VtxtPay_Code
    cmdCancel.Value = True
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "PAYMENT TERM"
        Call frmALL_AuditInquiry.DisplayHistory(labID, "PAYMENT TERM")
    End Select

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
    Dim rsPayTerm2                                As ADODB.Recordset
    lstPayTerm.Enabled = False
    lstPayTerm.Sorted = False: lstPayTerm.ListItems.Clear
    Set rsPayTerm2 = New ADODB.Recordset
    Set rsPayTerm2 = gconDMIS.Execute("select pay_code,pay_desc,no_days,ID from ALL_PayTerm")
    If Not (rsPayTerm2.EOF And rsPayTerm2.BOF) Then
        lstPayTerm.Enabled = True
        Listview_Loadval Me.lstPayTerm.ListItems, rsPayTerm2
        lstPayTerm.Refresh
        lstPayTerm.Enabled = True
    Else
        lstPayTerm.Enabled = False
    End If

End Sub

Private Sub lstPayTerm_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPayTerm
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

Private Sub lstPayTerm_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstPayTerm_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsPayTerm.Bookmark = rsFind(rsPayTerm.Clone, "pay_code", Item).Bookmark
    StoreMemVars
End Sub

Private Sub txtPay_Code_LostFocus()
    txtPay_Code.Text = UCase(txtPay_Code.Text)
End Sub


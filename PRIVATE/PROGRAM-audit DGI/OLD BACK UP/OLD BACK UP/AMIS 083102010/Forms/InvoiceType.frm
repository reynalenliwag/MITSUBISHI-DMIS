VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISMASTERFILEInvoiceType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Type Master List"
   ClientHeight    =   4140
   ClientLeft      =   1125
   ClientTop       =   855
   ClientWidth     =   5685
   ForeColor       =   &H00FFFFFF&
   Icon            =   "InvoiceType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5685
   Begin Crystal.CrystalReport rptInvoiceType 
      Left            =   5730
      Top             =   270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   45
      ScaleHeight     =   855
      ScaleWidth      =   5940
      TabIndex        =   9
      Top             =   3195
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
         MouseIcon       =   "InvoiceType.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":0A1C
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
         MouseIcon       =   "InvoiceType.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":0ED4
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
         MouseIcon       =   "InvoiceType.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":138C
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
         MouseIcon       =   "InvoiceType.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":1809
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
         MouseIcon       =   "InvoiceType.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":1CB7
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
         MouseIcon       =   "InvoiceType.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":211C
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
         MouseIcon       =   "InvoiceType.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":2568
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
         MouseIcon       =   "InvoiceType.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin Crystal.CrystalReport rptBanks 
      Left            =   2340
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Banks"
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
      Height          =   975
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   5625
      Begin VB.TextBox txtInvType 
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
         Text            =   "Text1"
         Top             =   540
         Width           =   4815
      End
      Begin VB.TextBox txtInvCode 
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
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   -540
         TabIndex        =   3
         Top             =   600
         Width           =   1215
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
         Left            =   4710
         TabIndex        =   6
         Top             =   600
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
         TabIndex        =   5
         Top             =   600
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
         Left            =   -660
         TabIndex        =   1
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2325
      Left            =   30
      TabIndex        =   7
      Top             =   840
      Width           =   5625
      Begin MSComctlLib.ListView lstInvoiceType 
         Height          =   2115
         Left            =   30
         TabIndex        =   8
         Top             =   150
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3731
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
         MouseIcon       =   "InvoiceType.frx":2D71
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "INVOICE TYPE"
            Object.Width           =   6702
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4185
      ScaleHeight     =   885
      ScaleWidth      =   1620
      TabIndex        =   18
      Top             =   3195
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
         MouseIcon       =   "InvoiceType.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":3025
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
         MouseIcon       =   "InvoiceType.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "InvoiceType.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAMISMASTERFILEInvoiceType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInvoiceType                                 As ADODB.Recordset
Dim AddorEdit                                     As String

Sub rsRefresh()
    Set rsInvoiceType = New ADODB.Recordset
    rsInvoiceType.Open "select * from ALL_InvoiceType order by InvCode asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtInvCode.Text = ""
    txtInvType.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        Frame1.Enabled = False
        labID.Caption = rsInvoiceType!ID
        txtInvCode.Text = Null2String(rsInvoiceType!InvCode)
        txtInvType.Text = Null2String(rsInvoiceType!INVTYPE)
    Else
        lstInvoiceType.ListItems.Clear
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsInvoiceType2                            As ADODB.Recordset
    Set rsInvoiceType2 = New ADODB.Recordset
    rsInvoiceType2.Open "select * from ALL_InvoiceType where ID = " & XXX, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsInvoiceType2.EOF And Not rsInvoiceType2.BOF Then
        fraDetails.Enabled = False
        lstInvoiceType.Enabled = False
        labID.Caption = rsInvoiceType2!ID
        txtInvCode.Text = Null2String(rsInvoiceType2!InvCode)
        txtInvType.Text = Null2String(rsInvoiceType2!INVTYPE)
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "INVOICE TYPES") = False Then Exit Sub

    AddorEdit = "ADD"
    initMemvars
    Picture1.Visible = False
    Picture2.Visible = True
    lstInvoiceType.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstInvoiceType.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
    FillGrid
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "INVOICE TYPES") = False Then Exit Sub

    If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
        SQL_STATEMENT = "delete from ALL_InvoiceType where ID = " & lstInvoiceType.SelectedItem.SubItems(2)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "X", "INVOICE TYPES", SQL_STATEMENT, labID.Caption, "", txtInvCode, "", ""
    End If
    rsRefresh
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "INVOICE TYPES") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    StoreEntry (lstInvoiceType.SelectedItem.SubItems(2))
    lstInvoiceType.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                                   As String
    findStr = InputBox("Please Input InvoiceType ...", "Find")
    If findStr <> "" Then
        On Error Resume Next
        rsInvoiceType.Bookmark = rsFind(rsInvoiceType.Clone, "InvCode", findStr).Bookmark
        If Err.Number = 3021 Then
            On Error GoTo Errorcode
            rsInvoiceType.Bookmark = rsFind(rsInvoiceType.Clone, "InvType", findStr).Bookmark
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

Private Sub cmdNext_Click()
    rsInvoiceType.MoveNext
    If rsInvoiceType.EOF Then
        rsInvoiceType.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsInvoiceType.MovePrevious
    If rsInvoiceType.BOF Then
        rsInvoiceType.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

'Upating Code       : AXP-0713200713:53
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Print", "INVOICE TYPES") = False Then Exit Sub

    Screen.MousePointer = 11

    rptInvoiceType.Reset
    rptInvoiceType.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInvoiceType.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptInvoiceType.ReportTitle = " INVOICE TYPE "
    PrintSQLReport rptInvoiceType, AMIS_REPORT_PATH & "\files\InvoiceType.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    NEW_LogAudit "V", "INVOICE TYPES", "", labID.Caption, "", txtInvCode, "", ""
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:54
Private Sub cmdSave_Click()
    Dim VtxtInvCode, VtxtInvType                  As String

    On Error GoTo Errorcode:

    VtxtInvCode = N2Str2Null(txtInvCode.Text)
    VtxtInvType = N2Str2Null(txtInvType.Text)

    If AddorEdit = "ADD" Then
        Dim rsInvoiceTypeDup                      As ADODB.Recordset
        Set rsInvoiceTypeDup = New ADODB.Recordset
        rsInvoiceTypeDup.Open "select InvCode from ALL_InvoiceType where InvCode = " & VtxtInvCode, gconDMIS
        If Not rsInvoiceTypeDup.EOF And Not rsInvoiceTypeDup.BOF Then
            MsgBox "Bank Code Already Exist!", vbCritical, "Duplicate Bank Code Not Allowed"
            Exit Sub
        End If
        SQL_STATEMENT = "Insert into ALL_InvoiceType " & _
                        "(InvCode,InvType) " & _
                        " values (" & VtxtInvCode & _
                        ", " & VtxtInvType & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "INVOICE TYPES", SQL_STATEMENT, labID.Caption, "", txtInvCode, "", ""
    Else
        SQL_STATEMENT = "update ALL_InvoiceType set" & _
                        " InvCode = " & VtxtInvCode & ", " & _
                        " InvType = " & VtxtInvType & _
                        " where ID = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "INVOICE TYPES", SQL_STATEMENT, labID.Caption, "", txtInvCode, "", ""
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsInvoiceType.Find "ID = " & labID.Caption
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
        frmALL_AuditInquiry.Caption = "INVOICE TYPES"
        Call frmALL_AuditInquiry.DisplayHistory(labID, "INVOICE TYPES")
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
    Dim rsInvoiceType2                            As ADODB.Recordset
    lstInvoiceType.Enabled = False
    lstInvoiceType.Sorted = False: lstInvoiceType.ListItems.Clear
    Set rsInvoiceType2 = New ADODB.Recordset
    Set rsInvoiceType2 = gconDMIS.Execute("select InvCode,InvType,ID from ALL_InvoiceType")
    If Not (rsInvoiceType2.EOF And rsInvoiceType2.BOF) Then
        lstInvoiceType.Enabled = True
        Listview_Loadval Me.lstInvoiceType.ListItems, rsInvoiceType2
        lstInvoiceType.Refresh
        lstInvoiceType.Enabled = True
    Else
        lstInvoiceType.Enabled = False
    End If

End Sub

Private Sub lstInvoiceType_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstInvoiceType
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

Private Sub lstInvoiceType_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstInvoiceType_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsInvoiceType.Bookmark = rsFind(rsInvoiceType.Clone, "InvCode", Item).Bookmark
    StoreMemVars
End Sub

Private Sub txtInvCode_LostFocus()
    txtInvCode.Text = UCase(txtInvCode.Text)
End Sub


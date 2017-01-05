VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISMASTERFILEATC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATC Codes"
   ClientHeight    =   5115
   ClientLeft      =   1665
   ClientTop       =   1275
   ClientWidth     =   5790
   ForeColor       =   &H00F5F5F5&
   Icon            =   "ATC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5790
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   5625
      TabIndex        =   13
      Top             =   4185
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
         MouseIcon       =   "ATC.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   21
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
         MouseIcon       =   "ATC.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   20
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
         MouseIcon       =   "ATC.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "ATC.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "ATC.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "ATC.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "ATC.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "ATC.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   5625
      Begin VB.TextBox txtRATE 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         Top             =   1050
         Width           =   555
      End
      Begin VB.TextBox txtNATURE 
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
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   630
         Width           =   4245
      End
      Begin VB.TextBox txtATC 
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
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "XX"
         Top             =   180
         Width           =   1665
      End
      Begin Crystal.CrystalReport rptATC 
         Left            =   5100
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "ATC Codes"
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
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1950
         TabIndex        =   9
         Top             =   1110
         Width           =   315
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate of Tax"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ATC Code"
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
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Nature of Income"
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
         Height          =   465
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   1155
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
         Top             =   630
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
         Top             =   630
         Width           =   225
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2685
      Left            =   60
      TabIndex        =   10
      Top             =   1440
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
         TabIndex        =   11
         Top             =   150
         Width           =   5445
      End
      Begin MSComctlLib.ListView lstATC 
         Height          =   2085
         Left            =   60
         TabIndex        =   12
         Top             =   540
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   3678
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
         MouseIcon       =   "ATC.frx":2D71
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ATC CODE"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NATURE OF INCOME"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RATE"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4230
      ScaleHeight     =   885
      ScaleWidth      =   1485
      TabIndex        =   22
      Top             =   4185
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
         MouseIcon       =   "ATC.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":3025
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
         MouseIcon       =   "ATC.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "ATC.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAMISMASTERFILEATC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim THEID                                              As Integer
Dim rsATC                                              As ADODB.Recordset
Dim AddorEdit, PrevATC                                 As String
Attribute PrevATC.VB_VarUserMemId = 1073938434

Sub rsRefresh()
    Set rsATC = New ADODB.Recordset
    Set rsATC = gconDMIS.Execute("select * from AMIS_ATC order by atc asc")


End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtATC.Text = ""
    txtNATURE.Text = ""
    txtRATE.Text = 0
End Sub

Sub StoreMemVars()
    If Not rsATC.EOF And Not rsATC.BOF Then
        Frame1.Enabled = False
        THEID = N2Str2Zero(rsATC!ID)
        txtATC.Text = Null2String(rsATC!ATC)
        txtNATURE.Text = Null2String(rsATC!NATURE)
        txtRATE.Text = N2Str2Zero(rsATC!Rate)
    Else
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsATC2                                         As ADODB.Recordset
    Set rsATC2 = New ADODB.Recordset
    Set rsATC2 = gconDMIS.Execute("select * from AMIS_ATC where ATC = '" & XXX & "'")
    If Not rsATC2.EOF And Not rsATC2.BOF Then
        fraDetails.Enabled = False
        lstATC.Enabled = False
        txtATC.Text = Null2String(rsATC2!ATC)
        txtNATURE.Text = Null2String(rsATC2!NATURE)
        txtRATE.Text = N2Str2Zero(rsATC2!Rate)
    End If
End Sub

Sub FillSearchGrid(XXX As Variant)
    Dim rsATC2                                         As ADODB.Recordset
    lstATC.Sorted = False: lstATC.ListItems.Clear
    Set rsATC2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsATC2 = gconDMIS.Execute("select ATC,NATURE,RATE from AMIS_ATC where NATURE like '" & ReplaceQuote(CStr(XXX)) & "%'")
    If Not (rsATC2.EOF And rsATC2.BOF) Then
        Listview_Loadval Me.lstATC.ListItems, rsATC2
        lstATC.Refresh
        lstATC.Enabled = True
    Else
        lstATC.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "ATC CODES") = False Then Exit Sub

    AddorEdit = "ADD": initMemvars: Picture1.Visible = False: Picture2.Visible = True
    On Error Resume Next
    txtATC.SetFocus
    lstATC.Enabled = False
    txtSearch.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False: StoreMemVars: fraDetails.Enabled = True: lstATC.Enabled = True: FillGrid
    On Error Resume Next
    lstATC.FindItem(txtATC.Text).EnsureVisible
    lstATC.Enabled = True
    txtSearch.Enabled = True
End Sub

'Upating Code       : AXP-0713200713:51
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "ATC CODES") = False Then Exit Sub

    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "delete from AMIS_ATC where ATC = " & N2Str2Null((lstATC.SelectedItem))
        gconDMIS.Execute SQL_STATEMENT
        rsRefresh
        StoreMemVars
        TransactionID = (FindTransactionID(N2Str2Null(txtATC), "ATC", "AMIS_ATC", ""))
        NEW_LogAudit "X", "ATC CODES", SQL_STATEMENT, TransactionID, txtNATURE, txtATC, "", ""
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:51
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "ATC CODES") = False Then Exit Sub

    AddorEdit = "EDIT": Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True:
    Call StoreEntry(labID)
    PrevATC = txtATC.Text
    On Error Resume Next
    txtATC.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200713:51
Private Sub cmdFind_Click()
    On Error Resume Next

    txtSearch.SetFocus
End Sub

'Upating Code       : AXP-0713200713:52
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsATC.MoveNext
    If rsATC.EOF Then
        rsATC.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:52
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsATC.MovePrevious
    If rsATC.BOF Then
        rsATC.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:52
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "ATC CODES") = False Then Exit Sub

    Screen.MousePointer = 11
    rptATC.Reset
    rptATC.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptATC.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptATC.ReportTitle = " ATC CODES"
    PrintSQLReport rptATC, AMIS_REPORT_PATH & "\files\ATC.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    TransactionID = (FindTransactionID(N2Str2Null(txtATC), "ATC", "AMIS_ATC", ""))
    NEW_LogAudit "V", "ATC CODES", "", TransactionID, N2Str2Null(txtNATURE), N2Str2Null(txtATC), "", ""
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:52
Private Sub cmdSave_Click()

    Dim VtxtATC, VtxtNATURE                            As String
    Dim VtxtRATE                                       As Double
    On Error GoTo ErrorCode:

    VtxtATC = N2Str2Null(txtATC.Text)
    VtxtNATURE = N2Str2Null(txtNATURE.Text)
    VtxtRATE = NumericVal(txtRATE.Text)

    If VtxtATC = "NULL" Then
        MessagePop RecSaveError, "Required Field", "ATC CODES is Required Field", 1000
        Exit Sub
    End If
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into AMIS_ATC " & _
                        "(ATC,NATURE,RATE) " & _
                        " values (" & VtxtATC & "," & VtxtNATURE & "," & VtxtRATE & ")"
        gconDMIS.Execute SQL_STATEMENT

        TransactionID = (FindTransactionID(N2Str2Null(VtxtATC), "ATC", "AMIS_ATC", ""))
        NEW_LogAudit "A", "ATC CODES", SQL_STATEMENT, TransactionID, "", N2Str2Null(VtxtATC), "", ""

    Else
        SQL_STATEMENT = "update AMIS_ATC set" & _
                        " ATC = " & VtxtATC & "," & _
                        " NATURE = " & VtxtNATURE & "," & _
                        " RATE = " & VtxtRATE & _
                        " where ATC = '" & PrevATC & "'"

        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(VtxtATC), "ATC", "AMIS_ATC", ""))
        NEW_LogAudit "E", "ATC CODES", SQL_STATEMENT, TransactionID, N2Str2Null(VtxtNATURE), N2Str2Null(VtxtATC), "", ""

    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsATC.Find "ATC = " & VtxtATC

    cmdCancel.Value = True

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyUp(KeyATC As Integer, Shift As Integer)
    MoveKeyPress KeyATC
    Select Case KeyATC
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "ATC CODES"
        Call frmALL_AuditInquiry.DisplayHistory(labID, "ATC CODES")
    End Select
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
    Dim rsATC2                                         As ADODB.Recordset
    lstATC.Sorted = False: lstATC.ListItems.Clear
    Set rsATC2 = New ADODB.Recordset
    Set rsATC2 = gconDMIS.Execute("select ATC,NATURE,RATE from AMIS_ATC order by ATC asc")
    If Not (rsATC2.EOF And rsATC2.BOF) Then
        Listview_Loadval Me.lstATC.ListItems, rsATC2
        lstATC.Refresh
        lstATC.Enabled = True
    Else
        lstATC.Enabled = False
    End If
End Sub

Private Sub lstATC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstATC
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstATC_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstATC_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsATC.Bookmark = rsFind(rsATC.Clone, "ATC", lstATC.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub


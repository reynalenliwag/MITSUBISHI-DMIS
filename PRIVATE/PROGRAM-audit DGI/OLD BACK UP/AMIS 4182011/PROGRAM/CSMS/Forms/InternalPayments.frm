VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSInternalPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVICE PAYMENT METHODS - INTERNAL"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   555
   ClientWidth     =   5535
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "InternalPayments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   5535
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   30
      TabIndex        =   18
      Top             =   30
      Width           =   5355
      Begin VB.ComboBox cboAccountCode 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1410
         Width           =   5205
      End
      Begin VB.TextBox txtAccountCode 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   50
         TabIndex        =   2
         Top             =   990
         Width           =   2475
      End
      Begin VB.TextBox txtCODE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   2
         TabIndex        =   0
         Top             =   210
         Width           =   555
      End
      Begin VB.TextBox txtDESCNAME 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   50
         TabIndex        =   1
         Top             =   570
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   22
         Top             =   1020
         Width           =   1290
      End
      Begin VB.Label labDESCNAME 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   21
         Top             =   600
         Width           =   2850
      End
      Begin VB.Label labCODE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   2115
      End
      Begin VB.Label labid 
         Caption         =   "Label1"
         Height          =   285
         Left            =   3600
         TabIndex        =   19
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3585
      Left            =   30
      TabIndex        =   15
      Top             =   2010
      Width           =   5355
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
         Left            =   60
         MaxLength       =   35
         TabIndex        =   13
         Top             =   150
         Width           =   5205
      End
      Begin MSComctlLib.ListView lstSBook 
         Height          =   2985
         Left            =   60
         TabIndex        =   14
         Top             =   540
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   5265
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
         MouseIcon       =   "InternalPayments.frx":1082
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TRANSACTION"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   300
      ScaleHeight     =   855
      ScaleWidth      =   6600
      TabIndex        =   17
      Top             =   5700
      Width           =   6600
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
         Left            =   4365
         MouseIcon       =   "InternalPayments.frx":11E4
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":1336
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exit Window"
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
         Left            =   3675
         MouseIcon       =   "InternalPayments.frx":169C
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":17EE
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   2985
         MouseIcon       =   "InternalPayments.frx":1B19
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":1C6B
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   2295
         MouseIcon       =   "InternalPayments.frx":1FC7
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":2119
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   1605
         MouseIcon       =   "InternalPayments.frx":242C
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":257E
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   915
         MouseIcon       =   "InternalPayments.frx":2878
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":29CA
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   225
         MouseIcon       =   "InternalPayments.frx":2D22
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":2E74
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   3885
      ScaleHeight     =   885
      ScaleWidth      =   2130
      TabIndex        =   16
      Top             =   5685
      Width           =   2130
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
         Left            =   780
         MouseIcon       =   "InternalPayments.frx":31D3
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":3325
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   90
         MouseIcon       =   "InternalPayments.frx":3663
         MousePointer    =   99  'Custom
         Picture         =   "InternalPayments.frx":37B5
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label LocalAcess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   2115
   End
End
Attribute VB_Name = "frmCSMSInternalPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSBOOK                                            As ADODB.Recordset
Dim ADDOREDIT                                          As String

Function SetAccountCode(XXX As String) As String
    Dim rsChartAccount                                 As New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Description = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function SetAccountDesc(XXX As String) As String
    Dim rsChartAccount                                 As New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where AcctCode = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccountDesc = Null2String(rsChartAccount!Description)
    End If
    Set rsChartAccount = Nothing
End Function

Sub InitCboAccountCode()
    Dim rsChartAccount                                 As New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Order by AcctCode asc")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        rsChartAccount.MoveFirst: cboAccountCode.Clear
        Do While Not rsChartAccount.EOF
            cboAccountCode.AddItem Null2String(rsChartAccount!Description)
            rsChartAccount.MoveNext
        Loop
    End If
    Set rsChartAccount = Nothing
End Sub

Sub InitMemVars()
    txtCode.Text = ""
    txtDESCNAME.Text = ""
    txtAccountCode.Text = ""
    txtAccountCode.Enabled = True
    cboAccountCode.Enabled = True
    cboAccountCode.ListIndex = -1
End Sub

Sub StoreMemvars()
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        labID.Caption = rsSBOOK!ID
        txtCode.Text = Null2String(rsSBOOK!Code)
        txtDESCNAME.Text = Null2String(rsSBOOK!DESCNAME)
        txtAccountCode.Text = Null2String(rsSBOOK!CHARTCODES)
        If SetAccountDesc(Null2String(rsSBOOK!CHARTCODES)) = "" Then
            cboAccountCode.ListIndex = -1
        Else
            cboAccountCode.Text = SetAccountDesc(Null2String(rsSBOOK!CHARTCODES))
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsSBOOK = New ADODB.Recordset
    rsSBOOK.Open "select * from CMIS_CBOOK where BOOK = 'S' order by ID asc", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsSBook2                                       As New ADODB.Recordset
    lstSBook.Sorted = False: lstSBook.ListItems.Clear
    lstSBook.Enabled = False
    Set rsSBook2 = gconDMIS.Execute("select CODE,DESCNAME , id from CMIS_SBOOK WHERE BOOK = 'S'")
    If Not (rsSBook2.EOF And rsSBook2.BOF) Then
        Listview_Loadval Me.lstSBook.ListItems, rsSBook2
        lstSBook.Refresh
        lstSBook.Enabled = True
    End If
End Sub

Sub FillSearchGrid(XXX As Variant)
    Dim rsSBook2                                       As New ADODB.Recordset
    lstSBook.Sorted = False: lstSBook.ListItems.Clear
    lstSBook.Enabled = False
    Set rsSBook2 = gconDMIS.Execute("select CODE,DESCNAME ,id from CMIS_SBOOK where BOOK = 'S' AND DESCNAME like '" & Replace(XXX, "'", "") & "%'")
    If Not (rsSBook2.EOF And rsSBook2.BOF) Then
        Listview_Loadval Me.lstSBook.ListItems, rsSBook2
        lstSBook.Refresh
        lstSBook.Enabled = True
    End If

End Sub

Private Sub cboAccountCode_Change()
    txtAccountCode.Text = SetAccountCode(cboAccountCode)
End Sub

Private Sub cboAccountCode_Click()
    txtAccountCode.Text = SetAccountCode(cboAccountCode)
End Sub

Private Sub cboAccountCode_KeyDown(KeyCode As Integer, Shift As Integer)
    txtAccountCode.Text = SetAccountCode(cboAccountCode)
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "SERVICE PAYMENT METHODS") = False Then: Exit Sub

    ADDOREDIT = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemVars
    lstSBook.Enabled = False
    txtSearch.Enabled = False
    On Error Resume Next
    txtCode.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstSBook.Enabled = True
    txtSearch.Enabled = True
    fraDetails.Enabled = True
    txtSearch.Enabled = True
    lstSBook.Enabled = True
    Call StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "SERVICE PAYMENT METHODS") = False Then Exit Sub
    On Error GoTo Errorcode
    If Not rsSBOOK.BOF Or Not rsSBOOK.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from CMIS_SBOOK where ID = " & labID.Caption
            gconDMIS.Execute SQL_STATEMENT

            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("X", "SERVICE PAYMENT METHODS", SQL_STATEMENT, labID, "", "CODE: " & txtCode, "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            'LogAudit "X", "CODE MAINTENANCE", Me.labDESCNAME.Caption & " " & Me.labCODE
            Call ShowDeletedMsg
        End If
    Else
        Call ShowNoRecord
    End If

    Call rsRefresh
    Call FillGrid
    StoreMemvars
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "SERVICE PAYMENT METHODS") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    txtSearch.Enabled = False
    lstSBook.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsSBOOK.MoveNext
    If rsSBOOK.EOF Then
        rsSBOOK.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsSBOOK.MovePrevious
    If rsSBOOK.BOF Then
        rsSBOOK.MoveFirst
        Call ShowFirstRecordMsg
    End If
    Call StoreMemvars
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    Dim rsfindDup                                       As ADODB.Recordset
    Dim VtxtCode                                        As String
    Dim VTXTDESCNAME                                    As String
    Dim VTXTACCOUNTCODE                                 As String

    If IsNull(txtCode.Text) = True Then
        MsgSpeechBox "Bank Code must not be empty"
        On Error Resume Next
        txtCode.SetFocus
        Exit Sub
    Else
        If ADDOREDIT = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select CODE from CMIS_SBOOK where CODE = 'S'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "SBook Code already exist!"
                On Error Resume Next
                txtCode.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtDESCNAME.Text = "" Then
        MsgSpeechBox "DESCNAME is Required"
        Exit Sub
    End If
    If txtAccountCode.Text = "" Then
        MsgSpeechBox "ACCOUNT CODE is Required"
        Exit Sub
    End If

    VtxtCode = N2Str2Null(txtCode.Text)
    VTXTDESCNAME = N2Str2Null(txtDESCNAME.Text)
    VTXTACCOUNTCODE = N2Str2Null(txtAccountCode.Text)
    If ADDOREDIT = "ADD" Then
        SQL_STATEMENT = "Insert into CMIS_SBook" & _
            " (CODE, DESCNAME, CHARTCODES, BOOK, DATECREATE, WHOCREATE)" & _
            " values (" & VtxtCode & _
            ", " & VTXTDESCNAME & _
            ", " & VTXTACCOUNTCODE & _
            ", 'S' " & _
            ", " & N2Str2Null(LOGDATE) & _
            ", " & N2Str2Null(LOGCODE) & ")"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("A", "SERVICE PAYMENT METHODS", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtCode), "CODE", "CMIS_SBOOK", "DETAILS", "S", "BOOK"), "", "CODE: " & txtCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CMIS_SBook set" & _
            " CODE = " & VtxtCode & "," & _
            " DESCNAME = " & VTXTDESCNAME & "," & _
            " CHARTCODES = " & VTXTACCOUNTCODE & "," & _
            " DATECREATE = " & "'" & LOGDATE & "'" & "," & _
            " WHOCREATE = " & "" & N2Str2Null(LOGCODE) & "" & _
            " where ID = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        'LogAudit "E", "CODE MAINTENANCE", Me.labDESCNAME.Caption & " " & Me.labCODE

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "SERVICE PAYMENT METHODS", SQL_STATEMENT, labID, "", "CODE: " & txtCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyUpdated
    End If
    
    Call rsRefresh
    Call FillGrid
    rsSBOOK.Find "CODE = " & VtxtCode
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub

            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SERVICE PAYMENT METHODS)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "SERVICE PAYMENT METHODS", "")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    
    Call rsRefresh
    Frame1.Enabled = False
    Call FillGrid
    Call InitCboAccountCode
    Call InitMemVars
    Call StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstSBook_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstSBook
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstSBook_DblClick()
    If Not lstSBook.ListItems.Count = 0 Then
        cmdEdit.Value = True
    End If
End Sub

Private Sub lstSBook_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsSBOOK.MoveFirst
    If IsNumeric(lstSBook.SelectedItem) = True Then
        rsSBOOK.Bookmark = rsFind(rsSBOOK.Clone, "CODE", lstSBook.SelectedItem).Bookmark
    Else
        On Error Resume Next
        rsSBOOK.Find ("CODE=" & N2Str2Null(lstSBook.SelectedItem))
    End If
    Call StoreMemvars
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

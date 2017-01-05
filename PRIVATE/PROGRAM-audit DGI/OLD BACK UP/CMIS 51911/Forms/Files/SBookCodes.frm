VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCMISSBookEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSACTION CODES"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   555
   ClientWidth     =   5445
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "SBookCodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   5445
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
         TabIndex        =   3
         Text            =   "cboAccountCode"
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
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1020
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
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   0
         Top             =   210
         Width           =   1395
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
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   3465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   135
         TabIndex        =   22
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label labDESCNAME 
         Alignment       =   1  'Right Justify
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
         Left            =   375
         TabIndex        =   21
         Top             =   660
         Width           =   1050
      End
      Begin VB.Label labCODE 
         Alignment       =   1  'Right Justify
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
         Left            =   930
         TabIndex        =   20
         Top             =   330
         Width           =   495
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
         MouseIcon       =   "SBookCodes.frx":08CA
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
      Top             =   5670
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
         MouseIcon       =   "SBookCodes.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":0B7E
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
         MouseIcon       =   "SBookCodes.frx":0EE4
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":1036
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
         MouseIcon       =   "SBookCodes.frx":1361
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":14B3
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
         MouseIcon       =   "SBookCodes.frx":180F
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":1961
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
         MouseIcon       =   "SBookCodes.frx":1C74
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":1DC6
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
         MouseIcon       =   "SBookCodes.frx":20C0
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":2212
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
         MouseIcon       =   "SBookCodes.frx":256A
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":26BC
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
         MouseIcon       =   "SBookCodes.frx":2A1B
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":2B6D
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
         MouseIcon       =   "SBookCodes.frx":2EAB
         MousePointer    =   99  'Custom
         Picture         =   "SBookCodes.frx":2FFD
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
Attribute VB_Name = "frmCMISSBookEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSBOOK                                                           As ADODB.Recordset
Dim AddorEdit                                                         As String

Function SetAccountCode(XXX As String) As String
    Dim rsChartAccount                                                As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Description = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function SetAccountDesc(XXX As String) As String
    Dim rsChartAccount                                                As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where AcctCode = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccountDesc = Null2String(rsChartAccount!Description)
    End If
    Set rsChartAccount = Nothing
End Function

Sub InitCboAccountCode()
    Dim rsChartAccount                                                As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    'Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where (HeaderCode = '2' OR HeaderCode = '4' OR HeaderCode = '7' OR HeaderCode = '8' OR HeaderCode = '9') Order by AcctCode asc")
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

Sub initMemvars()
    txtCode.Text = ""
    txtDESCNAME.Text = ""
    txtAccountCode.Text = ""
    If BOOKTYPE = "D" Then
        txtAccountCode.Enabled = True
        cboAccountCode.Enabled = True
    Else
        txtAccountCode.Enabled = False
        cboAccountCode.Enabled = False
    End If
End Sub

Sub StoreMemVars()
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        labid.Caption = rsSBOOK!Id
        txtCode.Text = Null2String(rsSBOOK!code)
        txtDESCNAME.Text = Null2String(rsSBOOK!DESCNAME)
        If BOOKTYPE = "I" Then
        Else
            txtAccountCode.Text = Null2String(rsSBOOK!CHARTCODES)
            cboAccountCode.Text = SetAccountDesc(Null2String(rsSBOOK!CHARTCODES))
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsSBOOK = New ADODB.Recordset
    If BOOKTYPE = "J" Then
        rsSBOOK.Open "select * from CMIS_CBOOK where BOOK = '" & BOOKTYPE & "' order by ID asc", gconDMIS, adOpenKeyset, adLockReadOnly
    Else
        If BOOKTYPE = "I" Then
            rsSBOOK.Open "select * from CMIS_vw_Vemployee order by DESCNAME asc", gconDMIS, adOpenKeyset, adLockReadOnly
        Else
            rsSBOOK.Open "select * from CMIS_CBOOK where BOOK = '" & BOOKTYPE & "' order by ID asc", gconDMIS, adOpenKeyset, adLockReadOnly
        End If
    End If
End Sub

Sub FillGrid()
    Dim rsSBook2                                                      As ADODB.Recordset
    lstSBook.Sorted = False: lstSBook.ListItems.Clear
    lstSBook.Enabled = False
    Set rsSBook2 = New ADODB.Recordset
    If BOOKTYPE = "J" Then
        Set rsSBook2 = gconDMIS.Execute("select CODE,DESCNAME ,id from CMIS_CBOOK where BOOK = '" & BOOKTYPE & "' order by CHARTCODES asc")
    Else
        If BOOKTYPE = "I" Then
            Set rsSBook2 = gconDMIS.Execute("select CODE,DESCNAME , id from CMIS_vw_Vemployee")
        Else
            Set rsSBook2 = gconDMIS.Execute("select CODE,DESCNAME , id from CMIS_SBOOK WHERE BOOK = '" & BOOKTYPE & "'")
        End If
    End If
    If Not (rsSBook2.EOF And rsSBook2.BOF) Then
        Listview_Loadval Me.lstSBook.ListItems, rsSBook2
        lstSBook.Refresh
        lstSBook.Enabled = True
    End If

End Sub

Sub FillSearchGrid(XXX As Variant)
    Dim rsSBook2                                                      As ADODB.Recordset
    lstSBook.Sorted = False: lstSBook.ListItems.Clear
    lstSBook.Enabled = False
    Set rsSBook2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    If BOOKTYPE = "J" Then
        Set rsSBook2 = gconDMIS.Execute("select CODE,DESCNAME ,id from CMIS_CBOOK where BOOK = '" & BOOKTYPE & "' AND DESCNAME like '" & XXX & "%'")
    Else
        If BOOKTYPE = "I" Then
            Set rsSBook2 = gconDMIS.Execute("select CODE,DESCNAME,id  from CMIS_vw_Vemployee where DESCNAME like '" & XXX & "%'")
        Else
            Set rsSBook2 = gconDMIS.Execute("select CODE,DESCNAME ,id from CMIS_SBOOK where BOOK = '" & BOOKTYPE & "' AND DESCNAME like '" & XXX & "%'")
        End If
    End If
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
    If Function_Access(LOGID, "Acess_Add", LocalAcess) = False Then: Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
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
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", LocalAcess) = False Then Exit Sub

    On Error GoTo Errorcode
    If Not rsSBOOK.BOF Or Not rsSBOOK.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from CMIS_SBOOK where ID = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'LogAudit "X", "CODE MAINTENANCE", "CODE: " & Me.txtCODE & ", DESCRIPTION: " & Me.txtDESCNAME
            Call NEW_LogAudit("X", LocalAcess, SQL_STATEMENT, labid, "", "CODE :" & txtCode, "", "")
            
            ShowDeletedMsg
        End If
    Else
        MsgSpeechBox "Nothing to delete!"
    End If

    rsRefresh
    FillGrid
    StoreMemVars
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LocalAcess) = False Then Exit Sub
    AddorEdit = "EDIT"
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
        ShowLastRecordMsg
        'MessagePop NaviEnd, "End of Record", "Last Record"
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSBOOK.MovePrevious
    If rsSBOOK.BOF Then
        rsSBOOK.MoveFirst
        ShowFirstRecordMsg
        'MsgSpeechBox "Beginning of Record"
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo ErrorCode
    Dim rsfindDup                                                     As ADODB.Recordset
    Dim OCODE As String: Dim ODESC                                    As String

    Dim VTXTCODE As String
    Dim VTXTDESCNAME As String
    Dim VTXTACCOUNTCODE                       As String
    
'    OCODE = Null2String(rsSBOOK!code)
'    ODESC = Null2String(rsSBOOK!DESCNAME)

    If IsNull(txtCode.Text) = True Then
        MsgSpeechBox "Bank Code must not be empty"
        On Error Resume Next
        txtCode.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select CODE from CMIS_SBOOK where CODE = '" & txtCode.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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

    VTXTCODE = N2Str2Null(txtCode.Text)
    VTXTDESCNAME = N2Str2Null(txtDESCNAME.Text)
    If BOOKTYPE = "D" Then
        VTXTACCOUNTCODE = N2Str2Null(txtAccountCode.Text)
    Else
        VTXTACCOUNTCODE = "NULL"
    End If
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into CMIS_SBook" & _
                       " (CODE,DESCNAME,CHARTCODES,BOOK,DATECREATE,WHOCREATE)" & _
                       " values (" & VTXTCODE & ", " & VTXTDESCNAME & "," & VTXTACCOUNTCODE & ", '" & BOOKTYPE & "','" & LOGDATE & "'" & ", " & N2Str2Null(LOGCODE) & ")"
        gconDMIS.Execute SQL_STATEMENT
    
        If BOOKTYPE = "I" Then
            'NEW LOG AUDIT-------------------------------------------------------------------
                Call NEW_LogAudit("A", LocalAcess, SQL_STATEMENT, FindTransactionID(N2Str2Null(txtCode), "CODE", "CMIS_VW_VEMPLOYEE"), "", "CODE: " & txtCode, "", "")
            'NEW LOG AUDIT-------------------------------------------------------------------
        Else
            'NEW LOG AUDIT-------------------------------------------------------------------
                Call NEW_LogAudit("A", LocalAcess, SQL_STATEMENT, FindTransactionID(N2Str2Null(txtCode), "CODE", "CMIS_SBOOK", "DETAILS", N2Str2Null(BOOKTYPE), "BOOK"), "", "CODE: " & txtCode, "", "")
            'NEW LOG AUDIT-------------------------------------------------------------------
        End If
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CMIS_SBook set" & _
                       " CODE = " & VTXTCODE & "," & _
                       " DESCNAME = " & VTXTDESCNAME & "," & _
                       " CHARTCODES = " & VTXTACCOUNTCODE & "," & _
                       " DATECREATE = " & "'" & LOGDATE & "'" & "," & _
                       " WHOCREATE = " & "" & N2Str2Null(LOGCODE) & "" & _
                       " where ID = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-------------------------------------------------------------------
            Call NEW_LogAudit("E", LocalAcess, SQL_STATEMENT, labid, "", "CODE: " & txtCode, "", "")
        'NEW LOG AUDIT-------------------------------------------------------------------
        ShowSuccessFullyUpdated
    End If

    rsRefresh
    FillGrid
    On Error Resume Next

    rsSBOOK.Find "BOOK= " & BOOKTYPE & "' AND CODE = " & VTXTCODE
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
            If Picture1.Visible = True Then
                Unload frmALL_AuditInquiry
                 
                frmALL_AuditInquiry.Show
                frmALL_AuditInquiry.ZOrder 0
                frmALL_AuditInquiry.Caption = "Audit Inquiry (" & LocalAcess & ")"
                Call frmALL_AuditInquiry.DisplayHistory(labid, LocalAcess)
            End If
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Frame1.Enabled = False
    FillGrid
    InitCboAccountCode
    initMemvars
    StoreMemVars
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

Private Sub lstSBook_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsSBOOK.MoveFirst
    If IsNumeric(lstSBook.SelectedItem) = True Then
        rsSBOOK.Bookmark = rsFind(rsSBOOK.Clone, "CODE", lstSBook.SelectedItem).Bookmark
        ' rsSBOOK.Find ("ID=" & Item.ListSubItems(3).Text)
    Else
        On Error Resume Next
        'rsSBOOK.Bookmark = rsFind(rsSBOOK.Clone, "CODE", Trim(lstSBook.SelectedItem)).Bookmark
        rsSBOOK.Find ("CODE=" & N2Str2Null(lstSBook.SelectedItem))
    End If
    'rsSBOOK.Find ("ID=" & Item.ListSubItems(2).Text)
    StoreMemVars
End Sub

Private Sub txtSEARCH_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub


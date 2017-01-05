VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmCreditCardCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card Terminal"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   ForeColor       =   &H00E0E0E0&
   Icon            =   "CreditCardCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   5535
   Begin VB.Frame fraDetails 
      Height          =   3195
      Left            =   90
      TabIndex        =   0
      Top             =   1950
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
         TabIndex        =   1
         Top             =   150
         Width           =   5205
      End
      Begin MSComctlLib.ListView lvBank 
         Height          =   2550
         Left            =   30
         TabIndex        =   26
         Top             =   600
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   4498
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
         MouseIcon       =   "CreditCardCompany.frx":09AA
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bank"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Bank Charges"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "EWT"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   60
      ScaleHeight     =   855
      ScaleWidth      =   5400
      TabIndex        =   3
      Top             =   5220
      Width           =   5400
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4680
         MouseIcon       =   "CreditCardCompany.frx":0B0C
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":0C5E
         Style           =   1  'Graphical
         TabIndex        =   4
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3960
         MouseIcon       =   "CreditCardCompany.frx":0FC4
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   5
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3240
         MouseIcon       =   "CreditCardCompany.frx":1441
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":1593
         Style           =   1  'Graphical
         TabIndex        =   2
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2520
         MouseIcon       =   "CreditCardCompany.frx":18EF
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":1A41
         Style           =   1  'Graphical
         TabIndex        =   6
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1800
         MouseIcon       =   "CreditCardCompany.frx":1D54
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":1EA6
         Style           =   1  'Graphical
         TabIndex        =   7
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1080
         MouseIcon       =   "CreditCardCompany.frx":21A0
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":22F2
         Style           =   1  'Graphical
         TabIndex        =   8
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   360
         MouseIcon       =   "CreditCardCompany.frx":264A
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":279C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1875
      Left            =   90
      TabIndex        =   13
      Top             =   90
      Width           =   5355
      Begin VB.TextBox txtEWT 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   22
         Top             =   1320
         Width           =   705
      End
      Begin VB.TextBox txtBankCharges 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   20
         Top             =   1320
         Width           =   705
      End
      Begin VB.TextBox cboCUSNAME 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   1560
         TabIndex        =   16
         Top             =   840
         Width           =   3645
      End
      Begin VB.TextBox txtCUSCDE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   15
         Top             =   360
         Width           =   945
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   14
         Top             =   380
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3840
         TabIndex        =   25
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   2280
         TabIndex        =   24
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EWT"
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
         Index           =   3
         Left            =   2640
         TabIndex        =   23
         Top             =   1400
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Charges"
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
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1400
         Width           =   1380
      End
      Begin VB.Label lblID 
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Code"
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
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   400
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   900
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   5400
      TabIndex        =   10
      Top             =   5220
      Width           =   5400
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4680
         MouseIcon       =   "CreditCardCompany.frx":2AFB
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":2C4D
         Style           =   1  'Graphical
         TabIndex        =   11
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3960
         MouseIcon       =   "CreditCardCompany.frx":2F8B
         MousePointer    =   99  'Custom
         Picture         =   "CreditCardCompany.frx":30DD
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCreditCardCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCreditCardBank                                                As ADODB.Recordset
Dim AddorEdit                                                       As String

Private Sub cmdFind_Click()
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    
    rsCreditCardBank.MoveNext
    If rsCreditCardBank.EOF Then
        rsCreditCardBank.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsCreditCardBank.MovePrevious
    If rsCreditCardBank.BOF Then
        rsCreditCardBank.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSelect_Click()
    SelectCustomer = "Terminal"
    frmCustomerSearch1.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False
    rsRefresh
    FillGrid
End Sub

Private Sub cmdAdd_Click()
    'If Function_Access(LOGID, "Acess_Add", LocalAcess) = False Then: Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    lvBank.Enabled = False
    txtSearch.Enabled = False
    On Error Resume Next
    'txtCODE.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    lvBank.Enabled = True
    txtSearch.Enabled = True
    fraDetails.Enabled = True
    txtSearch.Enabled = True
    lvBank.Enabled = True
'    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    'If Function_Access(LOGID, "Acess_Delete", LocalAcess) = False Then Exit Sub
    On Error GoTo ErrorCode
    
    If txtCUSCDE.Text = "" Then
        MessagePop InfoWarning, "INFORMATION", "Select record to delete"
        Exit Sub
    Else
        If Not rsCreditCardBank.BOF Or Not rsCreditCardBank.EOF Then
            If ShowConfirmDelete = True Then
                SQL_STATEMENT = "delete from CMIS_CardBank where ID = " & lblID.Caption
                gconDMIS.Execute SQL_STATEMENT
    
                'LogAudit "X", "CODE MAINTENANCE", "CODE: " & Me.txtCODE & ", DESCRIPTION: " & Me.txtDESCNAME
                'Call NEW_LogAudit("X", LocalAcess, SQL_STATEMENT, LABID, "", "CODE :" & txtCODE, "", "")

                ShowDeletedMsg
            Else
                Exit Sub
            End If
        Else
            MsgSpeechBox "No selected record to Delete!"
        End If
    End If
    
    rsRefresh
    FillGrid
    StoreMemVars
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    'If Function_Access(LOGID, "Acess_Edit", LocalAcess) = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    txtSearch.Enabled = False
    lvBank.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    Dim vCUSCDE                         As String
    Dim vCusName                        As String
    Dim vBankCharges                    As Double
    Dim vEWT                            As Double
    Dim rsFindBank                      As ADODB.Recordset

    vCUSCDE = N2Str2Null(txtCUSCDE)
    vCusName = N2Str2Null(cboCUSNAME)
    vBankCharges = NumericVal(txtBankCharges)
    vEWT = NumericVal(txtEWT)

    If AddorEdit = "ADD" Then
        Set rsFindBank = New ADODB.Recordset
        rsFindBank.Open "Select * from CMIS_CardBank where CusCde = " & vCUSCDE & "", gconDMIS, adOpenForwardOnly
        If Not rsFindBank.EOF And Not rsFindBank.BOF Then
            MsgBox "Bank Code already exists", vbExclamation, "Check Bank"
            On Error Resume Next
            txtCUSCDE.SetFocus
        Else
            If txtCUSCDE.Text = "" Then
                MsgBox "Please fill the missing entry.", vbInformation, "Empty field"
                txtCUSCDE.SetFocus
                Exit Sub
            ElseIf cboCUSNAME.Text = "" Then
                MsgBox "Please fill the missing entry.", vbInformation, "Empty field"
                cboCUSNAME.SetFocus
                Exit Sub
            ElseIf txtBankCharges.Text = "" Then
                MsgBox "Please fill the missing entry.", vbInformation, "Empty field"
                txtBankCharges.SetFocus
                Exit Sub
            ElseIf txtBankCharges.Text = "" Then
                MsgBox "Please fill missing entry.", vbInformation, "Empty field"
                txtEWT.SetFocus
                Exit Sub
            End If
            gconDMIS.Execute ("Insert into CMIS_CardBank (CusCde,AcctName,BankCharges,EWT)Values (" & vCUSCDE & "," & vCusName & "," & vBankCharges & "," & vEWT & ")")
            ShowSuccessFullyAdded
        End If
    Else
        gconDMIS.Execute ("Update CMIS_CardBank Set CusCde =" & vCUSCDE & ",  AcctName =" & vCusName & ", BankCharges=" & vBankCharges & ", EWT=" & vEWT & " where ID = '" & lblID.Caption & "'")
        ShowSuccessFullyUpdated
    End If
    
    FillGrid
    cmdCancel.Value = True
    
ErrorCode:
        ShowVBError
        Exit Sub
    End Sub

Sub initMemvars()
    txtCUSCDE.Text = ""
    cboCUSNAME.Text = ""
    txtBankCharges.Text = ""
    txtEWT.Text = ""
End Sub

Sub FillGrid()
    Dim xList                                                       As ListItem
    
    lvBank.ListItems.Clear
    Set rsCreditCardBank = New ADODB.Recordset
    rsCreditCardBank.Open "Select * from CMIS_CardBank order by CusCde,AcctName", gconDMIS, adOpenForwardOnly
    If Not rsCreditCardBank.EOF And Not rsCreditCardBank.BOF Then
        Do While Not rsCreditCardBank.EOF
            Set xList = lvBank.ListItems.Add(, , Null2String(rsCreditCardBank!CUSCDE))
                        xList.SubItems(1) = Null2String(rsCreditCardBank!AcctName)
                        xList.SubItems(2) = NumericVal(rsCreditCardBank!Bankcharges)
                        xList.SubItems(3) = NumericVal(rsCreditCardBank!EWT)
            rsCreditCardBank.MoveNext
        Loop
    End If
End Sub

Private Sub lvBank_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set rsCreditCardBank = New ADODB.Recordset
    rsCreditCardBank.Open "Select * from CMIS_CardBank where Cuscde ='" & lvBank.SelectedItem.Text & "'", gconDMIS, adOpenForwardOnly
    If Not rsCreditCardBank.EOF And Not rsCreditCardBank.BOF Then
        lblID.Caption = Null2String(rsCreditCardBank!Id)
        txtCUSCDE = Null2String(rsCreditCardBank!CUSCDE)
        cboCUSNAME = Null2String(rsCreditCardBank!AcctName)
        txtBankCharges = Null2String(rsCreditCardBank!Bankcharges)
        txtEWT = Null2String(rsCreditCardBank!EWT)
    End If
End Sub

Sub StoreMemVars()
    If Not rsCreditCardBank.EOF And Not rsCreditCardBank.BOF Then
        lblID.Caption = Null2String(rsCreditCardBank!Id)
        txtCUSCDE = Null2String(rsCreditCardBank!CUSCDE)
        cboCUSNAME = Null2String(rsCreditCardBank!AcctName)
        txtBankCharges = NumericVal(txtBankCharges)
        txtEWT = NumericVal(txtEWT)
    'Else
    '    ShowNoRecord
    '    cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsCreditCardBank = New ADODB.Recordset
    rsCreditCardBank.Open "Select * from CMIS_CardBank", gconDMIS, adOpenForwardOnly
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim xList                                                       As ListItem
    
    Set rsCreditCardBank = New ADODB.Recordset
    rsCreditCardBank.Open "Select * from CMIS_CardBank where CusCde like '" & XXX & "%'", gconDMIS, adOpenKeyset, adLockReadOnly
    lvBank.ListItems.Clear
    If Not rsCreditCardBank.EOF And Not rsCreditCardBank.BOF Then
        Do While Not rsCreditCardBank.EOF
            Set xList = lvBank.ListItems.Add(, , Null2String(rsCreditCardBank!CUSCDE))
                        xList.SubItems(1) = Null2String(rsCreditCardBank!AcctName)
            rsCreditCardBank.MoveNext
        Loop
    End If
End Sub

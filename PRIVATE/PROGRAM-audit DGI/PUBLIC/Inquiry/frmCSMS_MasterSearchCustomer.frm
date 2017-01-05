VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMS_MasterSearchCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_MasterSearchCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   915
      Left            =   6390
      MouseIcon       =   "frmCSMS_MasterSearchCustomer.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_MasterSearchCustomer.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancel"
      Top             =   5610
      Width           =   855
   End
   Begin VB.TextBox textSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      TabIndex        =   3
      Top             =   360
      Width           =   5685
   End
   Begin VB.OptionButton optFN 
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3330
      TabIndex        =   1
      Top             =   60
      Width           =   1305
   End
   Begin VB.OptionButton optLN 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Top             =   60
      Width           =   1305
   End
   Begin VB.OptionButton optFullName 
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5220
      TabIndex        =   2
      Top             =   60
      Value           =   -1  'True
      Width           =   1665
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5190
      Width           =   1545
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5190
      Width           =   5655
   End
   Begin MSComctlLib.ListView lstCustomer 
      Height          =   4395
      Left            =   30
      TabIndex        =   4
      Top             =   750
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   7752
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCSMS_MasterSearchCustomer.frx":1512
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CODE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "   Last Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "   First Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "   Account Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "    Mobile"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   " Home Phone"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "   Fax"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "     Address"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "    City"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "    Province"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   915
      Left            =   5550
      MouseIcon       =   "frmCSMS_MasterSearchCustomer.frx":1674
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_MasterSearchCustomer.frx":17C6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Select this Customer"
      Top             =   5610
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Customer"
      Height          =   915
      Left            =   4710
      MouseIcon       =   "frmCSMS_MasterSearchCustomer.frx":1B02
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_MasterSearchCustomer.frx":1C54
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Add Customer"
      Top             =   5610
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   90
      TabIndex        =   10
      Top             =   420
      Width           =   1350
   End
End
Attribute VB_Name = "frmCSMS_MasterSearchCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event SelectionMade(ByVal xCode As String, xName As String, FromForm As String)
Public Event NoSelectionMade()
Dim xFromFrom                                           As String

Sub PassVariable(xFromForm As String)
    xFromFrom = xFromForm
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsCustomer                                     As New ADODB.Recordset
    
    lstCustomer.Enabled = False
    lstCustomer.Sorted = False
    lstCustomer.ListItems.Clear

    XXX = Replace(LTrim(RTrim(XXX)), "'", "")
    If XXX = "" Then
        Set rsCustomer = gconDMIS.Execute("select TOP 100 CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where cuscde <> '999999' order by lastname asc")
    Else
        If optLN.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select TOP 100 CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where LastName like '" & XXX & "%' order by lastname asc")
        ElseIf optFN.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select TOP 100 CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where FirstName like '" & XXX & "%' order by firstname asc")
        ElseIf optFullName.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select TOP 100 CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where AcctName like '" & XXX & "%' order by AcctName asc")
        End If
    End If
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        lstCustomer.Enabled = True
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
        lstCustomer.Refresh
    Else
        txtCode.Text = "": txtName.Text = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    'COMMENT BY  : MJP07222009 0259PM
    'DESCRIPTION : TO UNIFY THE SEARCHING TOOL
        'On Error GoTo ERRORCODE:
        'If Trim(txtCode.Text) <> "" Then
        '    frmCSMSEditRO.txtAcct_No = txtCode.Text
        '    frmCSMSEditRO.SetCustomer
        '    Unload Me
        'End If
        'Exit Sub
    'COMMENT BY  : MJP07222009 0259PM
    
    If txtCode.Text = "" Then
        MsgBox "Select first a customer", vbInformation, "Info"
        Exit Sub
    End If
    
    RaiseEvent SelectionMade(txtCode.Text, txtName.Text, xFromFrom)
    Exit Sub
ERRORCODE:
    ShowVBError
End Sub

Private Sub cmdAdd_Click()
    If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub

    frmAllCustomer.cmdAdd.Value = True
    frmAllCustomer.Show 1
    Call textSearch_Change
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    'optFullName.Value = True
    Call textSearch_Change
End Sub

Private Sub lstCustomer_DblClick()
    If lstCustomer.ListItems.Count = 0 Then Exit Sub
    
    Call cmdSelect_Click
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtCode.Text = lstCustomer.SelectedItem
    txtName.Text = lstCustomer.SelectedItem.SubItems(3)
End Sub

Private Sub textSearch_Change()
    Call FillSearchGrid(textSearch.Text)
End Sub

Private Sub lstCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        If lstCustomer.ListItems.Count = 0 Then Exit Sub
        Call cmdSelect_Click
    End If
End Sub

Private Sub textSearch_GotFocus()
    textSearch.BackColor = &HC0FFC0
End Sub

Private Sub textSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstCustomer.ListItems.Count > 0 And lstCustomer.Enabled = True Then: lstCustomer.SetFocus

    ElseIf KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_LostFocus()
    textSearch.BackColor = vbWhite
End Sub

Private Sub txtCode_Change()
    If txtCode.Text <> "" Then
        cmdSelect.Enabled = True
    Else
        cmdSelect.Enabled = False
    End If
End Sub

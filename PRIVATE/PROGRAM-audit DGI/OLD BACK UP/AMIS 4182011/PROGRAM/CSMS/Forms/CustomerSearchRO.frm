VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerSearchRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer"
   ClientHeight    =   6525
   ClientLeft      =   450
   ClientTop       =   645
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "CustomerSearchRO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Height          =   735
      Left            =   60
      MouseIcon       =   "CustomerSearchRO.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "CustomerSearchRO.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Add Customer "
      Top             =   5700
      Width           =   735
   End
   Begin VB.TextBox txtActiveForm 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Text            =   "txtActiveForm"
      Top             =   60
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
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
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5250
      Width           =   5655
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H8000000F&
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
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5250
      Width           =   1545
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
      Left            =   4470
      TabIndex        =   4
      Top             =   120
      Width           =   1665
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
      Left            =   1650
      TabIndex        =   3
      Top             =   120
      Width           =   1305
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
      Left            =   3060
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
   Begin VB.TextBox textSearch 
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
      Left            =   1620
      TabIndex        =   0
      Top             =   420
      Width           =   5685
   End
   Begin MSComctlLib.ListView lstCustomer 
      Height          =   4395
      Left            =   60
      TabIndex        =   1
      Top             =   810
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "CustomerSearchRO.frx":076F
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   6480
      MouseIcon       =   "CustomerSearchRO.frx":08D1
      MousePointer    =   99  'Custom
      Picture         =   "CustomerSearchRO.frx":0A23
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel"
      Top             =   5700
      Width           =   735
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   735
      Left            =   5760
      MouseIcon       =   "CustomerSearchRO.frx":0D61
      MousePointer    =   99  'Custom
      Picture         =   "CustomerSearchRO.frx":0EB3
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Select Customer"
      Top             =   5700
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Add Customer"
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
      Left            =   870
      TabIndex        =   12
      Top             =   6000
      Width           =   1215
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
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1350
   End
End
Attribute VB_Name = "frmCustomerSearchRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAddRepor                                         As ADODB.Recordset

Function SetAddress(XXX As String) As String
    Dim rsCustomer                                     As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select CustomerAdd + ' ' + City as CusAdd from ALL_Customer Where CUSCDE = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetAddress = Null2String(rsCustomer!Cusadd)
    End If
End Function

Sub FillGrid()
    Dim rsCustomer                                     As New ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomer = gconDMIS.Execute("select TOP 100 CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer order by lastname asc")
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
        lstCustomer.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsCustomer                                     As New ADODB.Recordset
    
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    XXX = Replace(LTrim(RTrim(XXX)), "'", "")
    
    If optLN.Value = True Then
        Set rsCustomer = gconDMIS.Execute("select TOP 100 CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where LastName like '" & XXX & "%' order by lastname asc")
    ElseIf optFN.Value = True Then
        Set rsCustomer = gconDMIS.Execute("select TOP 100 CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where FirstName like '" & XXX & "%' order by firstname asc")
    ElseIf optFullName.Value = True Then
        Set rsCustomer = gconDMIS.Execute("select TOP 100 CusCde,LastName,FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where AcctName like '" & XXX & "%' order by AcctName asc")
    End If
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
        lstCustomer.Refresh
    End If
End Sub

Private Sub cmdCancel_Click()
    With frmCSMSDataEntry
        .Enabled = True
    End With
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If Trim(txtCode.Text) <> "" Then
        With frmCSMSDataEntry
            .txtParticipat.Text = txtCode.Text
            .txtParticipation.Text = txtName.Text
            .Enabled = True
        End With
        Unload Me
    End If
End Sub

Private Sub cmdAdd_Click()
    frmAllCustomer.cmdAdd.Value = True
    frmAllCustomer.Show 1
    FillGrid
End Sub

Private Sub Form_Load()
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    optFullName.Value = True
    FillGrid
End Sub

Private Sub lstCustomer_DblClick()
    If Not lstCustomer.ListItems.Count = 0 Then
        Call cmdSelect_Click
    End If
End Sub

Private Sub lstCustomer_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    txtCode = lstCustomer.SelectedItem
    txtName = lstCustomer.SelectedItem.SubItems(3)
End Sub

Private Sub lstCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not lstCustomer.ListItems.Count = 0 Then
            Call cmdSelect_Click
        End If
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        Call FillGrid
    Else
        Call FillSearchGrid(textSearch.Text)
    End If
End Sub

Private Sub lstCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then

        If lstCustomer.ListItems.Count > 0 And lstCustomer.Enabled = True Then
            lstCustomer.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

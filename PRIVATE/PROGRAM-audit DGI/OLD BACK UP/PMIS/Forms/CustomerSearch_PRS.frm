VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCustomerSearch_PRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ForeColor       =   &H8000000F&
   Icon            =   "CustomerSearch_PRS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   4605
      TabIndex        =   12
      Top             =   5700
      Visible         =   0   'False
      Width           =   4635
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"CustomerSearch_PRS.frx":030A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   825
         Left            =   60
         TabIndex        =   13
         Top             =   30
         Width           =   4515
      End
   End
   Begin VB.TextBox txtActiveForm 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Text            =   "txtActiveForm"
      Top             =   -510
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5250
      Width           =   5655
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
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
         Size            =   9
         Charset         =   0
         Weight          =   400
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
         Size            =   9
         Charset         =   0
         Weight          =   400
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
         Size            =   9
         Charset         =   0
         Weight          =   400
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
         Size            =   9.75
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "CustomerSearch_PRS.frx":0396
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6480
      MouseIcon       =   "CustomerSearch_PRS.frx":04F8
      MousePointer    =   99  'Custom
      Picture         =   "CustomerSearch_PRS.frx":064A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add Customer"
      Top             =   5700
      Width           =   855
   End
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
      Height          =   915
      Left            =   5640
      MouseIcon       =   "CustomerSearch_PRS.frx":095D
      MousePointer    =   99  'Custom
      Picture         =   "CustomerSearch_PRS.frx":0AAF
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel"
      Top             =   5700
      Width           =   855
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   4800
      MouseIcon       =   "CustomerSearch_PRS.frx":0DED
      MousePointer    =   99  'Custom
      Picture         =   "CustomerSearch_PRS.frx":0F3F
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Select this Customer"
      Top             =   5700
      Width           =   855
   End
   Begin VB.Label LABTYPE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   2355
   End
End
Attribute VB_Name = "frmCustomerSearch_PRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub FillGrid()
    Dim RSCUSTOMER                                     As ADODB.Recordset

    lstCustomer.Enabled = False: lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set RSCUSTOMER = New ADODB.Recordset
    '   If COUNTERTYPE = "CHG" Then
    '      Set rsCUSTOMER = gconDMIS.Execute("select CusCde,UPPER(LastName) AS LASTNAME,UPPER(FirstName) AS FIRSTNAME,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where cuscde <> '999999' and isnull(CREDITLIMIT,0)>0 order by lastname asc")
    '  Else
    Set RSCUSTOMER = gconDMIS.Execute("select top 1000 CusCde,UPPER(LastName) AS LASTNAME,UPPER(FirstName) AS FIRSTNAME,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where cuscde <> '999999' order by lastname asc")
    '  End If
    'john
    
    If Not (RSCUSTOMER.EOF And RSCUSTOMER.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, RSCUSTOMER
        lstCustomer.Enabled = True
        lstCustomer.Refresh
    Else
        txtCode.Text = "": txtName.Text = ""
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSCUSTOMER                                     As ADODB.Recordset
    Dim SEARCHFILTER                                   As String
    lstCustomer.Enabled = False: lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set RSCUSTOMER = New ADODB.Recordset
    'XXX = Repleys(LTrim(RTrim(XXX)))
    SEARCHFILTER = ""
    'If COUNTERTYPE = "CHG" Then
    '    SEARCHFILTER = " and isnull(CREDITLIMIT,0)>0 "
    'End If
    If optLN.Value = True Then
        Set RSCUSTOMER = gconDMIS.Execute("select top 1000 CusCde,UPPER(LastName) LastName,UPPER(FirstName) FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where LastName like '" & XXX & "%'" & SEARCHFILTER & " order by lastname asc")
    ElseIf optFN.Value = True Then
        Set RSCUSTOMER = gconDMIS.Execute("select top 1000 CusCde,UPPER(LastName) LastName,UPPER(FirstName) FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where FirstName like '" & XXX & "%'" & SEARCHFILTER & " order by firstname asc")
    ElseIf optFullName.Value = True Then
        Set RSCUSTOMER = gconDMIS.Execute("select  top 1000 CusCde,UPPER(LastName)LastName,UPPER(FirstName)FirstName,AcctName,Mobile,HomePhone,Fax,CustomerAdd,City,ProvincialAdd from ALL_Customer where AcctName like '" & XXX & "%'" & SEARCHFILTER & " order by AcctName asc")
    End If
    If Not (RSCUSTOMER.EOF And RSCUSTOMER.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, RSCUSTOMER
        lstCustomer.Enabled = True
        lstCustomer.Refresh
    Else
        txtCode.Text = "": txtName.Text = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If Trim(txtCode.Text) <> "" Then
        Select Case LABTYPE.Caption
            Case "P"
                frmPMISTrans_PrisForms.txtCustCode.Text = txtCode.Text
                frmPMISTrans_PrisForms.SetCustomer
                Unload Me
            Case "A"
                frmPMISAC_ARISForms.txtCustCode.Text = txtCode.Text
                frmPMISAC_ARISForms.SetCustomer
                Unload Me

            Case "M"
                frmPMISMAT_MRISForms.txtCustCode.Text = txtCode.Text
                frmPMISMAT_MRISForms.SetCustomer
                Unload Me
        End Select
    End If
End Sub

Private Sub cmdAdd_Click()
    frmAllCustomer.cmdAdd.Value = True
    frmAllCustomer.Show 1
    FillGrid
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    optFullName.Value = True
    FillGrid
End Sub

Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCustomer
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstCustomer_DblClick()
    If Not lstCustomer.SelectedItem Is Nothing Then
        cmdSelect.Value = True
    End If
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtCode = lstCustomer.SelectedItem
    txtName = lstCustomer.SelectedItem.SubItems(3)
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
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
        If lstCustomer.ListItems.Count > 0 And lstCustomer.Enabled = True Then: lstCustomer.SetFocus

    ElseIf KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub txtCode_Change()
    If txtCode.Text <> "" Then
        cmdSelect.Enabled = True
    Else
        cmdSelect.Enabled = False
    End If
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmFind2 
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vendor"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3960
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4380
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4380
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4380
      Width           =   1695
   End
   Begin VB.CommandButton cmdCust_al 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2625
      MouseIcon       =   "frmFind2.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmFind2.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Customers"
      Top             =   5700
      Width           =   705
   End
   Begin VB.CommandButton cmdEmp_al 
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1920
      MouseIcon       =   "frmFind2.frx":07DD
      MousePointer    =   99  'Custom
      Picture         =   "frmFind2.frx":092F
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Employees"
      Top             =   5700
      Width           =   705
   End
   Begin VB.CommandButton cmdVend_al 
      Caption         =   "Vendor"
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
      Left            =   3330
      MouseIcon       =   "frmFind2.frx":0D53
      MousePointer    =   99  'Custom
      Picture         =   "frmFind2.frx":0EA5
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Vendors"
      Top             =   5700
      Width           =   705
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   2
      Top             =   90
      Width           =   3465
   End
   Begin MSComctlLib.ListView lvCustomer 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Acct Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Name"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Entity"
         Object.Width           =   671
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   915
      Left            =   1830
      TabIndex        =   7
      Top             =   5640
      Width           =   2265
      _Version        =   655364
      _ExtentX        =   4004
      _ExtentY        =   1614
      _StockProps     =   14
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   1245
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   11445
      _Version        =   655364
      _ExtentX        =   20188
      _ExtentY        =   2196
      _StockProps     =   14
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      ForeColor       =   4210752
   End
   Begin XtremeShortcutBar.ShortcutCaption sc 
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11445
      _Version        =   655364
      _ExtentX        =   20188
      _ExtentY        =   926
      _StockProps     =   14
      Caption         =   "Account Name"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "frmFind2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                                    As ADODB.Recordset
Dim rsVENDOR                                      As ADODB.Recordset
Dim rsEntity                                      As ADODB.Recordset
Dim xEntity                                       As ListItem
Dim Xcode                                         As String
Dim XaCCOUNTNAME                                  As String
Dim xEntityClass                                  As String
Dim xJOURNALTYPE                                  As String
Dim X_ent                                         As String
Event EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)

Sub initMemvars()
    If xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "CRJ" Then
        Set rsEntity = New ADODB.Recordset
        rsEntity.Open "Select Top 20 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsEntity.EOF And Not rsEntity.BOF Then
            Do While Not rsEntity.EOF
                Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
                xEntity.SubItems(1) = Null2String(rsEntity!accountname)
                xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
                rsEntity.MoveNext
            Loop
        End If
        Set rsEntity = Nothing
    ElseIf xJOURNALTYPE = "CDJ_HD" Or xJOURNALTYPE = "CDJ_DET" Or xJOURNALTYPE = "APJ_HD" Then
        If SelectEntity = "Customer" Then
        X_ent = "C"
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "Select Top 20 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND AccountName IS NOT NULL order by accountname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!Code)
                    xEntity.SubItems(1) = Null2String(rsCustomer!accountname)
                    xEntity.SubItems(2) = Null2String(rsCustomer!ENTITYCODE)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        ElseIf SelectEntity = "Vendor" Then
            X_ent = "V"
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "Select Top 20 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND AccountName IS NOT NULL order by accountname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!Code)
                    xEntity.SubItems(1) = Null2String(rsCustomer!accountname)
                    xEntity.SubItems(2) = Null2String(rsCustomer!ENTITYCODE)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsVENDOR = Nothing
        ElseIf SelectEntity = "Employee" Then
            X_ent = "E"
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "Select Top 20 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND AccountName IS NOT NULL order by accountname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!Code)
                    xEntity.SubItems(1) = Null2String(rsCustomer!accountname)
                    xEntity.SubItems(2) = Null2String(rsCustomer!ENTITYCODE)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        End If
    Else
        If SelectEntity = "Customer" Then
        X_ent = "C"
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "Select Top 20 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND AccountName IS NOT NULL order by accountname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!Code)
                    xEntity.SubItems(1) = Null2String(rsCustomer!accountname)
                    xEntity.SubItems(2) = Null2String(rsCustomer!ENTITYCODE)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        ElseIf SelectEntity = "Vendor" Then
            X_ent = "V"
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "Select Top 20 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND AccountName IS NOT NULL order by accountname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!Code)
                    xEntity.SubItems(1) = Null2String(rsCustomer!accountname)
                    xEntity.SubItems(2) = Null2String(rsCustomer!ENTITYCODE)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsVENDOR = Nothing
        ElseIf SelectEntity = "Employee" Then
            X_ent = "E"
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "Select Top 20 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND AccountName IS NOT NULL order by accountname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!Code)
                    xEntity.SubItems(1) = Null2String(rsCustomer!accountname)
                    xEntity.SubItems(2) = Null2String(rsCustomer!ENTITYCODE)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        End If
    End If
End Sub

Private Sub cmdCust_al_Click()
    lvCustomer.ListItems.Clear
    SelectEntity = "Customer"
    frmFind2.Caption = "SEARCH CUSTOMER"
    initMemvars
End Sub

Private Sub cmdEmp_al_Click()
    lvCustomer.ListItems.Clear
    SelectEntity = "Employee"
    frmFind2.Caption = "SEARCH EMPLOYEE"
    initMemvars
End Sub

Private Sub cmdVend_al_Click()
    lvCustomer.ListItems.Clear
    SelectEntity = "Vendor"
    frmFind2.Caption = "SEARCH VENDOR"
    initMemvars
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    SelectEntity = "Vendor"
    initMemvars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
    Unload frmEntity
End Sub

Private Sub lvCustomer_DblClick()
If xJOURNALTYPE = "CDJ1" Then
    
        If X_ent = "C" Then
            frmAMISJournalEntry_CDJ.labPV1.Caption = "Cust.Code"
            
        ElseIf X_ent = "E" Then
            frmAMISJournalEntry_CDJ.labPV1.Caption = "Emp.Code"
        Else
            frmAMISJournalEntry_CDJ.labPV1.Caption = "Ven.Code"
        End If
    
        frmAMISJournalEntry_CDJ.txtPO_No = X_ent & (Trim(Me.lvCustomer.SelectedItem.Text))
        'frmAMISJournalEntry_CDJ.txtMRR_No = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
        
ElseIf xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "CRJ" Then
        'frmAMISCustomerAROpening.txtCustCode.Text = (Trim(Me.lvCustomer.SelectedItem.Text))
        'frmAMISCustomerAROpening.txtCustName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
        Xcode = Trim(Me.lvCustomer.SelectedItem.Text)
        XaCCOUNTNAME = Trim(Me.lvCustomer.SelectedItem.SubItems(1))
        xEntityClass = Trim(Me.lvCustomer.SelectedItem.SubItems(2))
        RaiseEvent EntitySelected(Xcode, XaCCOUNTNAME, xEntityClass)
        Unload Me
ElseIf xJOURNALTYPE = "CDJ_HD" Then
        'frmAMISCustomerAROpening.txtCustCode.Text = (Trim(Me.lvCustomer.SelectedItem.Text))
        'frmAMISCustomerAROpening.txtCustName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
        frmAMISJournalEntry_CDJ.lblEntityD.Caption = Trim(Me.lvCustomer.SelectedItem.SubItems(2))
        'Xcode = Trim(Me.lvCustomer.SelectedItem.Text)
        frmAMISJournalEntry_CDJ.txtCode.Text = Trim(Me.lvCustomer.SelectedItem.Text)
        frmAMISJournalEntry_CDJ.cboNameofVendor.Text = Trim(Me.lvCustomer.SelectedItem.SubItems(1))
        frmAMISJournalEntry_CDJ.txtCode.Text = Trim(Me.lvCustomer.SelectedItem.Text)
       ' XaCCOUNTNAME = Trim(Me.lvCustomer.SelectedItem.SubItems(1))
        'xEntityClass = Trim(Me.lvCustomer.SelectedItem.SubItems(2))
        RaiseEvent EntitySelected(Xcode, XaCCOUNTNAME, xEntityClass)
        Unload Me
ElseIf xJOURNALTYPE = "CDJ_DET" Then
        'frmAMISCustomerAROpening.txtCustCode.Text = (Trim(Me.lvCustomer.SelectedItem.Text))
        'frmAMISCustomerAROpening.txtCustName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
        'frmAMISJournalEntry_CDJ.lblEntityD.Caption = Trim(Me.lvCustomer.SelectedItem.SubItems(2))
        'Xcode = Trim(Me.lvCustomer.SelectedItem.Text)
        frmAMISJournalEntry_CDJ.txtPO_No.Text = Trim(Me.lvCustomer.SelectedItem.SubItems(2)) + Trim(Me.lvCustomer.SelectedItem.Text)
        RaiseEvent EntitySelected(Xcode, XaCCOUNTNAME, xEntityClass)
        Unload Me

ElseIf xJOURNALTYPE = "APJ_HD" Then
        frmAMISJournalEntry_APJ.lblentity.Caption = Trim(Me.lvCustomer.SelectedItem.SubItems(2))
        frmAMISJournalEntry_APJ.txtCode.Text = Trim(Me.lvCustomer.SelectedItem.Text)
        frmAMISJournalEntry_APJ.cboNameofVendor.Text = Trim(Me.lvCustomer.SelectedItem.SubItems(1))
        frmAMISJournalEntry_APJ.txtCode.Text = Trim(Me.lvCustomer.SelectedItem.Text)
        RaiseEvent EntitySelected(Xcode, XaCCOUNTNAME, xEntityClass)
        Unload Me
ElseIf xJOURNALTYPE = "APJ_DET" Then
        frmAMISJournalEntry_APJ.lblentitydet.Caption = Trim(Me.lvCustomer.SelectedItem.SubItems(2))
        frmAMISJournalEntry_APJ.txtPO_No.Text = Trim(Me.lvCustomer.SelectedItem.Text)
        RaiseEvent EntitySelected(Xcode, XaCCOUNTNAME, xEntityClass)
        Unload Me
End If
    Unload Me
End Sub

Private Sub lvCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lvCustomer_DblClick
    End If
End Sub

Private Sub Option1_Click()
    lvCustomer.ListItems.Clear
    SelectEntity = "Employee"
    frmFind2.Caption = "SEARCH EMPLOYEE"
    initMemvars
End Sub

Private Sub Option2_Click()
    lvCustomer.ListItems.Clear
    SelectEntity = "Customer"
    frmFind2.Caption = "SEARCH CUSTOMER"
    initMemvars
End Sub

Private Sub Option3_Click()
    lvCustomer.ListItems.Clear
    SelectEntity = "Vendor"
    frmFind2.Caption = "SEARCH VENDOR"
    initMemvars
End Sub

Private Sub txtSearch_Change()
    If xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "CRJ" Then
        Set rsEntity = New ADODB.Recordset
        lvCustomer.ListItems.Clear
        If txtSearch.Text = "" Then
            rsEntity.Open "Select Top 20 Code,AccountName,EntityCode from ALL_ENTITY where AccountName like '%" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsEntity.Open "Select Code,AccountName,EntityCode from ALL_ENTITY where AccountName like '%" & txtSearch.Text & "%'  ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If Not rsEntity.EOF And Not rsEntity.BOF Then
            Do While Not rsEntity.EOF
                Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
                xEntity.SubItems(1) = Null2String(rsEntity!accountname)
                xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
                rsEntity.MoveNext
            Loop
        End If
        Set rsEntity = Nothing
    ElseIf xJOURNALTYPE = "CDJ1" Then
        Set rsEntity = New ADODB.Recordset
        lvCustomer.ListItems.Clear
        If txtSearch.Text = "" Then
            rsEntity.Open "Select Top 20 Code,AccountName,EntityCode from ALL_ENTITY where AccountName like '%" & txtSearch.Text & "%'  ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsEntity.Open "Select Code,AccountName,EntityCode from ALL_ENTITY where AccountName like '%" & txtSearch.Text & "%' and ENTITYCODE='" & Left(SelectEntity, 1) & "' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If Not rsEntity.EOF And Not rsEntity.BOF Then
            Do While Not rsEntity.EOF
                Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
                xEntity.SubItems(1) = Null2String(rsEntity!accountname)
                xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
                rsEntity.MoveNext
            Loop
        End If
        Set rsEntity = Nothing
    Else
        If SelectEntity = "Customer" Then
            Set rsCustomer = New ADODB.Recordset
            lvCustomer.ListItems.Clear
            If txtSearch.Text = "" Then
                rsCustomer.Open "Select Top 20 code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND ACCOUNTNAME like '%" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsCustomer.Open "Select code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND ACCOUNTNAME like '%" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!Code)
                    xEntity.SubItems(1) = rsCustomer!accountname
                    xEntity.SubItems(2) = rsCustomer!ENTITYCODE
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        ElseIf SelectEntity = "Vendor" Then
            Set rsVENDOR = New ADODB.Recordset
            lvCustomer.ListItems.Clear
            If txtSearch.Text = "" Then
                rsVENDOR.Open "Select Top 20 Code,NameofVendor from ALL_Vendor where NameofVendor like '%" & txtSearch.Text & "%' ORDER BY Code,NameofVendor", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsVENDOR.Open "Select Code,NameofVendor from ALL_Vendor where NameofVendor like '%" & txtSearch.Text & "%' ORDER BY Code,NameofVendor", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
                Do While Not rsVENDOR.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsVENDOR!Code)
                    xEntity.SubItems(1) = rsVENDOR!nameofvendor
                    rsVENDOR.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        End If
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtSearch.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lvCustomer.ListItems.Count > 0 And lvCustomer.Enabled = True Then: lvCustomer.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Sub LoadJournal2(XXX As String)
    xJOURNALTYPE = XXX
End Sub


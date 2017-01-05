VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmEntity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmEntity.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6210
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
      Left            =   1650
      TabIndex        =   0
      Top             =   90
      Width           =   3465
   End
   Begin MSComctlLib.ListView lvCustomer 
      Height          =   3915
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   6906
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
   Begin XtremeShortcutBar.ShortcutCaption sc1 
      Height          =   525
      Left            =   0
      TabIndex        =   2
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
Attribute VB_Name = "frmEntity"
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
Event EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)

Sub initMemvars()
    If xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "CRJ" Then
        Set rsEntity = New ADODB.Recordset
        rsEntity.Open "Select Top 20 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsEntity.EOF And Not rsEntity.BOF Then
            Do While Not rsEntity.EOF
                Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!code)
                xEntity.SubItems(1) = Null2String(rsEntity!AccountName)
                xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
                rsEntity.MoveNext
            Loop
        End If
        Set rsEntity = Nothing
    Else
        If SelectEntity = "Customer" Then
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "Select Top 20 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND AccountName IS NOT NULL", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!code)
                    xEntity.SubItems(1) = Null2String(rsCustomer!AccountName)
                    xEntity.SubItems(2) = Null2String(rsCustomer!ENTITYCODE)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        ElseIf SelectEntity = "Vendor" Then
            Set rsVENDOR = New ADODB.Recordset
            rsVENDOR.Open "Select Top 20 Code,NameofVendor from ALL_Vendor", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
                Do While Not rsVENDOR.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsVENDOR!code)
                    xEntity.SubItems(1) = rsVENDOR!nameofvendor
                    rsVENDOR.MoveNext
                Loop
            End If
            Set rsVENDOR = Nothing
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    initMemvars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
    Unload frmEntity
End Sub

Private Sub lvCustomer_DblClick()
'    If xJOURNALTYPE = "APJ" Then
'        frmAMIS_APJ_JOURNAL_ENTRY.txtAccountName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
'    ElseIf xJOURNALTYPE = "CDJ" Then
'        frmAMIS_CDJ_JOURNAL_ENTRY.txtAccountName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
'    ElseIf xJOURNALTYPE = "CRJ" Then
'        frmAMIS_CRJ_JOURNAL_ENTRY.txtAccountName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
'    ElseIf xJOURNALTYPE = "DRJ" Then
'        frmAMIS_DRJ_JOURNAL_ENTRY.txtAccountName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
'    ElseIf xJOURNALTYPE = "SJ" Then
'        frmAMIS_SJ_JOURNAL_ENTRY.txtAccountName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
'    Else

    If xJOURNALTYPE = "GJ" Then
        frmAMISJournalEntry_GJDetails.txtCode = (Trim(Me.lvCustomer.SelectedItem.Text))
        frmAMISJournalEntry_GJDetails.txtName = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))

        If SelectEntity = "Customer" Then
            frmAMISJournalEntry_GJDetails.lblCode.Caption = "Cust. Code"
            frmAMISJournalEntry_GJDetails.lblName.Caption = "Cust. Name"
            frmAMISJournalEntry_GJDetails.lblInvoiceNo.Caption = "Inv. No"
            frmAMISJournalEntry_GJDetails.lblInvoiceType.Caption = "Inv. Type"

            If frmAMISJournalEntry_GJDetails.chkOther.Value = 1 Then
                'do nothing
            Else
                'frmAMIS_GJ_ENTRY.txtInvoiceNo.SetFocus
            End If

        Else
            frmAMISJournalEntry_GJDetails.lblCode.Caption = "Supp. Code"
            frmAMISJournalEntry_GJDetails.lblName.Caption = "Supp. Name"
            frmAMISJournalEntry_GJDetails.lblInvoiceNo.Caption = "MRR No"
            frmAMISJournalEntry_GJDetails.lblInvoiceType.Caption = "MRR Type"

            If frmAMISJournalEntry_GJDetails.chkOther.Value = 1 Then
                'do nothing
            Else
                'frmAMIS_GJ_ENTRY.txtInvoiceNo.SetFocus
            End If
        End If
    ElseIf xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "CRJ" Then
        'frmAMISCustomerAROpening.txtCustCode.Text = (Trim(Me.lvCustomer.SelectedItem.Text))
        'frmAMISCustomerAROpening.txtCustName.Text = (Trim(Me.lvCustomer.SelectedItem.SubItems(1)))
        Xcode = Trim(Me.lvCustomer.SelectedItem.Text)
        XaCCOUNTNAME = Trim(Me.lvCustomer.SelectedItem.SubItems(1))
        xEntityClass = Trim(Me.lvCustomer.SelectedItem.SubItems(2))
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

Private Sub txtSEARCH_Change()
    If xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "CRJ" Then
        Set rsEntity = New ADODB.Recordset
        lvCustomer.ListItems.Clear
        rsEntity.Open "Select Top 20 Code,AccountName,EntityCode from ALL_ENTITY where AccountName like '" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsEntity.EOF And Not rsEntity.BOF Then
            Do While Not rsEntity.EOF
                Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!code)
                xEntity.SubItems(1) = Null2String(rsEntity!AccountName)
                xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
                rsEntity.MoveNext
            Loop
        End If
        Set rsEntity = Nothing
    Else
        If SelectEntity = "Customer" Then
            Set rsCustomer = New ADODB.Recordset
            lvCustomer.ListItems.Clear
            rsCustomer.Open "Select Top 20 code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND ACCOUNTNAME like '" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!code)
                    xEntity.SubItems(1) = rsCustomer!AccountName
                    xEntity.SubItems(2) = rsCustomer!ENTITYCODE
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        ElseIf SelectEntity = "Vendor" Then
            Set rsVENDOR = New ADODB.Recordset
            lvCustomer.ListItems.Clear
            rsVENDOR.Open "Select Top 20 Code,NameofVendor from ALL_Vendor where NameofVendor like '" & txtSearch.Text & "%' ORDER BY Code,NameofVendor", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
                Do While Not rsVENDOR.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsVENDOR!code)
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

Sub LoadJournal(XXX As String)
    xJOURNALTYPE = XXX
End Sub

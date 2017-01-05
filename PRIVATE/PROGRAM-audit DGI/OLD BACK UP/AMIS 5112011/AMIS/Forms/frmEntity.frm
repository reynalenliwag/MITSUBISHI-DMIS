VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmEntity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmEntity.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6195
   Begin VB.PictureBox picEntity 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   60
      ScaleHeight     =   465
      ScaleWidth      =   6075
      TabIndex        =   4
      Top             =   570
      Visible         =   0   'False
      Width           =   6075
      Begin VB.OptionButton optCustomer 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -30
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Value           =   -1  'True
         Width           =   2040
      End
      Begin VB.OptionButton optVendor 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vendor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2010
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   2040
      End
      Begin VB.OptionButton optEmployee 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1500
      TabIndex        =   0
      Top             =   90
      Width           =   4635
   End
   Begin MSComctlLib.ListView lvCustomer 
      Height          =   5955
      Left            =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10504
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
      Appearance      =   0
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
         Text            =   "Code"
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
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   3
      Top             =   210
      Width           =   1380
   End
   Begin XtremeShortcutBar.ShortcutCaption sc1 
      Height          =   7125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11445
      _Version        =   655364
      _ExtentX        =   20188
      _ExtentY        =   12568
      _StockProps     =   14
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
Dim rsCustomer                                         As ADODB.Recordset
Dim rsVENDOR                                           As ADODB.Recordset
Dim rsEntity                                           As ADODB.Recordset
Dim xEntity                                            As ListItem
Dim Xcode                                              As String
Dim XaCCOUNTNAME                                       As String
Dim xEntityClass                                       As String
Dim xJOURNALTYPE                                       As String
Event EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)
Dim xEntityCode As String

Sub initMemvars()
    lvCustomer.ListItems.Clear
    If xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CRJ" Then
        picEntity.Visible = True
        If xJOURNALTYPE = "COB" Then
            optCustomer.Value = True
        ElseIf xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Then
            optVendor.Value = True
        End If
        If optCustomer.Value = True Then
            Set rsEntity = New ADODB.Recordset
            rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='C' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsEntity.EOF And Not rsEntity.BOF Then
                Do While Not rsEntity.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
                    xEntity.SubItems(1) = Null2String(rsEntity!ACCOUNTNAME)
                    xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
                    rsEntity.MoveNext
                Loop
            End If
            xEntityCode = "C"
        ElseIf optVendor.Value = True Then
            Set rsEntity = New ADODB.Recordset
            rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='V' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsEntity.EOF And Not rsEntity.BOF Then
                Do While Not rsEntity.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
                    xEntity.SubItems(1) = Null2String(rsEntity!ACCOUNTNAME)
                    xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
                    rsEntity.MoveNext
                Loop
            End If
            xEntityCode = "V"
        ElseIf optEmployee.Value = True Then
            Set rsEntity = New ADODB.Recordset
            rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='E' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsEntity.EOF And Not rsEntity.BOF Then
                Do While Not rsEntity.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
                    xEntity.SubItems(1) = Null2String(rsEntity!ACCOUNTNAME)
                    xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
                    rsEntity.MoveNext
                Loop
            End If
            xEntityCode = "E"
        End If
        Set rsEntity = Nothing
        lvCustomer.Top = 1080
        lvCustomer.Height = 5955
    Else
        picEntity.Visible = False
        If SelectEntity = "Customer" Then
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "Select Top 100 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='" & Left(SelectEntity, 1) & "' AND AccountName IS NOT NULL ORDER BY ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                Do While Not rsCustomer.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsCustomer!Code)
                    xEntity.SubItems(1) = Null2String(rsCustomer!ACCOUNTNAME)
                    xEntity.SubItems(2) = Null2String(rsCustomer!ENTITYCODE)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        ElseIf SelectEntity = "Vendor" Then
            Set rsVENDOR = New ADODB.Recordset
            rsVENDOR.Open "Select Top 100 Code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='V' AND AccountName IS NOT NULL ORDER BY ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
                Do While Not rsVENDOR.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsVENDOR!Code)
                    xEntity.SubItems(1) = Null2String(rsVENDOR!ACCOUNTNAME)
                    xEntity.SubItems(2) = Null2String(rsVENDOR!ENTITYCODE)
                    rsVENDOR.MoveNext
                Loop
            End If
            Set rsVENDOR = Nothing
        End If
        lvCustomer.Top = 570
        lvCustomer.Height = 6465
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
    ElseIf xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CRJ" Then
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

Private Sub optCustomer_Click()
    xEntityCode = "C"
    lvCustomer.ListItems.Clear
    Set rsEntity = New ADODB.Recordset
    If txtSearch.Text = "" Then
        rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='C' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='C' AND ACCOUNTNAME LIKE '%" & txtSearch.Text & "%' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsEntity.EOF And Not rsEntity.BOF Then
        Do While Not rsEntity.EOF
            Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
            xEntity.SubItems(1) = Null2String(rsEntity!ACCOUNTNAME)
            xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
            rsEntity.MoveNext
        Loop
    End If
    txtSearch.SetFocus
End Sub

Private Sub optEmployee_Click()
    xEntityCode = "E"
    lvCustomer.ListItems.Clear
    Set rsEntity = New ADODB.Recordset
    If txtSearch.Text = "" Then
        rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='E' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='E' AND ACCOUNTNAME LIKE '%" & txtSearch.Text & "%' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsEntity.EOF And Not rsEntity.BOF Then
        Do While Not rsEntity.EOF
            Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
            xEntity.SubItems(1) = Null2String(rsEntity!ACCOUNTNAME)
            xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
            rsEntity.MoveNext
        Loop
    End If
    txtSearch.SetFocus
End Sub

Private Sub optVendor_Click()
On Error Resume Next
    xEntityCode = "V"
    lvCustomer.ListItems.Clear
    Set rsEntity = New ADODB.Recordset
    If txtSearch.Text = "" Then
        rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='V' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsEntity.Open "Select Top 100 EntityCode,Code,AccountName from ALL_ENTITY where AccountName IS NOT NULL AND ENTITYCODE='V' AND ACCOUNTNAME LIKE '%" & txtSearch.Text & "%' ORDER BY EntityCode,ACCOUNTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsEntity.EOF And Not rsEntity.BOF Then
        Do While Not rsEntity.EOF
            Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
            xEntity.SubItems(1) = Null2String(rsEntity!ACCOUNTNAME)
            xEntity.SubItems(2) = Null2String(rsEntity!ENTITYCODE)
            rsEntity.MoveNext
        Loop
    End If
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_Change()
    If xJOURNALTYPE = "COB" Or xJOURNALTYPE = "VPJ" Or xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "CRJ" Then
        Set rsEntity = New ADODB.Recordset
        lvCustomer.ListItems.Clear
        If txtSearch.Text = "" Then
            rsEntity.Open "Select Top 100 Code,AccountName,EntityCode from ALL_ENTITY where EntityCode = '" & xEntityCode & "' AND AccountName like '%" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsEntity.Open "Select Code,AccountName,EntityCode from ALL_ENTITY where EntityCode = '" & xEntityCode & "' AND AccountName like '%" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If Not rsEntity.EOF And Not rsEntity.BOF Then
            Do While Not rsEntity.EOF
                Set xEntity = lvCustomer.ListItems.Add(, , rsEntity!Code)
                xEntity.SubItems(1) = Null2String(rsEntity!ACCOUNTNAME)
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
                    xEntity.SubItems(1) = rsCustomer!ACCOUNTNAME
                    xEntity.SubItems(2) = rsCustomer!ENTITYCODE
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        ElseIf SelectEntity = "Vendor" Then
            Set rsVENDOR = New ADODB.Recordset
            lvCustomer.ListItems.Clear
            If txtSearch.Text = "" Then
                rsVENDOR.Open "Select Top 20 code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='V' AND ACCOUNTNAME like '%" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsVENDOR.Open "Select code,AccountName,ENTITYCODE from ALL_ENTITY where ENTITYCODE='V' AND ACCOUNTNAME like '%" & txtSearch.Text & "%' ORDER BY Code,AccountName", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
                Do While Not rsVENDOR.EOF
                    Set xEntity = lvCustomer.ListItems.Add(, , rsVENDOR!Code)
                    xEntity.SubItems(1) = rsVENDOR!ACCOUNTNAME
                    xEntity.SubItems(2) = rsVENDOR!ENTITYCODE
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

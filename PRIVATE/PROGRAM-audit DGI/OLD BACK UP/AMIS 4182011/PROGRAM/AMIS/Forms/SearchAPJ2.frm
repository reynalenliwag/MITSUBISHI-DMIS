VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmAMISSearchAPJ2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Accounts Payable Journals"
   ClientHeight    =   6330
   ClientLeft      =   2970
   ClientTop       =   3495
   ClientWidth     =   8715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "SearchAPJ2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8715
   Begin XtremeSuiteControls.TabControl SearchTab 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   8700
      _Version        =   655364
      _ExtentX        =   15346
      _ExtentY        =   10504
      _StockProps     =   64
      Appearance      =   3
      Color           =   4
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "By &Voucher No"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By Vendor/&Payee Name"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   5325
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll2 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7560
            TabIndex        =   11
            Top             =   120
            Width           =   1035
         End
         Begin VB.TextBox txtVendorPayeeName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1215
            TabIndex        =   3
            Top             =   45
            Width           =   6195
         End
         Begin MSComctlLib.ListView ListVendorPayeeName 
            Height          =   4815
            Left            =   45
            TabIndex        =   4
            Top             =   450
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   8493
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "SearchAPJ2.frx":000C
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VENDOR NAME"
               Object.Width           =   7055
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "VOUCHER NO."
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "BALANCE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "DUE DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "JTYPE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "PARTICULARS"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "ACCT CODE"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   45
            TabIndex        =   2
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5325
         Left            =   -69970
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7560
            TabIndex        =   10
            Top             =   120
            Width           =   1035
         End
         Begin VB.TextBox txtVoucherNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1215
            TabIndex        =   7
            Top             =   45
            Width           =   6240
         End
         Begin MSComctlLib.ListView ListVoucherNo 
            Height          =   4845
            Left            =   45
            TabIndex        =   8
            Top             =   450
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   8546
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "SearchAPJ2.frx":016E
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VOUCHER NO."
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ENTITY CODE"
               Object.Width           =   5291
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "BALANCE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "DUE DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "JTYPE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "PARTICULAR"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "ACCT CODE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "INVOICE NO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "INVOICE TYPE"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   45
            TabIndex        =   6
            Top             =   90
            Width           =   1125
         End
      End
   End
   Begin MSForms.CheckBox chkCurrentVend 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   30
      Width           =   8685
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   4
      Size            =   "15319;556"
      Value           =   "1"
      Caption         =   "Show Journals for Current Vendor Only"
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmAMISSearchAPJ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                       As New ADODB.Recordset
Dim Y, k                                               As Long
Attribute k.VB_VarUserMemId = 1073938433
Dim xJOURNALTYPE                                       As String
Sub LoadJournal(XXX As String)
    xJOURNALTYPE = XXX
End Sub

Sub clearListView()
    For Y = 1 To Me.ListVoucherNo.ListItems.Count
        If Me.ListVoucherNo.ListItems.Count <= 0 Then Exit For
        Me.ListVoucherNo.Sorted = False
        Me.ListVoucherNo.ListItems.Remove Me.ListVoucherNo.SelectedItem.Index
    Next Y
    For Y = 1 To Me.ListVendorPayeeName.ListItems.Count
        If Me.ListVendorPayeeName.ListItems.Count <= 0 Then Exit For
        Me.ListVendorPayeeName.Sorted = False
        Me.ListVendorPayeeName.ListItems.Remove Me.ListVendorPayeeName.SelectedItem.Index
    Next Y
End Sub

Private Sub chkCurrentVend_Click()
    SearchTab.SelectedItem = SEARCH_TAB
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
End Sub

Private Sub chkShowAll_Click()
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
End Sub

Private Sub chkShowAll2_Click()
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Select Case SEARCH_TAB
        Case 0:
            On Error Resume Next
            txtVoucherNo.SetFocus
        Case 1:
            On Error Resume Next
            txtVendorPayeeName.SetFocus
        End Select
    End If
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SearchTab.SelectedItem = SEARCH_TAB
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
End Sub

Private Sub ListVendorPayeeName_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListVendorPayeeName
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub ListVendorPayeeName_DblClick()
    On Error Resume Next
    If xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "APJ" Then
        frmAMISJournalEntry_CDJ.txtPO_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(6)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
        frmAMISJournalEntry_CDJ.txtPVAmount.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(4)))
        frmAMISJournalEntry_CDJ.lblPVAmount.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(4)))
        frmAMISJournalEntry_CDJ.lblJ_CLASS.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(8)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtProd_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtINV_No.Enabled = False
    Else
        frmAMISJournalEntry.txtPO_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(5)))
        frmAMISJournalEntry.txtMRR_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
        frmAMISJournalEntry.txtPVAmount.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(4)))
        frmAMISJournalEntry.txtMRR_No.Enabled = False
        frmAMISJournalEntry.txtProd_No.Enabled = False
        frmAMISJournalEntry.txtINV_No.Enabled = False
    End If
    'frmAMISJournalEntry.cmdPVSave_Click
    Unload Me
End Sub

Private Sub ListVoucherNo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListVoucherNo
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub ListVoucherNo_DblClick()
    On Error Resume Next
    ' UPDATE BTT 1/27/2009
    If xJOURNALTYPE = "CDJ" Then
        frmAMISJournalEntry_CDJ.txtPO_No.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(6)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Text = (Trim(Me.ListVoucherNo.SelectedItem))
        frmAMISJournalEntry_CDJ.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(4)))
        frmAMISJournalEntry_CDJ.lblPVAmount.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(4)))
        frmAMISJournalEntry_CDJ.lblJ_CLASS.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(8)))
        frmAMISJournalEntry_CDJ.lblInvoiceNo.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(9)))
        frmAMISJournalEntry_CDJ.lblInvoiceType.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(10)))
        'frmAMISJournalEntry.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(3)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtProd_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtINV_No.Enabled = False
        
    ElseIf xJOURNALTYPE = "APJ" Then
        frmAMISJournalEntry_CDJ.txtPO_No.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(5)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Text = (Trim(Me.ListVoucherNo.SelectedItem))
        frmAMISJournalEntry_CDJ.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(4)))
        'frmAMISJournalEntry.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(3)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtProd_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtINV_No.Enabled = False
    Else
        frmAMISJournalEntry.txtPO_No.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(5)))
        frmAMISJournalEntry.txtMRR_No.Text = (Trim(Me.ListVoucherNo.SelectedItem))
        frmAMISJournalEntry.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(4)))
        'frmAMISJournalEntry.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(3)))
        frmAMISJournalEntry.txtMRR_No.Enabled = False
        frmAMISJournalEntry.txtProd_No.Enabled = False
        frmAMISJournalEntry.txtINV_No.Enabled = False
        'frmAMISJournalEntry.cmdPVSave_Click
    End If
    Unload Me
End Sub

Private Sub ListVoucherNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtVoucherNo.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListVoucherNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        If xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "APJ" Then
            frmAMISJournalEntry_CDJ.txtPO_No.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(6)))
            frmAMISJournalEntry_CDJ.txtMRR_No.Text = (Trim(Me.ListVoucherNo.SelectedItem))
            frmAMISJournalEntry_CDJ.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(4)))
            frmAMISJournalEntry_CDJ.lblJ_CLASS.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(8)))
            frmAMISJournalEntry_CDJ.lblPVAmount.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(4)))
            frmAMISJournalEntry_CDJ.txtMRR_No.Enabled = False
            frmAMISJournalEntry_CDJ.txtProd_No.Enabled = False
            frmAMISJournalEntry_CDJ.txtINV_No.Enabled = False
        Else
            frmAMISJournalEntry.txtPO_No.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(5)))
            frmAMISJournalEntry.txtMRR_No.Text = (Trim(Me.ListVoucherNo.SelectedItem))
            frmAMISJournalEntry.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(4)))
            frmAMISJournalEntry.txtMRR_No.Enabled = False
            frmAMISJournalEntry.txtProd_No.Enabled = False
            frmAMISJournalEntry.txtINV_No.Enabled = False
            'frmAMISJournalEntry.cmdPVSave_Click
        End If
        Unload Me
    End If
End Sub

Private Sub ListVendorPayeeName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtVendorPayeeName.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListVendorPayeeName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "APJ" Then
            frmAMISJournalEntry_CDJ.txtPO_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(6)))
            frmAMISJournalEntry_CDJ.txtMRR_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
            frmAMISJournalEntry_CDJ.txtPVAmount.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(4)))
            frmAMISJournalEntry_CDJ.lblPVAmount.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(4)))
            frmAMISJournalEntry_CDJ.lblJ_CLASS.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(8)))
            frmAMISJournalEntry_CDJ.txtMRR_No.Enabled = False
            frmAMISJournalEntry_CDJ.txtProd_No.Enabled = False
            frmAMISJournalEntry_CDJ.txtINV_No.Enabled = False
        Else
            frmAMISJournalEntry.txtPO_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(5)))
            frmAMISJournalEntry.txtMRR_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
            frmAMISJournalEntry.txtPVAmount.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(4)))
            frmAMISJournalEntry.txtMRR_No.Enabled = False
            frmAMISJournalEntry.txtProd_No.Enabled = False
            frmAMISJournalEntry.txtINV_No.Enabled = False
            'frmAMISJournalEntry.cmdPVSave_Click
        End If
        Unload Me
    End If
End Sub

Private Sub SearchTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    SEARCH_TAB = SearchTab.SelectedItem
    DoEvents
    txtVoucherNo.Enabled = False
    txtVendorPayeeName.Enabled = False
    ListVoucherNo.Enabled = False
    ListVendorPayeeName.Enabled = False
    Select Case SEARCH_TAB
    Case 0
        txtVoucherNo.Enabled = True: ListVoucherNo.Enabled = True
        Me.Caption = "Search Item by Voucher Number"
        On Error Resume Next
        txtVoucherNo.SetFocus
        txtVoucherNo_Change
    Case 1
        txtVendorPayeeName.Enabled = True: ListVendorPayeeName.Enabled = True
        Me.Caption = "Search Item by All_VENDOR/Payee Name"
        On Error Resume Next
        txtVendorPayeeName.SetFocus
        txtVendorPayeeName_Change
    End Select
End Sub

Private Sub txtVoucherNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtVoucherNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListVoucherNo.Enabled = True Then ListVoucherNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVoucherNo_Change()
' UPDATE BY BTT
'UPDATED BY: ACL 9202010
'    Dim xOPTION1 As String
'    Dim xOPTION2 As String
'    Dim CMD As New ADODB.Command
'    xOPTION1 = "VOUCHER"
    If txtVoucherNo = "" Then
        Me.ListVoucherNo.Sorted = False: Me.ListVoucherNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentVend.Value = True Then
            'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_VENDOR.nameofvendor,AMIS_Journal_Hd.Balance,AMIS_Journal_Hd.JType,AMIS_Journal_Hd.remarks from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where AMIS_Journal_Hd.vendorcode = '" & CURRENT_VENDORCODE & "' AND (Jtype = 'VPJ' OR Jtype = 'APJ' or Jtype = 'VDJ') and status = 'P' and (PaidStatus = 'N' OR AMIS_Journal_Hd.Balance > 0) order by VoucherNo asc")
            'Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (SELECT AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.Jdate,vendorcode,AMIS_JOURNAL_HD.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_JOURNAL_HD.jtype,AMIS_JOURNAL_HD.remarks From AMIS_JOURNAL_HD WHERE AMIS_JOURNAL_HD.JTYPE IN('APJ','VPJ','VDJ')AND AMIS_JOURNAL_HD.STATUS ='P') AS T WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' and T.XBALANCE > 0")
            'Set rsJournal_HD = gconDMIS.Execute("SELECT TOP 22 * FROM (SELECT HD.VOUCHERNO,HD.JDATE,VENDORCODE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END),(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END) - " & _
             "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS XBALANCE, " & _
             "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE INNER JOIN AMIS_CHARTACCOUNT AC ON DET.ACCT_CODE=AC.ACCTCODE WHERE HD.JTYPE IN('APJ','VPJ','VDJ') AND HD.STATUS ='P' AND AC.IS_SCHEDULE_ACCNT=1) AS T " & _
             "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' and T.XBALANCE <> 0 ORDER BY VOUCHERNO,JTYPE")

            'Set rsJournal_HD = gconDMIS.Execute("SELECT TOP 22 * FROM (SELECT HD.VOUCHERNO,HD.JDATE,VENDORCODE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END),(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END) - " & _
             "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS XBALANCE, " & _
             "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ','VDJ') AND HD.STATUS ='P') AS T " & _
             "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' and T.XBALANCE <> 0 ORDER BY VOUCHERNO,JTYPE")

            SQL_STATEMENT = "SELECT VOUCHERNO,JDATE,VENDORCODE,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),DUEDATE,JTYPE,REMARKS,ACCT_CODE,NULL AS INVOICENO,NULL AS INVOICETYPE FROM (SELECT HD.VOUCHERNO,HD.JDATE,HD.DUEDATE,VENDORCODE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END), " & _
                                                "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AMOUNTPAID, " & _
                                                "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ') AND HD.STATUS ='P') AS T " & _
                                                "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND (AMOUNTTOPAY-AMOUNTPAID) <> 0 " '& vbCrLf
            'SQL_STATEMENT = SQL_STATEMENT & "UNION" & vbCrLf
            'SQL_STATEMENT = SQL_STATEMENT & "SELECT VOUCHERNO,JDATE,CUSTOMERCODE,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),JTYPE,REMARKS,ACCT_CODE,INVOICENO,INVOICETYPE FROM " & _
                            "(SELECT HD.VOUCHERNO,HD.JDATE,HD.CUSTOMERCODE, ISNULL((SELECT SUM(ISNULL(INVOICEAMOUNT,0)) FROM AMIS_CRJ_DETAIL " & _
                            "WHERE (HD.JTYPE=AMIS_CRJ_DETAIL.CR_TYPE  AND VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AMOUNTTOPAY, " & _
                            "ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AMOUNTPAID, " & _
                            "HD.jtype , HD.REMARKS, Det.ACCT_CODE,ACD.INVOICENO,ACD.INVOICETYPE " & _
                            "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE " & _
                            "INNER JOIN AMIS_CRJ_DETAIL ACD ON HD.VOUCHERNO=ACD.VOUCHERNO AND HD.JTYPE=ACD.CR_TYPE " & _
                            "WHERE HD.JTYPE IN('CRJ') AND HD.STATUS ='P' AND DET.CREDIT > 0) AS T WHERE CUSTOMERCODE = '" & CURRENT_VENDORCODE & "' AND (ISNULL(AMOUNTTOPAY,0)-ISNULL(AMOUNTPAID,0)) <> 0 " & vbCrLf
            'SQL_STATEMENT = SQL_STATEMENT & "UNION" & vbCrLf
            'SQL_STATEMENT = SQL_STATEMENT & "SELECT RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE,AMOUNT2PAY,BALANCE=(AMOUNT2PAY-AMOUNTPAID),LEFT(VOUCHERNO,3) AS JTYPE,NULL AS REMARKS,ACCT_CODE,INVOICENO,INVOICETYPE FROM " & _
                            "AMIS_AP WHERE LEFT(VOUCHERNO,3)='COB' AND VENDOR_CODE = '" & CURRENT_VENDORCODE & "' AND (ISNULL(AMOUNT2PAY,0)-ISNULL(AMOUNTPAID,0)) <> 0"

            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
            ',ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=AMIS_CRJ_DETAIL.CR_TYPE) GROUP BY J_CLASS) ,0) AMOUNTPAID
            '            xOPTION2 = "CURRENT"
            '
            '            CMD.ActiveConnection = gconDMIS
            '            CMD.CommandType = adCmdStoredProc
            '            CMD.CommandText = "USP_SEARCH_AP"
            '            CMD.CommandTimeout = 500
            '
            '            With CMD.Parameters
            '                .Append CMD.CreateParameter("@VENDORCODE", adVarChar, adParamInput, 6, CURRENT_VENDORCODE)
            '                .Append CMD.CreateParameter("@OPTION1", adVarChar, adParamInput, 7, xOPTION1)
            '                .Append CMD.CreateParameter("@OPTION2", adVarChar, adParamInput, 7, xOPTION2)
            '            End With
            '
            '            Set rsJournal_HD = CMD.Execute
        Else
            '            xOPTION2 = "ALL"
            'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_VENDOR.nameofvendor,AMIS_Journal_Hd.Balance,AMIS_Journal_Hd.JType,AMIS_Journal_Hd.remarks from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where (Jtype = 'VPJ' OR Jtype = 'APJ' or Jtype = 'VDJ') and status = 'P' and (PaidStatus = 'N' OR AMIS_Journal_Hd.Balance > 0) order by VoucherNo asc")
            'Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (SELECT AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.Jdate,AMIS_JOURNAL_HD.vendorcode,AMIS_JOURNAL_HD.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_JOURNAL_HD.jtype,AMIS_JOURNAL_HD.remarks From AMIS_JOURNAL_HD WHERE AMIS_JOURNAL_HD.JTYPE IN('APJ','VPJ','VDJ')AND AMIS_JOURNAL_HD.STATUS ='P') AS T WHERE T.XBALANCE > 0")
            'APJ & CDJ REFERENCE===============================================
            'Set rsJournal_HD = gconDMIS.Execute("SELECT TOP 22 * FROM (SELECT HD.VOUCHERNO,HD.JDATE,VENDORCODE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END),(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END) - " & _
             "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS XBALANCE, " & _
             "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ','VDJ') AND HD.STATUS ='P') AS T " & _
             "WHERE T.XBALANCE <> 0 ORDER BY VOUCHERNO,JTYPE")
            Set rsJournal_HD = gconDMIS.Execute("SELECT VOUCHERNO,JDATE,VENDORCODE,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),DUEDATE,JTYPE,REMARKS,ACCT_CODE FROM (SELECT HD.VOUCHERNO,HD.JDATE,HD.DUEDATE,VENDORCODE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END), " & _
                                                "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AMOUNTPAID, " & _
                                                "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ') AND HD.STATUS ='P') AS T " & _
                                                "WHERE (AMOUNTTOPAY-AMOUNTPAID) <> 0 ORDER BY VOUCHERNO,JTYPE")
            '            CMD.ActiveConnection = gconDMIS
            '            CMD.CommandType = adCmdStoredProc
            '            CMD.CommandText = "USP_SEARCH_AP"
            '            CMD.CommandTimeout = 500
            '
            '            With CMD.Parameters
            '                .Append CMD.CreateParameter("@VENDORCODE", adVarChar, adParamInput, 6, CURRENT_VENDORCODE)
            '                .Append CMD.CreateParameter("@OPTION1", adVarChar, adParamInput, 7, xOPTION1)
            '                .Append CMD.CreateParameter("@OPTION2", adVarChar, adParamInput, 7, xOPTION2)
            '            End With
            '
            '            Set rsJournal_HD = CMD.Execute


        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
            ListVoucherNo.Enabled = True
        Else
            ListVoucherNo.Enabled = False
        End If
    Else
        Me.ListVoucherNo.Sorted = False: Me.ListVoucherNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentVend.Value = True Then
            'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_VENDOR.CODE,AMIS_Journal_Hd.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_Journal_Hd.JType,AMIS_Journal_Hd.remarks from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code Where AMIS_Journal_Hd.vendorcode = '" & CURRENT_VENDORCODE & "' AND (Jtype = 'VPJ' OR Jtype = 'APJ' or Jtype = 'VDJ') and status = 'P' and (PaidStatus = 'N' OR AMIS_Journal_Hd.Balance > 0)  and VoucherNo like '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' order by VoucherNo asc")
            'Commented by BTT due to TCN concern
            'Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (SELECT AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.Jdate,AMIS_JOURNAL_HD.vendorcode,AMIS_JOURNAL_HD.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_JOURNAL_HD.jtype,AMIS_JOURNAL_HD.remarks From AMIS_JOURNAL_HD WHERE AMIS_JOURNAL_HD.JTYPE IN('APJ','VPJ','VDJ')AND AMIS_JOURNAL_HD.STATUS ='P') AS T WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' and T.XBALANCE>0 and VOUCHERNO like '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' order by VoucherNo asc")
            Set rsJournal_HD = gconDMIS.Execute("SELECT VOUCHERNO,JDATE,VENDORCODE,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),JTYPE,REMARKS,ACCT_CODE FROM (SELECT HD.VOUCHERNO,HD.JDATE,HD.VENDORCODE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END), " & _
                                                "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AMOUNTPAID, " & _
                                                "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ')AND HD.STATUS ='P') AS T " & _
                                                "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND (AMOUNTTOPAY-AMOUNTPAID) <> 0 AND VOUCHERNO LIKE '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' ORDER BY VOUCHERNO,ACCT_CODE ASC")

        Else
            '            Dim rsCRJ_CDJ As ADODB.Recordset
            '            Set rsCRJ_CDJ = New ADODB.Recordset
            '            rsCRJ_CDJ.Open "SELECT ACCT_CODE FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) = '21-07' AND VOUCHERNO='" & frmAMISJournalEntry_CDJ.txtVoucherNo.Text & "' AND JTYPE='" & xJOURNALTYPE & "' AND DEBIT > 0 ", gconDMIS, adOpenForwardOnly
            '            If Not rsCRJ_CDJ.EOF And Not rsCRJ_CDJ.BOF Then
            '                'CRJ FOR CUSTOMER DEPOSITS=========================================
            '                Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (SELECT HD.VOUCHERNO,HD.JDATE,CUSTOMERCODE,INVOICEAMOUNT=(CASE WHEN HD.JTYPE='COB' THEN HD.INVOICEAMT WHEN HD.JTYPE='CRJ' THEN DET.CREDIT END),(CASE WHEN HD.JTYPE='COB' THEN HD.INVOICEAMT WHEN HD.JTYPE='CRJ' THEN DET.CREDIT END) - " & _
                             '                            "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_CV_DETAIL.JTYPE=HD.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS XBALANCE, " & _
                             '                            "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE INNER JOIN AMIS_CHARTACCOUNT AC ON DET.ACCT_CODE=AC.ACCTCODE WHERE HD.JTYPE IN('CRJ','COB') AND HD.STATUS ='P' AND AC.IS_SCHEDULE_ACCNT=1) AS T " & _
                             '                            "WHERE T.XBALANCE > 0 and VOUCHERNO like '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' ORDER BY VOUCHERNO,JTYPE")
            '            Else
            'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_VENDOR.CODE,AMIS_JOURNAL_HD.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_Journal_Hd.JType,AMIS_Journal_Hd.remarks from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code Where (Jtype = 'VPJ' OR Jtype = 'APJ' or Jtype = 'VDJ') and status = 'P' and (PaidStatus = 'N' OR AMIS_Journal_Hd.Balance > 0)  and VoucherNo like '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' order by VoucherNo asc")
            'Commented by BTT due to TCN concern
            'Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (SELECT AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.Jdate,AMIS_JOURNAL_HD.vendorcode,AMIS_JOURNAL_HD.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_JOURNAL_HD.jtype,AMIS_JOURNAL_HD.remarks From AMIS_JOURNAL_HD WHERE AMIS_JOURNAL_HD.JTYPE IN('APJ','VPJ','VDJ')AND AMIS_JOURNAL_HD.STATUS ='P') AS T WHERE T.XBALANCE>0 and VOUCHERNO like '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' order by VoucherNo asc")
            Set rsJournal_HD = gconDMIS.Execute("SELECT VOUCHERNO,JDATE,VENDORCODE,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),JTYPE,REMARKS,ACCT_CODE FROM (SELECT HD.VOUCHERNO,HD.JDATE,HD.VENDORCODE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END), " & _
                                                "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AMOUNTPAID, " & _
                                                "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ')AND HD.STATUS ='P') AS T " & _
                                                "WHERE (AMOUNTTOPAY-AMOUNTPAID) <> 0 AND VOUCHERNO LIKE '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' ORDER BY VOUCHERNO,ACCT_CODE ASC")
            '            End If

        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
            ListVoucherNo.Enabled = True
        Else
            ListVoucherNo.Enabled = False
        End If
    End If
End Sub

Private Sub txtVendorPayeeName_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtVendorPayeeName.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListVendorPayeeName.Enabled = True Then ListVendorPayeeName.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVendorPayeeName_Change()
'UPDATED BY: ACL 9202010
'DESCRIPTION: ONLY TAG DETAILS WITH SCHEDULE
    If txtVendorPayeeName = "" Then
        Me.ListVendorPayeeName.Sorted = False: Me.ListVendorPayeeName.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentVend.Value = True Then
            'Set rsJournal_HD = gconDMIS.Execute("select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.Balance,AMIS_Journal_Hd.JType from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where AMIS_Journal_Hd.vendorcode = '" & CURRENT_VENDORCODE & "' AND (Jtype = 'VPJ' OR Jtype = 'APJ' OR Jtype = 'VDJ') and status = 'P' and (PaidStatus = 'N' OR AMIS_Journal_Hd.Balance > 0)  order by All_VENDOR.nameofvendor asc")
            'Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (SELECT (SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = VENDORCODE) AS VENDORNAME,AMIS_JOURNAL_HD.Jdate,AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_JOURNAL_HD.jtype,AMIS_JOURNAL_HD.remarks,AMIS_JOURNAL_HD.VENDORCODE From AMIS_JOURNAL_HD WHERE AMIS_JOURNAL_HD.JTYPE IN('APJ','VPJ','VDJ')AND AMIS_JOURNAL_HD.STATUS ='P') AS T WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' and T.XBALANCE>0")
            Set rsJournal_HD = gconDMIS.Execute("SELECT VENDORNAME,JDATE,VOUCHERNO,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),DUEDATE,JTYPE,REMARKS,ACCT_CODE FROM (SELECT (SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = HD.VENDORCODE) AS VENDORNAME, " & _
                                                "HD.JDATE,HD.VENDORCODE,HD.VOUCHERNO,HD.DUEDATE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END), " & _
                                                "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS AMOUNTPAID, " & _
                                                "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ')AND HD.STATUS ='P') AS T " & _
                                                "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND (AMOUNTTOPAY-AMOUNTPAID) <> 0 ORDER BY VOUCHERNO,JTYPE")
        Else
            'Set rsJournal_HD = gconDMIS.Execute("select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.Balance,AMIS_Journal_Hd.JType from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where (Jtype = 'VPJ' OR Jtype = 'APJ' OR Jtype = 'VDJ') and status = 'P' and (PaidStatus = 'N' OR AMIS_Journal_Hd.Balance > 0)  order by All_VENDOR.nameofvendor asc")
            'Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (SELECT (SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = VENDORCODE) AS VENDORNAME,AMIS_JOURNAL_HD.Jdate,AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_JOURNAL_HD.jtype,AMIS_JOURNAL_HD.remarks From AMIS_JOURNAL_HD WHERE AMIS_JOURNAL_HD.JTYPE IN('APJ','VPJ','VDJ')AND AMIS_JOURNAL_HD.STATUS ='P') AS T WHERE T.XBALANCE>0")
            Set rsJournal_HD = gconDMIS.Execute("SELECT VENDORNAME,JDATE,VOUCHERNO,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),DUEDATE,JTYPE,ISNULL(REMARKS,'') AS REMARKS,ACCT_CODE FROM (SELECT (SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = HD.VENDORCODE) AS VENDORNAME, " & _
                                                "HD.JDATE,HD.VENDORCODE,HD.VOUCHERNO,HD.DUEDATE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END), " & _
                                                "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS AMOUNTPAID, " & _
                                                "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ')AND HD.STATUS ='P') AS T " & _
                                                "WHERE (AMOUNTTOPAY-AMOUNTPAID) <> 0 ORDER BY VOUCHERNO,JTYPE")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVendorPayeeName.ListItems, rsJournal_HD
            ListVendorPayeeName.Enabled = True
        Else
            ListVendorPayeeName.Enabled = False
        End If
    Else
        Me.ListVendorPayeeName.Sorted = False: Me.ListVendorPayeeName.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentVend.Value = True Then
            'Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_Journal_Hd.JType,AMIS_Journal_Hd.VENDORCODE,AMIS_Journal_Hd.STATUS from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code ) AS X Where X.vendorcode = '" & CURRENT_VENDORCODE & "' AND (Jtype = 'VPJ' OR Jtype = 'APJ' OR Jtype = 'VDJ') and status = 'P' and X.XBALANCE > 0  and X.nameofvendor like '" & Trim(Me.txtVendorPayeeName) & "%' order by X.nameofvendor asc")
            'Set rsJournal_HD = gconDMIS.Execute(" SELECT * FROM (SELECT (SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = VENDORCODE) AS VENDORNAME, " & _
             "HD.JDATE,HD.VOUCHERNO,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END),(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END) - " & _
             "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS XBALANCE, " & _
             "HD.JTYPE,HD.REMARKS,HD.VENDORCODE,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ','VDJ')AND HD.STATUS ='P') AS T " & _
             "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND T.XBALANCE <>0 AND VENDORNAME like '" & Format(Trim(Repleys(Me.txtVendorPayeeName)), "000000") & "%' ORDER BY VOUCHERNO,JTYPE")

            Set rsJournal_HD = gconDMIS.Execute("SELECT VENDORNAME,JDATE,VOUCHERNO,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),DUEDATE,JTYPE,REMARKS,ACCT_CODE FROM (SELECT (SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = HD.VENDORCODE) AS VENDORNAME, " & _
                                                "HD.JDATE,HD.VENDORCODE,HD.VOUCHERNO,HD.DUEDATE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END), " & _
                                                "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS AMOUNTPAID, " & _
                                                "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ')AND HD.STATUS ='P') AS T " & _
                                                "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND (AMOUNTTOPAY-AMOUNTPAID) <> 0 AND VENDORNAME like '" & Format(Trim(Repleys(Me.txtVendorPayeeName)), "000000") & "%' ORDER BY VOUCHERNO,JTYPE")
        Else
            'Set rsJournal_HD = gconDMIS.Execute("SELECT * FROM (select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.AMOUNTTOPAY,AMIS_JOURNAL_HD.AMOUNTTOPAY - ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (AMIS_JOURNAL_HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO)) ,0) AS XBALANCE,AMIS_Journal_Hd.JType,AMIS_Journal_Hd.VENDORCODE,AMIS_Journal_Hd.STATUS from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code ) AS X Where (Jtype = 'VPJ' OR Jtype = 'APJ' OR Jtype = 'VDJ') and status = 'P' and X.XBalance > 0 and X.nameofvendor like '" & Trim(Me.txtVendorPayeeName) & "%' order by X.nameofvendor asc")
            Set rsJournal_HD = gconDMIS.Execute("SELECT VENDORNAME,JDATE,VOUCHERNO,AMOUNTTOPAY,BALANCE=(AMOUNTTOPAY-AMOUNTPAID),DUEDATE,JTYPE,REMARKS,ACCT_CODE FROM (SELECT (SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = HD.VENDORCODE) AS VENDORNAME, " & _
                                                "HD.JDATE,HD.VENDORCODE,HD.VOUCHERNO,HD.DUEDATE,AMOUNTTOPAY=(CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY WHEN HD.JTYPE='APJ' THEN DET.CREDIT END), " & _
                                                "ISNULL((SELECT SUM(AMOUNT) FROM AMIS_CV_DETAIL WHERE (HD.JTYPE=AMIS_CV_DETAIL.JTYPE  AND PV_VOUCHERNO=HD.VOUCHERNO AND J_CLASS=DET.ACCT_CODE) GROUP BY J_CLASS) ,0) AS AMOUNTPAID, " & _
                                                "HD.JTYPE,HD.REMARKS,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.JTYPE IN('APJ','VPJ')AND HD.STATUS ='P') AS T " & _
                                                "WHERE (AMOUNTTOPAY-AMOUNTPAID) <> 0 AND VENDORNAME like '" & Format(Trim(Repleys(Me.txtVendorPayeeName)), "000000") & "%' ORDER BY VOUCHERNO,JTYPE")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVendorPayeeName.ListItems, rsJournal_HD
            ListVendorPayeeName.Enabled = True
        Else
            ListVendorPayeeName.Enabled = False
        End If
    End If
End Sub

Function ReturnVendor(XXX As String) As String
    Dim RS                                             As New ADODB.Recordset
    Set RS = gconDMIS.Execute("Select nameofvendor,code from ALL_vendor where code='" & XXX & "'")
    If Not (RS.EOF And RS.BOF) Then
        ReturnVendor = Null2String(RS!nameofvendor)
    Else
        ReturnVendor = ""
    End If
    Set RS = Nothing
End Function


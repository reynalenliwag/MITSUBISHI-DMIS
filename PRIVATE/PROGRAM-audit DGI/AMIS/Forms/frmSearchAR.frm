VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
Begin VB.Form frmAMISSearchAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Sales Journals"
   ClientHeight    =   6375
   ClientLeft      =   2970
   ClientTop       =   3495
   ClientWidth     =   8700
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
   Icon            =   "frmSearchAR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8700
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
      ItemCount       =   3
      Item(0).Caption =   "By &Voucher No"
      Item(0).Tooltip =   "Search Sales Journals by Invoice Number"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By Entity Name"
      Item(1).Tooltip =   "Search Sales Journals by Customer Name"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "By Invoice No"
      Item(2).Tooltip =   "Search Sales Journals by Invoice Number"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   5325
         Left            =   -69970
         TabIndex        =   1
         Top             =   30
         Visible         =   0   'False
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll3 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7560
            TabIndex        =   14
            Top             =   120
            Width           =   1035
         End
         Begin VB.TextBox txtInvoiceNo 
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
            TabIndex        =   2
            Top             =   45
            Width           =   6240
         End
         Begin MSComctlLib.ListView ListInvoiceNo 
            Height          =   4875
            Left            =   30
            TabIndex        =   16
            Top             =   450
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   8599
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
            MouseIcon       =   "frmSearchAR.frx":000C
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "INVOICE NO "
               Object.Width           =   3422
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "JTYPE"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "VOUCHER NO."
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ENTITY NAME"
               Object.Width           =   5291
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "INVOICE TYPE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "INVOICE NO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "INVOICE DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "INVOICE AMT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "BALANCE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "CODE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "ENTITY CODE"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label3 
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
            TabIndex        =   3
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   5325
         Left            =   -69970
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll2 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7560
            TabIndex        =   13
            Top             =   120
            Width           =   1035
         End
         Begin VB.TextBox txtCustomer 
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
            TabIndex        =   5
            Top             =   45
            Width           =   6240
         End
         Begin MSComctlLib.ListView ListCustomer 
            Height          =   4875
            Left            =   30
            TabIndex        =   15
            Top             =   450
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   8599
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
            MouseIcon       =   "frmSearchAR.frx":016E
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ENTITY NAME"
               Object.Width           =   5291
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "JTYPE"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "VOUCHER NO."
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "INVOICE NO "
               Object.Width           =   3422
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "INVOICE TYPE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "INVOICE NO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "INVOICE DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "INVOICE AMT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "BALANCE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "CODE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "ENTITY CODE"
               Object.Width           =   0
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
            TabIndex        =   6
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5325
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7560
            TabIndex        =   12
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
            TabIndex        =   9
            Top             =   45
            Width           =   6240
         End
         Begin MSComctlLib.ListView ListVoucherNo 
            Height          =   4875
            Left            =   30
            TabIndex        =   10
            Top             =   450
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   8599
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
            MouseIcon       =   "frmSearchAR.frx":02D0
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "JTYPE"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "VOUCHER NO."
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "ENTITY NAME"
               Object.Width           =   5291
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "INVOICE NO "
               Object.Width           =   3422
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "INVOICE TYPE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "INVOICE NO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "INVOICE DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "INVOICE AMT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "BALANCE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "CODE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "ENTITY CODE"
               Object.Width           =   0
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
            TabIndex        =   8
            Top             =   90
            Width           =   1125
         End
      End
   End
   Begin MSForms.CheckBox chkCurrentCust 
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   30
      Width           =   8625
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   4
      Size            =   "15214;556"
      Value           =   "1"
      Caption         =   "Show Journals for Current Customer Only"
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmAMISSearchAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                            As New ADODB.Recordset
Dim Y, k                                                    As Long
Attribute k.VB_VarUserMemId = 1073938433
Dim xJOURNALTYPE                                            As String
Dim ACCT_CODE                                               As String

Sub LOADJOURNAL(XXX As String)
    xJOURNALTYPE = XXX
End Sub
Sub clearListView()
    For Y = 1 To Me.ListVoucherNo.ListItems.Count
        If Me.ListVoucherNo.ListItems.Count <= 0 Then Exit For
        Me.ListVoucherNo.Sorted = False
        Me.ListVoucherNo.ListItems.Remove Me.ListVoucherNo.SelectedItem.Index
    Next Y
    For Y = 1 To Me.ListCustomer.ListItems.Count
        If Me.ListCustomer.ListItems.Count <= 0 Then Exit For
        Me.ListCustomer.Sorted = False
        Me.ListCustomer.ListItems.Remove Me.ListCustomer.SelectedItem.Index
    Next Y
    For Y = 1 To Me.ListInvoiceNo.ListItems.Count
        If Me.ListInvoiceNo.ListItems.Count <= 0 Then Exit For
        Me.ListInvoiceNo.Sorted = False
        Me.ListInvoiceNo.ListItems.Remove Me.ListInvoiceNo.SelectedItem.Index
    Next Y
End Sub

Private Sub chkShowAll2_Click()
    If SEARCH_TAB = 1 Then txtCustomer_Change
End Sub

Private Sub chkCurrentCust_Click()
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub chkShowAll_Click()
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Select Case SEARCH_TAB
        Case 0:
            On Error Resume Next
            txtVoucherNo.SetFocus
        Case 1:
            On Error Resume Next
            txtCustomer.SetFocus
        Case 2:
            On Error Resume Next
            txtInvoiceNo.SetFocus
        End Select
    End If
    If Shift = 2 Then
        Select Case KeyCode
        Case vbKeyV: SearchTab.SelectedItem = 0
            txtVoucherNo_Change
        Case vbKeyP: SearchTab.SelectedItem = 1
            txtCustomer_Change
        Case vbKeyI: SearchTab.SelectedItem = 2
            txtInvoiceNo_Change
        End Select
        On Error Resume Next
        SEARCH_TAB = SearchTab.SelectedItem: SearchTab_SelectedChanged (SearchTab.Selected)
    End If
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SearchTab.SelectedItem = SEARCH_TAB
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub ListCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListCustomer
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub ListCustomer_DblClick()
    frmAMISJournalEntry_DetailPayment.txtInvoiceNo.Text = (Trim(Me.ListCustomer.SelectedItem.SubItems(6)))
    frmAMISJournalEntry_DetailPayment.txtInvoiceDate.Text = (Trim(Me.ListCustomer.SelectedItem.SubItems(7)))
    frmAMISJournalEntry_DetailPayment.txtInvoiceAmt = (Trim(Me.ListCustomer.SelectedItem.SubItems(9)))
    frmAMISJournalEntry_DetailPayment.txtCusCde.Text = (Trim(Me.ListCustomer.SelectedItem.SubItems(10)))
    frmAMISJournalEntry_DetailPayment.txtCustomerName.Text = (Trim(Me.ListCustomer.SelectedItem.Text))
    frmAMISJournalEntry_DetailPayment.lblInvoiceType.Caption = (Trim(Me.ListCustomer.SelectedItem.SubItems(5)))
    frmAMISJournalEntry_DetailPayment.lblBalance.Caption = (Trim(Me.ListCustomer.SelectedItem.SubItems(9)))
    frmAMISJournalEntry_DetailPayment.lblSJVoucherNo.Caption = (Trim(Me.ListCustomer.SelectedItem.SubItems(1))) + "-" + (Trim(Me.ListCustomer.SelectedItem.SubItems(2)))
    frmAMISJournalEntry_DetailPayment.lblEntityClass.Caption = (Trim(Me.ListCustomer.SelectedItem.SubItems(11)))
    Unload Me
End Sub

Private Sub ListInvoiceNo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListInvoiceNo
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
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
    frmAMISJournalEntry_DetailPayment.txtInvoiceNo.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(6)))
    frmAMISJournalEntry_DetailPayment.txtInvoiceDate.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(7)))
    frmAMISJournalEntry_DetailPayment.txtInvoiceAmt = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(9)))
    frmAMISJournalEntry_DetailPayment.txtCusCde.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(10)))
    frmAMISJournalEntry_DetailPayment.txtCustomerName.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(3)))
    frmAMISJournalEntry_DetailPayment.lblInvoiceType.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(5)))
    frmAMISJournalEntry_DetailPayment.lblBalance.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(9)))
    frmAMISJournalEntry_DetailPayment.lblSJVoucherNo.Caption = (Trim(Me.ListVoucherNo.SelectedItem.Text)) + "-" + (Trim(Me.ListVoucherNo.SelectedItem.SubItems(1)))
    frmAMISJournalEntry_DetailPayment.lblEntityClass.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(11)))
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
        ListVoucherNo_DblClick
    End If
End Sub

Private Sub ListCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtCustomer.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListCustomer_DblClick
    End If
End Sub

Private Sub ListInvoiceNo_DblClick()
    frmAMISJournalEntry_DetailPayment.txtInvoiceNo.Text = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(6)))
    frmAMISJournalEntry_DetailPayment.txtInvoiceDate.Text = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(7)))
    frmAMISJournalEntry_DetailPayment.txtInvoiceAmt = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(9)))
    frmAMISJournalEntry_DetailPayment.txtCusCde.Text = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(10)))
    frmAMISJournalEntry_DetailPayment.txtCustomerName.Text = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(4)))
    frmAMISJournalEntry_DetailPayment.lblInvoiceType.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(5)))
    frmAMISJournalEntry_DetailPayment.lblBalance.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(9)))
    frmAMISJournalEntry_DetailPayment.lblSJVoucherNo.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(1))) + "-" + (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(2)))
    frmAMISJournalEntry_DetailPayment.lblEntityClass.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(11)))
    Unload Me
End Sub

Private Sub ListInvoiceNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtInvoiceNo.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListInvoiceNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListInvoiceNo_DblClick
    End If
End Sub

Private Sub SearchTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    SEARCH_TAB = SearchTab.SelectedItem
    DoEvents
    txtVoucherNo.Enabled = False
    txtCustomer.Enabled = False
    txtInvoiceNo.Enabled = False
    ListVoucherNo.Enabled = False
    ListCustomer.Enabled = False
    ListInvoiceNo.Enabled = False
    Select Case SEARCH_TAB
    Case 0
        txtVoucherNo.Enabled = True: ListVoucherNo.Enabled = True
        Me.Caption = "Search Item by Voucher Number"
        On Error Resume Next
        txtVoucherNo.SetFocus
        chkCurrentCust_Click
    Case 1
        txtCustomer.Enabled = True: ListCustomer.Enabled = True
        Me.Caption = "Search Item by Customer/Payee Name"
        On Error Resume Next
        txtCustomer.SetFocus
        chkCurrentCust_Click
    Case 2
        txtInvoiceNo.Enabled = True: ListInvoiceNo.Enabled = True
        Me.Caption = "Search Item by Invoice Number"
        On Error Resume Next
        txtInvoiceNo.SetFocus
        chkCurrentCust_Click
    End Select
End Sub

Private Sub txtInvoiceNo_Change()
    If txtInvoiceNo = "" Then
        ListInvoiceNo.Enabled = False
        Me.ListInvoiceNo.Sorted = False: Me.ListInvoiceNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentCust.Value = True Then
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0  AND CUSTOMERCODE = '" & CURRENT_CUSCODE & "' " & _
                                                    " ORDER BY INVOICENO,INVOICETYPE,JTYPE,VOUCHERNO,ACCOUNTNAME ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND VENDORCODE = '" & CURRENT_CUSCODE & "' " & _
                                                    " ORDER BY INVOICENO,INVOICETYPE,JTYPE,VOUCHERNO,ACCOUNTNAME ASC ")
            End If
        Else
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 " & _
                                                    " ORDER BY INVOICENO,INVOICETYPE,JTYPE,VOUCHERNO,ACCOUNTNAME ASC ")

            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 " & _
                                                    " ORDER BY INVOICENO,INVOICETYPE,JTYPE,VOUCHERNO,ACCOUNTNAME ASC ")
            End If
        End If

        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListInvoiceNo.Enabled = True
            Listview_Loadval Me.ListInvoiceNo.ListItems, rsJournal_HD
        Else
            ListInvoiceNo.Enabled = False
        End If
    Else
        Me.ListInvoiceNo.Sorted = False: Me.ListInvoiceNo.ListItems.Clear: chkCurrentCust.Value = False
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentCust.Value = True Then
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND CUSTOMERCODE = '" & CURRENT_CUSCODE & "' AND INVOICENO LIKE '" & Format(Replace(Trim(Me.txtInvoiceNo), "'", ""), "000000") & "%' " & _
                                                    " ORDER BY INVOICENO,INVOICETYPE,JTYPE,VOUCHERNO,ACCOUNTNAME ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND VENDORCODE = '" & CURRENT_CUSCODE & "' AND INVOICENO LIKE '" & Format(Replace(Trim(Me.txtInvoiceNo), "'", ""), "000000") & "%'" & _
                                                    " ORDER BY INVOICENO,INVOICETYPE,JTYPE,VOUCHERNO,ACCOUNTNAME ASC ")
            End If
        Else
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0  AND INVOICENO LIKE '" & Format(Replace(Trim(Me.txtInvoiceNo), "'", ""), "000000") & "%' " & _
                                                    " ORDER BY INVOICENO,INVOICETYPE,JTYPE,VOUCHERNO,ACCOUNTNAME ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND INVOICENO LIKE '" & Format(Replace(Trim(Me.txtInvoiceNo), "'", ""), "000000") & "%'" & _
                                                    " ORDER BY INVOICENO,INVOICETYPE,JTYPE,VOUCHERNO,ACCOUNTNAME ASC ")
            End If
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListInvoiceNo.Enabled = True
            Listview_Loadval Me.ListInvoiceNo.ListItems, rsJournal_HD
        Else
            ListInvoiceNo.Enabled = False
        End If
    End If
End Sub

Private Sub txtInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtInvoiceNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListInvoiceNo.ListItems.Count > 0 And ListInvoiceNo.Enabled = True Then: ListInvoiceNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVoucherNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtVoucherNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListVoucherNo.ListItems.Count > 0 And ListVoucherNo.Enabled = True Then: ListVoucherNo.SetFocus

    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVoucherNo_Change()
    If txtVoucherNo = "" Then
        ListVoucherNo.Enabled = False
        Me.ListVoucherNo.Sorted = False: Me.ListVoucherNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentCust.Value = True Then
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0  AND CUSTOMERCODE = '" & CURRENT_CUSCODE & "' " & _
                                                    " ORDER BY JTYPE,VOUCHERNO ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND VENDORCODE = '" & CURRENT_CUSCODE & "' " & _
                                                    " ORDER BY JTYPE,VOUCHERNO ASC ")
            End If
        Else
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 " & _
                                                    " ORDER BY JTYPE,VOUCHERNO ASC ")

            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 " & _
                                                    " ORDER BY JTYPE,VOUCHERNO ASC ")
            End If
        End If

        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListVoucherNo.Enabled = True
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
        Else
            ListVoucherNo.Enabled = False
        End If
    Else
        Me.ListVoucherNo.Sorted = False: Me.ListVoucherNo.ListItems.Clear: chkCurrentCust.Value = False
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentCust.Value = True Then
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND CUSTOMERCODE = '" & CURRENT_CUSCODE & "' AND VOUCHERNO LIKE '" & Format(Replace(Trim(Me.txtVoucherNo), "'", ""), "000000") & "%' " & _
                                                    " ORDER BY JTYPE,VOUCHERNO ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND VENDORCODE = '" & CURRENT_CUSCODE & "' AND VOUCHERNO LIKE '" & Format(Replace(Trim(Me.txtVoucherNo), "'", ""), "000000") & "%'" & _
                                                    " ORDER BY JTYPE,VOUCHERNO ASC ")
            End If
        Else
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0  AND VOUCHERNO LIKE '" & Format(Replace(Trim(Me.txtVoucherNo), "'", ""), "000000") & "%' " & _
                                                    " ORDER BY JTYPE,VOUCHERNO ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND VOUCHERNO LIKE '" & Format(Replace(Trim(Me.txtVoucherNo), "'", ""), "000000") & "%'" & _
                                                    " ORDER BY JTYPE,VOUCHERNO ASC ")
            End If
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListVoucherNo.Enabled = True
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
        Else
            ListVoucherNo.Enabled = False
        End If
    End If
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCustomer.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListCustomer.ListItems.Count > 0 And ListCustomer.Enabled = True Then: ListCustomer.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCustomer_Change()
    If txtCustomer = "" Then
        Me.ListCustomer.Sorted = False: Me.ListCustomer.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentCust.Value = True Then
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0  AND CUSTOMERCODE = '" & CURRENT_CUSCODE & "' " & _
                                                    " ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND VENDORCODE = '" & CURRENT_CUSCODE & "' " & _
                                                    " ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO ASC ")
            End If
        Else
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 " & _
                                                    " ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO ASC ")

            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 " & _
                                                    " ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO ASC ")
            End If
        End If

        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListCustomer.Enabled = True
            Listview_Loadval Me.ListCustomer.ListItems, rsJournal_HD
        Else
            ListCustomer.Enabled = False
        End If
    Else
        Me.ListCustomer.Sorted = False: Me.ListCustomer.ListItems.Clear: chkCurrentCust.Value = False
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentCust.Value = True Then
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND CUSTOMERCODE = '" & CURRENT_CUSCODE & "' AND ACCOUNTNAME LIKE '" & Format(Replace(Trim(Me.txtCustomer), "'", ""), "000000") & "%' " & _
                                                    " ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND VENDORCODE = '" & CURRENT_CUSCODE & "' AND ACCOUNTNAME LIKE '" & Format(Replace(Trim(Me.txtCustomer), "'", ""), "000000") & "%'" & _
                                                    " ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO ASC ")
            End If
        Else
            If CheckIfARAccount(N2Str2Null(ACCT_CODE)) = True Then
                Set rsJournal_HD = gconDMIS.Execute("SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,CUSTOMERCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(SJVOUCHERNO)=10 THEN LEFT(SJVOUCHERNO,3) ELSE LEFT(SJVOUCHERNO,2) END AS JTYPE,RIGHT(SJVOUCHERNO,6) AS VOUCHERNO,JDATE,CUSTOMERCODE,CUSTOMERNAME AS ACCOUNTNAME,INVOICETYPE + ' ' + INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT_TOPAY AS INVOICEAMT,ENTITYCODE,AMOUNT_TOPAY - ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL AD WHERE (AR.INVOICETYPE=AD.INVOICETYPE AND AR.INVOICENO=AD.INVOICENO AND AR.CUSTOMERCODE=AD.CUSTOMERCODE AND AR.ACCOUNT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AR AR WHERE STATUS='P' AND ACCOUNT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0  AND ACCOUNTNAME LIKE '" & Format(Replace(Trim(Me.txtCustomer), "'", ""), "000000") & "%' " & _
                                                    " ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO ASC ")
            Else
                Set rsJournal_HD = gconDMIS.Execute("SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                    " FROM (" & _
                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                    " FROM AMIS_AP AP WHERE STATUS='P' AND ACCT_CODE = " & N2Str2Null(ACCT_CODE) & ") T " & _
                                                    " WHERE BALANCE <> 0 AND ACCOUNTNAME LIKE '" & Format(Replace(Trim(Me.txtCustomer), "'", ""), "000000") & "%'" & _
                                                    " ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO ASC ")
            End If
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListCustomer.Enabled = True
            Listview_Loadval Me.ListCustomer.ListItems, rsJournal_HD
        Else
            ListCustomer.Enabled = False
        End If
    End If
End Sub


Sub CHECKACCOUNT(XXX As String)
    ACCT_CODE = XXX
End Sub

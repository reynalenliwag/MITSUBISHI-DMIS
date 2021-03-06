VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmAMISSearchSJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Sales Journals"
   ClientHeight    =   6450
   ClientLeft      =   2970
   ClientTop       =   3495
   ClientWidth     =   8775
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
   Icon            =   "SearchSJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8775
   Begin VB.OptionButton optCancelled 
      Caption         =   "Cancelled Journals"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   5220
      TabIndex        =   2
      Top             =   60
      Width           =   3795
   End
   Begin VB.OptionButton optUnPosted 
      Caption         =   "Un-Posted Journals"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Top             =   60
      Width           =   3795
   End
   Begin VB.OptionButton optPosted 
      Caption         =   "Posted Journals"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Value           =   -1  'True
      Width           =   3255
   End
   Begin XtremeSuiteControls.TabControl SearchTab 
      Height          =   5955
      Left            =   30
      TabIndex        =   3
      Top             =   450
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
      Item(0).Tooltip =   "Search Sales Journals by Voucher Number"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By Customer Name"
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
         TabIndex        =   4
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
            TabIndex        =   18
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
            Left            =   1350
            TabIndex        =   5
            Top             =   0
            Width           =   6135
         End
         Begin MSComctlLib.ListView ListInvoiceNo 
            Height          =   4875
            Left            =   0
            TabIndex        =   7
            Top             =   420
            Width           =   8625
            _ExtentX        =   15214
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
            MouseIcon       =   "SearchSJ.frx":000C
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "INVOICE TYPE WITH NO."
               Object.Width           =   3598
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   8995
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "VOUCHER NO."
               Object.Width           =   2822
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
            TabIndex        =   6
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   5325
         Left            =   -69970
         TabIndex        =   8
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
            Left            =   7530
            TabIndex        =   17
            Top             =   90
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
            Left            =   1350
            TabIndex        =   9
            Top             =   0
            Width           =   6135
         End
         Begin MSComctlLib.ListView ListCustomer 
            Height          =   4890
            Left            =   0
            TabIndex        =   11
            Top             =   420
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   8625
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
            MouseIcon       =   "SearchSJ.frx":016E
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CUSTOMER NAME"
               Object.Width           =   8995
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
               Text            =   "PAY TYPE AND OR NO."
               Object.Width           =   3598
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
            TabIndex        =   10
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5325
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7530
            TabIndex        =   16
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
            TabIndex        =   14
            Top             =   45
            Width           =   6270
         End
         Begin MSComctlLib.ListView ListVoucherNo 
            Height          =   4875
            Left            =   45
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
            MouseIcon       =   "SearchSJ.frx":02D0
            NumItems        =   3
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
               Text            =   "CUSTOMER NAME"
               Object.Width           =   12347
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
            TabIndex        =   13
            Top             =   90
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frmAMISSearchSJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As New ADODB.Recordset
Dim Y, k                                          As Long
Attribute k.VB_VarUserMemId = 1073938433
Dim StatusToSearch                                As String

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

Private Sub chkShowAll_Click()
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
End Sub

Private Sub chkShowAll2_Click()
    If SEARCH_TAB = 1 Then txtCustomer_Change
End Sub

Private Sub chkShowAll3_Click()
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
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
        Case vbKeyV: SearchTab.SelectedItem = 0: txtVoucherNo_Change
        Case vbKeyP: SearchTab.SelectedItem = 1: txtCustomer_Change
        Case vbKeyI: SearchTab.SelectedItem = 2: txtInvoiceNo_Change
        End Select
        On Error Resume Next
        SEARCH_TAB = SearchTab.SelectedItem: SearchTab_SelectedChanged (SearchTab.Selected)
    End If
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    StatusToSearch = "P"
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
    On Error Resume Next
    'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListCustomer.SelectedItem.SubItems(2)))
    frmAMISJournalEntry_SJ.LoadJournal ("SJ")
    frmAMISJournalEntry_SJ.SearchVoucherNo (Trim(Me.ListCustomer.SelectedItem.SubItems(2)))
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

Private Sub ListInvoiceNo_DblClick()
    On Error Resume Next
    'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(3)))
    frmAMISJournalEntry_SJ.LoadJournal ("SJ")
    frmAMISJournalEntry_SJ.SearchVoucherNo (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(3)))
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
    'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
    frmAMISJournalEntry_SJ.LoadJournal ("SJ")
    frmAMISJournalEntry_SJ.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
    Unload Me
End Sub

Private Sub ListVoucherNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txtVoucherNo.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListVoucherNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
        frmAMISJournalEntry_SJ.LoadJournal ("SJ")
        frmAMISJournalEntry_SJ.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
        Unload Me
    End If
End Sub

Private Sub ListCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtCustomer.SetFocus: SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListCustomer.SelectedItem.SubItems(2)))
        frmAMISJournalEntry_SJ.LoadJournal ("SJ")
        frmAMISJournalEntry_SJ.SearchVoucherNo (Trim(Me.ListCustomer.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub ListInvoiceNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtInvoiceNo.SetFocus: SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListInvoiceNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(3)))
        frmAMISJournalEntry_SJ.LoadJournal ("SJ")
        frmAMISJournalEntry_SJ.SearchVoucherNo (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(3)))
        Unload Me
    End If
End Sub

Private Sub optCancelled_Click()
    StatusToSearch = "C"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub optCancelled_GotFocus()
    StatusToSearch = "C"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub optPosted_Click()
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub optPosted_GotFocus()
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub optUnPosted_Click()
    StatusToSearch = "N"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub optUnPosted_GotFocus()
    StatusToSearch = "N"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub SearchTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    SEARCH_TAB = SearchTab.SelectedItem
    DoEvents
    txtVoucherNo.Enabled = False
    txtCustomer.Enabled = False
    ListVoucherNo.Enabled = False
    ListCustomer.Enabled = False
    ListInvoiceNo.Enabled = False
    ListInvoiceNo.Enabled = False
    Select Case SEARCH_TAB
    Case 0
        txtVoucherNo.Enabled = True: ListVoucherNo.Enabled = True
        Me.Caption = "Search Item by Voucher Number"
        On Error Resume Next
        txtVoucherNo.SetFocus
    Case 1
        txtCustomer.Enabled = True: ListCustomer.Enabled = True
        Me.Caption = "Search Item by Customer Name"
        On Error Resume Next
        txtCustomer.SetFocus
    Case 2
        txtInvoiceNo.Enabled = True: ListInvoiceNo.Enabled = True
        Me.Caption = "Search Item by Invoice No."
        On Error Resume Next
        txtInvoiceNo.SetFocus
    End Select
End Sub

Private Sub txtInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtInvoiceNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListInvoiceNo.ListItems.Count > 0 And ListInvoiceNo.Enabled = True Then: ListInvoiceNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtInvoiceNo_Change()
    If txtInvoiceNo = "" Then
        ListInvoiceNo.Enabled = False
        Me.ListInvoiceNo.Sorted = False: Me.ListInvoiceNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll3.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO,ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo from AMIS_Journal_HD LEFT join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' order by AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO,ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo from AMIS_Journal_HD LEFT join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' order by AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListInvoiceNo.Enabled = True
            Listview_Loadval Me.ListInvoiceNo.ListItems, rsJournal_HD
        Else
            ListInvoiceNo.Enabled = False
        End If
    Else
        Me.ListInvoiceNo.Sorted = False: Me.ListInvoiceNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll3.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO,ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo from AMIS_Journal_HD LEFT join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' and AMIS_Journal_Hd.InvoiceNo like '" & Trim(Me.txtInvoiceNo) & "%' order by AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO,ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo from AMIS_Journal_HD LEFT join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' and AMIS_Journal_Hd.InvoiceNo like '" & Replace(Trim(Me.txtInvoiceNo), "'", "") & "%' order by AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListInvoiceNo.Enabled = True
            Listview_Loadval Me.ListInvoiceNo.ListItems, rsJournal_HD
        Else
            ListInvoiceNo.Enabled = False
        End If
    End If
End Sub

Private Sub txtVoucherNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtVoucherNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
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
        If chkShowAll.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO  from AMIS_Journal_HD inner join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' order by VoucherNo asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO  from AMIS_Journal_HD inner join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' order by VoucherNo asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListVoucherNo.Enabled = True
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
        Else
            ListVoucherNo.Enabled = False
        End If
    Else
        Me.ListVoucherNo.Sorted = False: Me.ListVoucherNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO from AMIS_Journal_HD inner join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' and VoucherNo like '" & Format(Replace(Trim(Me.txtVoucherNo), "'", ""), "000000") & "%' order by VoucherNo asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO from AMIS_Journal_HD inner join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' and VoucherNo like '" & Format(Replace(Trim(Me.txtVoucherNo), "'", ""), "000000") & "%' order by VoucherNo asc")
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
        ListCustomer.Enabled = False
        Me.ListCustomer.Sorted = False: Me.ListCustomer.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll2.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO from AMIS_Journal_HD inner join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' order by ALL_CUSTMASTER_AMIS.CustName asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO from AMIS_Journal_HD inner join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' order by ALL_CUSTMASTER_AMIS.CustName asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListCustomer.Enabled = True
            Listview_Loadval Me.ListCustomer.ListItems, rsJournal_HD
        Else
            ListCustomer.Enabled = False
        End If
    Else
        Me.ListCustomer.Sorted = False: Me.ListCustomer.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll2.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO from AMIS_Journal_HD inner join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' and ALL_CUSTMASTER_AMIS.CustName like '" & Trim(Me.txtCustomer) & "%' order by ALL_CUSTMASTER_AMIS.CustName asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 ALL_CUSTMASTER_AMIS.CustName,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.Invoicetype + ' ' + AMIS_Journal_Hd.InvoiceNo as INVOICETYPEWITHNO from AMIS_Journal_HD inner join ALL_CUSTMASTER_AMIS on AMIS_Journal_Hd.customercode = ALL_CUSTMASTER_AMIS.CustCode where AMIS_Journal_Hd.jtype='SJ' and status = '" & StatusToSearch & "' and ALL_CUSTMASTER_AMIS.CustName like '" & Replace(Trim(Me.txtCustomer), "'", "") & "%' order by ALL_CUSTMASTER_AMIS.CustName asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListCustomer.Enabled = True
            Listview_Loadval Me.ListCustomer.ListItems, rsJournal_HD
        Else
            ListCustomer.Enabled = False
        End If
    End If
End Sub


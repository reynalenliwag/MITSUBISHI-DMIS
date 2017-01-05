VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmAMISSearchAPJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Accounts Payable Journals"
   ClientHeight    =   6435
   ClientLeft      =   2970
   ClientTop       =   3735
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
   Icon            =   "SearchAPJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8715
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
      Left            =   5250
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
      Left            =   2430
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Value           =   -1  'True
      Width           =   3255
   End
   Begin XtremeSuiteControls.TabControl SearchTab 
      Height          =   5955
      Left            =   0
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
      ItemCount       =   4
      Item(0).Caption =   "By &Voucher No"
      Item(0).Tooltip =   "Search Accounts Payable Journal by Voucher Number"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By Vendor/&Payee Name"
      Item(1).Tooltip =   "Search Accounts Payable Journal by Vendor/Payee Name"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "By Invoice No."
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Item(3).Caption =   "By RR NO"
      Item(3).Tooltip =   "Search By RR NO"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage4"
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
            Left            =   7530
            TabIndex        =   21
            Top             =   150
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
            Left            =   1170
            TabIndex        =   6
            Top             =   45
            Width           =   6315
         End
         Begin MSComctlLib.ListView ListVendorPayeeName 
            Height          =   4815
            Left            =   45
            TabIndex        =   7
            Top             =   495
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
            MouseIcon       =   "SearchAPJ.frx":000C
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VENDOR NAME"
               Object.Width           =   8819
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
               Text            =   "INV NO."
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
            TabIndex        =   5
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5325
         Left            =   30
         TabIndex        =   8
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
            TabIndex        =   20
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
            TabIndex        =   10
            Top             =   45
            Width           =   6270
         End
         Begin MSComctlLib.ListView ListVoucherNo 
            Height          =   4845
            Left            =   45
            TabIndex        =   11
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
            MouseIcon       =   "SearchAPJ.frx":016E
            NumItems        =   4
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
               Text            =   "VENDOR NAME"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "INV. NO"
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
            TabIndex        =   9
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   5325
         Left            =   -69970
         TabIndex        =   12
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
            TabIndex        =   22
            Top             =   150
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
            Left            =   1170
            TabIndex        =   13
            Top             =   60
            Width           =   6345
         End
         Begin MSComctlLib.ListView ListInvoiceNo 
            Height          =   4755
            Left            =   45
            TabIndex        =   14
            Top             =   495
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   8387
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
            MouseIcon       =   "SearchAPJ.frx":02D0
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "INV. NO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PO NO."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "VENDOR NAME"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
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
            TabIndex        =   15
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   5325
         Left            =   -69970
         TabIndex        =   16
         Top             =   30
         Visible         =   0   'False
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll4 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7560
            TabIndex        =   23
            Top             =   150
            Width           =   1035
         End
         Begin VB.TextBox txtsearch 
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
            Left            =   1170
            TabIndex        =   17
            Top             =   60
            Width           =   6345
         End
         Begin MSComctlLib.ListView ListRR 
            Height          =   4755
            Left            =   60
            TabIndex        =   18
            Top             =   495
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   8387
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
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
            MouseIcon       =   "SearchAPJ.frx":0432
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "RR No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "INVOICE NO"
               Object.Width           =   8819
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
         Begin VB.Label Label4 
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
            TabIndex        =   19
            Top             =   90
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frmAMISSearchAPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                       As New ADODB.Recordset
Dim Y, k                                               As Long
Attribute k.VB_VarUserMemId = 1073938433
Dim StatusToSearch                                     As String

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
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
End Sub

Private Sub chkShowAll3_Click()
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub chkShowAll4_Click()
    If SEARCH_TAB = 3 Then txtSearch_Change
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
        Case 2:
            On Error Resume Next
            txtInvoiceNo.SetFocus
        End Select
    End If
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    StatusToSearch = "P"
    SearchTab.SelectedItem = SEARCH_TAB
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
End Sub

Private Sub ListRR_DblClick()
'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListRR.SelectedItem.SubItems(3)))
    Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
    Call frmAMISJournalEntry_APJ.SearchVoucherNo(Trim(Me.ListRR.SelectedItem.SubItems(3)))
    Unload Me
End Sub

Private Sub ListRR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
        Call frmAMISJournalEntry_APJ.SearchVoucherNo(Trim(Me.ListRR.SelectedItem.SubItems(3)))
        Unload Me
    End If
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

Private Sub ListVendorPayeeName_DblClick()
'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
    Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
    Call frmAMISJournalEntry_APJ.SearchVoucherNo(Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListInvoiceNo_DblClick()

'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(3)))
    Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
    Call frmAMISJournalEntry_APJ.SearchVoucherNo(Trim(Me.ListInvoiceNo.SelectedItem.SubItems(4)))
    Unload Me
End Sub

Private Sub ListVoucherNo_DblClick()
'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
    Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
    Call frmAMISJournalEntry_APJ.SearchVoucherNo(Trim(Me.ListVoucherNo.SelectedItem))
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
        'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
        'Unload Me
        Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
        Call frmAMISJournalEntry_APJ.SearchVoucherNo(Trim(Me.ListVoucherNo.SelectedItem))
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
        'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
        'Unload Me
        Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
        Call frmAMISJournalEntry_APJ.SearchVoucherNo(Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
        Unload Me
    End If
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
        'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(4)))
        'Unload Me
        Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
        Call frmAMISJournalEntry_APJ.SearchVoucherNo(Trim(Me.ListInvoiceNo.SelectedItem.SubItems(4)))
        Unload Me
    End If
End Sub

Private Sub optCancelled_Click()
    StatusToSearch = "C"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
    If SEARCH_TAB = 3 Then txtSearch_Change
End Sub

Private Sub optCancelled_GotFocus()
    StatusToSearch = "C"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
    If SEARCH_TAB = 3 Then txtSearch_Change
End Sub

Private Sub optPosted_Click()
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
    If SEARCH_TAB = 3 Then txtSearch_Change
End Sub

Private Sub optPosted_GotFocus()
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
    If SEARCH_TAB = 3 Then txtSearch_Change
End Sub

Private Sub optUnPosted_Click()
    StatusToSearch = "N"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
    If SEARCH_TAB = 3 Then txtSearch_Change
End Sub

Private Sub optUnPosted_GotFocus()
    StatusToSearch = "N"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtInvoiceNo_Change
    If SEARCH_TAB = 3 Then txtSearch_Change
End Sub

Private Sub SearchTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    SEARCH_TAB = SearchTab.SelectedItem
    DoEvents
    txtVoucherNo.Enabled = False
    txtVendorPayeeName.Enabled = False
    txtInvoiceNo.Enabled = False
    ListVoucherNo.Enabled = False
    ListVendorPayeeName.Enabled = False
    ListInvoiceNo.Enabled = False
    Select Case Item.Index
    Case 0
        txtVoucherNo.Enabled = True: ListVoucherNo.Enabled = True
        Me.Caption = "Search Item by Voucher Number"
        On Error Resume Next
        txtVoucherNo.SetFocus
    Case 1
        txtVendorPayeeName.Enabled = True: ListVendorPayeeName.Enabled = True
        Me.Caption = "Search Item by All_VENDOR/Payee Name"
        On Error Resume Next
        txtVendorPayeeName.SetFocus
    Case 2
        txtInvoiceNo.Enabled = True: ListInvoiceNo.Enabled = True
        Me.Caption = "Search Item by Invoice No."
        On Error Resume Next
        txtInvoiceNo.SetFocus
    Case 3
        DisplayRR
        Me.Caption = "Search Item by RR No."
        On Error Resume Next
        txtSearch.SetFocus
    End Select
End Sub

Private Sub txtSearch_Change()
    DisplayRR
End Sub
Sub DisplayRR()
' Update By BTT 04/3/2009
    Dim Item                                           As ListItem
    Dim RS                                             As New ADODB.Recordset
    Dim keyword                                        As String
    keyword = Trim(txtSearch.Text)
    If keyword = "" Then
        If chkShowAll4.Value = 1 Then
            Set RS = gconDMIS.Execute("SELECT MRR_NO,INV_NO,AMIS_PV_DETAIL.JDATE,AMIS_PV_DETAIL.VOUCHERNO from Amis_pv_Detail INNER JOIN AMIS_JOURNAL_HD ON AMIS_JOURNAL_HD.VOUCHERNO=AMIS_PV_DETAIL.VOUCHERNO AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE WHERE AMIS_JOURNAL_HD.JTYPE='APJ' AND AMIS_JOURNAL_HD.STATUS = '" & StatusToSearch & "' ORDER BY AMIS_JOURNAL_HD.VOUCHERNO")
        Else
            Set RS = gconDMIS.Execute("SELECT TOP 18 MRR_NO,INV_NO,AMIS_PV_DETAIL.JDATE,AMIS_PV_DETAIL.VOUCHERNO from Amis_pv_Detail INNER JOIN AMIS_JOURNAL_HD ON AMIS_JOURNAL_HD.VOUCHERNO=AMIS_PV_DETAIL.VOUCHERNO AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE WHERE AMIS_JOURNAL_HD.JTYPE='APJ' AND AMIS_JOURNAL_HD.STATUS = '" & StatusToSearch & "' ORDER BY AMIS_JOURNAL_HD.VOUCHERNO")
        End If
    Else
        If chkShowAll4.Value = 1 Then
            Set RS = gconDMIS.Execute("SELECT MRR_NO,INV_NO,AMIS_JOURNAL_HD.JDATE,AMIS_JOURNAL_HD.VOUCHERNO from Amis_pv_Detail INNER JOIN AMIS_JOURNAL_HD ON AMIS_JOURNAL_HD.VOUCHERNO=AMIS_PV_DETAIL.VOUCHERNO AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE where AMIS_JOURNAL_HD.JTYPE ='APJ' AND AMIS_JOURNAL_HD.STATUS = '" & StatusToSearch & "' AND mrr_no like '" & keyword & "%' ORDER BY AMIS_JOURNAL_HD.VOUCHERNO")
        Else
            Set RS = gconDMIS.Execute("SELECT TOP 18 MRR_NO,INV_NO,AMIS_JOURNAL_HD.JDATE,AMIS_JOURNAL_HD.VOUCHERNO from Amis_pv_Detail INNER JOIN AMIS_JOURNAL_HD ON AMIS_JOURNAL_HD.VOUCHERNO=AMIS_PV_DETAIL.VOUCHERNO AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE where AMIS_JOURNAL_HD.JTYPE ='APJ' AND AMIS_JOURNAL_HD.STATUS = '" & StatusToSearch & "' AND mrr_no like '" & Replace(keyword, "'", "") & "%' AND AMIS_JOURNAL_HD.STATUS = '" & StatusToSearch & "' ORDER BY AMIS_JOURNAL_HD.VOUCHERNO")
        End If
    End If
    ListRR.ListItems.Clear
    If Not (RS.EOF And RS.BOF) Then
        Do While Not RS.EOF
            Set Item = ListRR.ListItems.Add(, , Null2String(RS!MRR_No))
            Item.SubItems(1) = Null2String(RS!INV_NO)
            Item.SubItems(2) = Null2String(RS!JDATE)
            Item.SubItems(3) = Null2String(RS!VOUCHERNO)
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtSearch.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListRR.Enabled = True Then ListRR.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVoucherNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtVoucherNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListVoucherNo.Enabled = True Then ListVoucherNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVoucherNo_Change()
    If txtVoucherNo = "" Then
        ListVoucherNo.Enabled = False
        Me.ListVoucherNo.Sorted = False: Me.ListVoucherNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_VENDOR.nameofvendor,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where Jtype='APJ' and status = '" & StatusToSearch & "' order by VoucherNo asc")
        If chkShowAll.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_VENDOR.nameofvendor, amis_pv_detail.inv_no from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_Journal_Hd.JTYPE=amis_pv_detail.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' order by AMIS_Journal_Hd.VoucherNo asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_VENDOR.nameofvendor, amis_pv_detail.inv_no from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_Journal_Hd.JTYPE=amis_pv_detail.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' order by AMIS_Journal_Hd.VoucherNo asc")
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
        'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_VENDOR.nameofvendor,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code Where Jtype='APJ' and status = '" & StatusToSearch & "' and VoucherNo like '" & Format(Trim(Me.txtVoucherNo), "000000") & "%' order by VoucherNo asc")
        If chkShowAll.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_VENDOR.nameofvendor,amis_pv_detail.inv_no from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_Journal_Hd.JTYPE=amis_pv_detail.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' and AMIS_Journal_Hd.VoucherNo like '" & Format(Trim(Replace(Me.txtVoucherNo, "'", "")), "000000") & "%' order by AMIS_Journal_Hd.VoucherNo asc")
            'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_ENTITY.ACCOUNTNAME,amis_pv_detail.inv_no from AMIS_Journal_HD inner join All_ENTITY on AMIS_Journal_Hd.vendorcode = All_ENTITY.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_Journal_Hd.JTYPE=amis_pv_detail.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' and AMIS_Journal_Hd.VoucherNo like '" & Format(Trim(Replace(Me.txtVoucherNo, "'", "")), "000000") & "%' order by AMIS_Journal_Hd.VoucherNo asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_VENDOR.nameofvendor,amis_pv_detail.inv_no from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_Journal_Hd.JTYPE=amis_pv_detail.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' and AMIS_Journal_Hd.VoucherNo like '" & Format(Trim(Replace(Me.txtVoucherNo, "'", "")), "000000") & "%' order by AMIS_Journal_Hd.VoucherNo asc")
            'Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_ENTITY.ACCOUNTNAME,amis_pv_detail.inv_no from AMIS_Journal_HD inner join All_ENTITY on AMIS_Journal_Hd.vendorcode = All_ENTITY.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_Journal_Hd.JTYPE=amis_pv_detail.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' and AMIS_Journal_Hd.VoucherNo like '" & Format(Trim(Replace(Me.txtVoucherNo, "'", "")), "000000") & "%' order by AMIS_Journal_Hd.VoucherNo asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
            ListVoucherNo.Enabled = True
        Else
            ListVoucherNo.Enabled = False
        End If
        ListVoucherNo.Enabled = True
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
    If txtVendorPayeeName = "" Then
        ListVendorPayeeName.Enabled = False
        Me.ListVendorPayeeName.Sorted = False: Me.ListVendorPayeeName.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        'Set rsJournal_HD = gconDMIS.Execute("select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where Jtype='APJ' and status = '" & StatusToSearch & "' order by All_VENDOR.nameofvendor asc")
        If chkShowAll2.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.VoucherNo,amis_pv_detail.inv_no from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' order by All_VENDOR.nameofvendor asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18  All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.VoucherNo,amis_pv_detail.inv_no from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' order by All_VENDOR.nameofvendor asc")
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
        'Set rsJournal_HD = gconDMIS.Execute("select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where Jtype='APJ' and status = '" & StatusToSearch & "' and All_VENDOR.nameofvendor like '" & Trim(Me.txtVendorPayeeName) & "%' order by All_VENDOR.nameofvendor asc")
        If chkShowAll2.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT ALL_VENDOR.NAMEOFVENDOR,AMIS_JOURNAL_HD.JDATE,AMIS_JOURNAL_HD.VOUCHERNO,AMIS_PV_DETAIL.INV_NO FROM AMIS_JOURNAL_HD INNER JOIN ALL_VENDOR ON AMIS_JOURNAL_HD.VENDORCODE = ALL_VENDOR.CODE LEFT OUTER JOIN AMIS_PV_DETAIL ON AMIS_JOURNAL_HD.VOUCHERNO= AMIS_PV_DETAIL.VOUCHERNO AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE WHERE AMIS_JOURNAL_HD.JTYPE='APJ' AND AMIS_JOURNAL_HD.STATUS = '" & StatusToSearch & "' and All_VENDOR.nameofvendor like '" & Trim(Me.txtVendorPayeeName) & "%' order by All_VENDOR.nameofvendor asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT TOP 18 ALL_VENDOR.NAMEOFVENDOR,AMIS_JOURNAL_HD.JDATE,AMIS_JOURNAL_HD.VOUCHERNO,AMIS_PV_DETAIL.INV_NO FROM AMIS_JOURNAL_HD INNER JOIN ALL_VENDOR ON AMIS_JOURNAL_HD.VENDORCODE = ALL_VENDOR.CODE LEFT OUTER JOIN AMIS_PV_DETAIL ON AMIS_JOURNAL_HD.VOUCHERNO= AMIS_PV_DETAIL.VOUCHERNO AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE WHERE AMIS_JOURNAL_HD.JTYPE='APJ' AND AMIS_JOURNAL_HD.STATUS = '" & StatusToSearch & "' and All_VENDOR.nameofvendor like '" & Replace(Trim(Me.txtVendorPayeeName), "'", "") & "%' order by All_VENDOR.nameofvendor asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVendorPayeeName.ListItems, rsJournal_HD
            ListVendorPayeeName.Enabled = True
        Else
            ListVendorPayeeName.Enabled = False
        End If
        ListVendorPayeeName.Enabled = True
    End If
End Sub

Private Sub txtInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtInvoiceNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListInvoiceNo.Enabled = True Then ListInvoiceNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtInvoiceNo_Change()
    If txtInvoiceNo = "" Then
        ListInvoiceNo.Enabled = False
        Me.ListInvoiceNo.Sorted = False: Me.ListInvoiceNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_VENDOR.nameofvendor,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where Jtype='APJ' and status = '" & StatusToSearch & "' order by VoucherNo asc")
        If chkShowAll3.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select amis_pv_detail.inv_no,amis_pv_detail.PO_NO,All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.VoucherNo from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "'")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 amis_pv_detail.inv_no,amis_pv_detail.PO_NO,All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.VoucherNo from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "'")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListInvoiceNo.ListItems, rsJournal_HD
            ListInvoiceNo.Enabled = True
        Else
            ListInvoiceNo.Enabled = False
        End If
    Else
        Me.ListInvoiceNo.Sorted = False: Me.ListInvoiceNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        'Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDate,All_VENDOR.nameofvendor,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code Where Jtype='APJ' and status = '" & StatusToSearch & "' and VoucherNo like '" & Format(Trim(Me.txtVoucherNo), "000000") & "%' order by VoucherNo asc")
        If chkShowAll3.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("select amis_pv_detail.inv_no,amis_pv_detail.PO_NO,All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.VoucherNo from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' and AMIS_PV_Detail.INV_No like '" & Trim(Me.txtInvoiceNo) & "%' order by AMIS_PV_Detail.INV_No asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 amis_pv_detail.inv_no,amis_pv_detail.PO_NO,All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.VoucherNo from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code left outer join amis_pv_detail on AMIS_Journal_Hd.VoucherNo= amis_pv_detail.voucherno AND AMIS_JOURNAL_HD.JTYPE=AMIS_PV_DETAIL.JTYPE where AMIS_Journal_Hd.Jtype='APJ' and AMIS_Journal_Hd.status = '" & StatusToSearch & "' and AMIS_PV_Detail.INV_No like '" & Replace(Trim(Me.txtInvoiceNo), "'", "") & "%' order by AMIS_PV_Detail.INV_No asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListInvoiceNo.ListItems, rsJournal_HD
            ListInvoiceNo.Enabled = True
        Else
            ListInvoiceNo.Enabled = False
        End If
        ListInvoiceNo.Enabled = True
    End If
End Sub


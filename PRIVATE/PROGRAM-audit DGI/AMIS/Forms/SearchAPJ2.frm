VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
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
      ItemCount       =   3
      Item(0).Caption =   "By &Voucher No"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By Vendor/&Payee Name"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "By Invoice No."
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
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
         Begin VB.CheckBox chkShowAll2 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7560
            TabIndex        =   9
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
            Left            =   30
            TabIndex        =   11
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
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VENDOR NAME"
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
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "ACCT CODE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "DUE DATE"
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
         Left            =   30
         TabIndex        =   4
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
            TabIndex        =   8
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
            TabIndex        =   6
            Top             =   45
            Width           =   6240
         End
         Begin MSComctlLib.ListView ListVoucherNo 
            Height          =   4815
            Left            =   30
            TabIndex        =   10
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
            MouseIcon       =   "SearchAPJ2.frx":016E
            NumItems        =   14
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
               Text            =   "VENDOR NAME"
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
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "ACCT CODE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "DUE DATE"
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
            TabIndex        =   5
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
            TabIndex        =   14
            Top             =   45
            Width           =   6240
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7560
            TabIndex        =   13
            Top             =   120
            Width           =   1035
         End
         Begin MSComctlLib.ListView ListInvoiceNo 
            Height          =   4815
            Left            =   30
            TabIndex        =   15
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
            MouseIcon       =   "SearchAPJ2.frx":02D0
            NumItems        =   14
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
               Text            =   "VENDOR NAME"
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
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "ACCT CODE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "DUE DATE"
               Object.Width           =   2540
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
            TabIndex        =   16
            Top             =   90
            Width           =   1125
         End
      End
   End
   Begin MSForms.CheckBox chkCurrentVend 
      Height          =   315
      Left            =   0
      TabIndex        =   7
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
Dim rsJournal_HD                                            As New ADODB.Recordset
Dim Y, k                                                    As Long
Attribute k.VB_VarUserMemId = 1073938433
Dim xJOURNALTYPE                                            As String
Dim CURRENT_VENDORCODE                                      As String

Sub CURRENT_VENDOR(VENDOR_CODE As String)
    CURRENT_VENDORCODE = VENDOR_CODE
End Sub

Sub LOADJOURNAL(XXX As String)
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
    'UPDATE BY: ACL 6232011 ADDTIONAL SEARCH FIELD INVOICENO
    If xJOURNALTYPE = "CDJ" Then
        frmAMISJournalEntry_CDJ.txtPO_No.Text = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(1)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Text = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(2)))
        frmAMISJournalEntry_CDJ.lblInvoiceType.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(5)))
        frmAMISJournalEntry_CDJ.lblinvoiceno.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(6)))
        frmAMISJournalEntry_CDJ.lblInvoiceDate.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(7)))
        frmAMISJournalEntry_CDJ.txtPVAmount.Text = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(8)))
        frmAMISJournalEntry_CDJ.lblPVAmount.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(8)))
        frmAMISJournalEntry_CDJ.lblJ_CLASS.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(12)))
        frmAMISJournalEntry_CDJ.lblCode.Caption = (Trim(Me.ListInvoiceNo.SelectedItem.SubItems(10)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtProd_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtINV_No.Enabled = False
    End If
    Unload Me
End Sub

Private Sub ListInvoiceNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If xJOURNALTYPE = "CDJ" Then
            ListInvoiceNo_DblClick
        End If
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

Private Sub ListVendorPayeeName_DblClick()
    On Error Resume Next
    If xJOURNALTYPE = "CDJ" Then
        frmAMISJournalEntry_CDJ.txtPO_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(1)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
        frmAMISJournalEntry_CDJ.lblInvoiceType.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(5)))
        frmAMISJournalEntry_CDJ.lblinvoiceno.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(6)))
        frmAMISJournalEntry_CDJ.lblInvoiceDate.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(7)))
        frmAMISJournalEntry_CDJ.txtPVAmount.Text = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(9)))
        frmAMISJournalEntry_CDJ.lblPVAmount.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(9)))
        frmAMISJournalEntry_CDJ.lblJ_CLASS.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(12)))
        frmAMISJournalEntry_CDJ.lblCode.Caption = (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(10)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtProd_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtINV_No.Enabled = False
    End If
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
        frmAMISJournalEntry_CDJ.txtPO_No.Text = (Trim(Me.ListVoucherNo.SelectedItem))
        frmAMISJournalEntry_CDJ.txtMRR_No.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(1)))
        frmAMISJournalEntry_CDJ.lblInvoiceType.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(5)))
        frmAMISJournalEntry_CDJ.lblinvoiceno.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(6)))
        frmAMISJournalEntry_CDJ.lblInvoiceDate.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(7)))
        frmAMISJournalEntry_CDJ.txtPVAmount.Text = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(9)))
        frmAMISJournalEntry_CDJ.lblPVAmount.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(9)))
        frmAMISJournalEntry_CDJ.lblJ_CLASS.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(12)))
        frmAMISJournalEntry_CDJ.lblCode.Caption = (Trim(Me.ListVoucherNo.SelectedItem.SubItems(10)))
        frmAMISJournalEntry_CDJ.txtMRR_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtProd_No.Enabled = False
        frmAMISJournalEntry_CDJ.txtINV_No.Enabled = False
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
        If xJOURNALTYPE = "CDJ" Then
            ListVoucherNo_DblClick
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
        If xJOURNALTYPE = "CDJ" Then
            ListVendorPayeeName_DblClick
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
        Me.Caption = "Search Item by Vendor/Payee Name"
        On Error Resume Next
        txtVendorPayeeName.SetFocus
        txtVendorPayeeName_Change
    Case 2
        txtInvoiceNo.Enabled = True: ListInvoiceNo.Enabled = True
        Me.Caption = "Search Item by Invoice No."
        On Error Resume Next
        txtInvoiceNo.SetFocus
        txtInvoiceNo_Change
    End Select
End Sub

Private Sub txtInvoiceNo_Change()
'UPDATED BY: ACL 9202010
'    Dim xOPTION1 As String
'    Dim xOPTION2 As String
'    Dim CMD As New ADODB.Command
'    xOPTION1 = "VOUCHER"
    If txtInvoiceNo = "" Then
        Me.ListInvoiceNo.Sorted = False: Me.ListInvoiceNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkCurrentVend.Value = True Then
            SQL_STATEMENT = "SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P') T " & _
                            "WHERE BALANCE <> 0 AND VENDORCODE = '" & CURRENT_VENDORCODE & "' " & _
                            "ORDER BY JTYPE,VOUCHERNO ASC"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
        Else
            SQL_STATEMENT = "SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            " FROM (" & _
                            " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P')T " & _
                            "WHERE BALANCE <> 0 " & _
                            "ORDER BY JTYPE,VOUCHERNO ASC"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
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
        If chkCurrentVend.Value = True Then
            SQL_STATEMENT = "SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P')T " & _
                            "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND BALANCE <> 0 AND INVOICENO LIKE '" & Trim(Repleys(Me.txtInvoiceNo)) & "%' " & _
                            "ORDER BY JTYPE,VOUCHERNO ASC"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
        Else
            SQL_STATEMENT = "SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P')T " & _
                            "WHERE BALANCE <> 0 AND INVOICENO LIKE '" & Trim(Repleys(Me.txtInvoiceNo)) & "%' " & _
                            "ORDER BY VOUCHERNO,JTYPE ASC"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListInvoiceNo.ListItems, rsJournal_HD
            ListInvoiceNo.Enabled = True
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
        If ListInvoiceNo.Enabled = True Then ListInvoiceNo.SetFocus
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
            SQL_STATEMENT = "SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P') T " & _
                            "WHERE BALANCE <> 0 AND VENDORCODE = '" & CURRENT_VENDORCODE & "' " & _
                            "ORDER BY JTYPE,VOUCHERNO ASC"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
        Else
            SQL_STATEMENT = "SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            " FROM (" & _
                            " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P')T " & _
                            "WHERE BALANCE <> 0 " & _
                            "ORDER BY JTYPE,VOUCHERNO ASC"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
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
            SQL_STATEMENT = "SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P')T " & _
                            "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND BALANCE <> 0 AND VOUCHERNO LIKE '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' " & _
                            "ORDER BY JTYPE,VOUCHERNO ASC"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
        Else
            SQL_STATEMENT = "SELECT JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P')T " & _
                            "WHERE BALANCE <> 0 AND VOUCHERNO LIKE '" & Format(Trim(Repleys(Me.txtVoucherNo)), "000000") & "%' " & _
                            "ORDER BY VOUCHERNO,JTYPE ASC"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
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
            SQL_STATEMENT = "SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P') T " & _
                            "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND BALANCE <> 0 " & _
                            "ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
        Else
            SQL_STATEMENT = "SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P') T " & _
                            "WHERE BALANCE <> 0 " & _
                            "ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
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
            SQL_STATEMENT = "SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P') T " & _
                            "WHERE VENDORCODE = '" & CURRENT_VENDORCODE & "' AND BALANCE <> 0 AND ACCOUNTNAME like '" & Format(Trim(Repleys(Me.txtVendorPayeeName)), "000000") & "%' " & _
                            "ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
        Else
            SQL_STATEMENT = "SELECT ACCOUNTNAME,JTYPE,VOUCHERNO,JDATE,INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE,ACCT_CODE,DUEDATE " & _
                            "FROM (" & _
                            "SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE, " & _
                            "AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE,ACCT_CODE,DUEDATE " & _
                            "FROM AMIS_AP AP WHERE STATUS='P') T " & _
                            "WHERE BALANCE <> 0 AND ACCOUNTNAME like '" & Format(Trim(Repleys(Me.txtVendorPayeeName)), "000000") & "%' " & _
                            "ORDER BY ACCOUNTNAME,JTYPE,VOUCHERNO"
            Set rsJournal_HD = gconDMIS.Execute(SQL_STATEMENT)
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
    Dim RS                                                  As New ADODB.Recordset
    Set RS = gconDMIS.Execute("Select nameofvendor,code from ALL_vendor where code='" & XXX & "'")
    If Not (RS.EOF And RS.BOF) Then
        ReturnVendor = Null2String(RS!nameofvendor)
    Else
        ReturnVendor = ""
    End If
    Set RS = Nothing
End Function


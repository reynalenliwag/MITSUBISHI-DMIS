VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmAMISSearchCDJ 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Cash Disbursement Journals"
   ClientHeight    =   6450
   ClientLeft      =   2970
   ClientTop       =   3495
   ClientWidth     =   8745
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
   ForeColor       =   &H00F5F5F5&
   Icon            =   "SearchCDJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8745
   Begin VB.OptionButton optCancelled 
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      Top             =   420
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
      SelectedItem    =   2
      Item(0).Caption =   "By &Voucher No"
      Item(0).Tooltip =   "Search Cash Disbursement Journal by Voucher Number"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By Vendor/&Payee Name"
      Item(1).Tooltip =   "Search Cash Disbursement Journal by Vendor/Payee Name"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "By &Bank Name"
      Item(2).Tooltip =   "Search Cash Disbursement Journal by Bank Name"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   5325
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.TextBox txtBankName 
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
            Width           =   7365
         End
         Begin MSComctlLib.ListView ListBankName 
            Height          =   4815
            Left            =   0
            TabIndex        =   7
            Top             =   495
            Width           =   8610
            _ExtentX        =   15187
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
            MouseIcon       =   "SearchCDJ.frx":000C
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "BANK NAME"
               Object.Width           =   3528
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
               Text            =   "VOUCHER NO."
               Object.Width           =   3528
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
            TabIndex        =   5
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
            TabIndex        =   10
            Top             =   45
            Width           =   7395
         End
         Begin MSComctlLib.ListView ListVendorPayeeName 
            Height          =   4815
            Left            =   45
            TabIndex        =   11
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
            MouseIcon       =   "SearchCDJ.frx":016E
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VENDOR NAME"
               Object.Width           =   12347
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
            TabIndex        =   9
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
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
            Width           =   7350
         End
         Begin MSComctlLib.ListView ListVoucherNo 
            Height          =   4845
            Left            =   45
            TabIndex        =   15
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
            MouseIcon       =   "SearchCDJ.frx":02D0
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
               Text            =   "VENDOR NAME"
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
Attribute VB_Name = "frmAMISSearchCDJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                                      As New ADODB.Recordset
Dim y, k                                                              As Long
Attribute k.VB_VarUserMemId = 1073938433
Dim StatusToSearch                                                    As String

Sub clearListView()
    For y = 1 To Me.ListVoucherNo.ListItems.Count
        If Me.ListVoucherNo.ListItems.Count <= 0 Then Exit For
        Me.ListVoucherNo.Sorted = False
        Me.ListVoucherNo.ListItems.Remove Me.ListVoucherNo.SelectedItem.Index
    Next y
    For y = 1 To Me.ListVendorPayeeName.ListItems.Count
        If Me.ListVendorPayeeName.ListItems.Count <= 0 Then Exit For
        Me.ListVendorPayeeName.Sorted = False
        Me.ListVendorPayeeName.ListItems.Remove Me.ListVendorPayeeName.SelectedItem.Index
    Next y
    For y = 1 To Me.ListBankName.ListItems.Count
        If Me.ListBankName.ListItems.Count <= 0 Then Exit For
        Me.ListBankName.Sorted = False
        Me.ListBankName.ListItems.Remove Me.ListBankName.SelectedItem.Index
    Next y
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
                txtBankName.SetFocus
        End Select
    End If
    '    If Shift = 2 Then
    '        Select Case KeyCode
    '            Case vbKeyV: SearchTab.SelectedItem = 0
    '            Case vbKeyP: SearchTab.SelectedItem = 1
    '            Case vbKeyB: SearchTab.SelectedItem = 2
    '        End Select
    '        SEARCH_TAB = SearchTab.SelectedItem: SearchTab_SelectedChanged (SearchTab.Selected)
    '    End If
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SearchTab.SelectedItem = SEARCH_TAB
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtBankName_Change
End Sub

Private Sub ListBankName_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListBankName
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub ListBankName_DblClick()
    frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListBankName.SelectedItem.SubItems(3)))
    Unload Me
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
    frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
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
    frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
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
        frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
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
        frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVendorPayeeName.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub ListBankName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtBankName.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListBankName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListBankName.SelectedItem.SubItems(3)))
        Unload Me
    End If
End Sub

Private Sub optCancelled_Click()
    StatusToSearch = "C"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtBankName_Change
End Sub

Private Sub optCancelled_GotFocus()
    StatusToSearch = "C"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtBankName_Change
End Sub

Private Sub optPosted_Click()
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtBankName_Change
End Sub

Private Sub optPosted_GotFocus()
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtBankName_Change
End Sub

Private Sub optUnPosted_Click()
    StatusToSearch = "N"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtBankName_Change
End Sub

Private Sub optUnPosted_GotFocus()
    StatusToSearch = "N"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtVendorPayeeName_Change
    If SEARCH_TAB = 2 Then txtBankName_Change
End Sub

Private Sub SearchTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    SEARCH_TAB = SearchTab.SelectedItem
    DoEvents
    txtVoucherNo.Enabled = False
    txtVendorPayeeName.Enabled = False
    txtBankName.Enabled = False
    ListVoucherNo.Enabled = False
    ListVendorPayeeName.Enabled = False
    ListBankName.Enabled = False
    Select Case SEARCH_TAB
        Case 0
            txtVoucherNo.Enabled = True: ListVoucherNo.Enabled = True
            Me.Caption = "Search Item by Voucher Number"
            On Error Resume Next
            'txtVoucherNo.SetFocus
        Case 1
            txtVendorPayeeName.Enabled = True: ListVendorPayeeName.Enabled = True
            Me.Caption = "Search Item by All_VENDOR/Payee Name"
            On Error Resume Next
            txtVendorPayeeName.SetFocus
        Case 2
            txtBankName.Enabled = True: ListBankName.Enabled = True
            Me.Caption = "Search Item by Bank Name"
            On Error Resume Next
            txtBankName.SetFocus
    End Select
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
        Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,All_VENDOR.nameofvendor,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where Jtype='CDJ' and status = '" & StatusToSearch & "' order by VoucherNo asc")
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
            ListVoucherNo.Enabled = True
        Else
            ListVoucherNo.Enabled = False
        End If
    Else
        Me.ListVoucherNo.Sorted = False: Me.ListVoucherNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        Set rsJournal_HD = gconDMIS.Execute("select AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,ALL_VENDOR.nameofvendor,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where Jtype='CDJ' and status = '" & StatusToSearch & "' and VoucherNo like '" & Format(Trim(Me.txtVoucherNo), "000000") & "%' order by VoucherNo asc")
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
        If ListVendorPayeeName.ListItems.Count > 0 And ListVendorPayeeName.Enabled = True Then: ListVendorPayeeName.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVendorPayeeName_Change()
    If txtVendorPayeeName = "" Then
        Me.ListVendorPayeeName.Sorted = False: Me.ListVendorPayeeName.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        Set rsJournal_HD = gconDMIS.Execute("select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where Jtype='CDJ' and status = '" & StatusToSearch & "' order by All_VENDOR.nameofvendor asc")
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVendorPayeeName.ListItems, rsJournal_HD
            ListVendorPayeeName.Enabled = True
        Else
            ListVendorPayeeName.Enabled = False
        End If
    Else
        Me.ListVendorPayeeName.Sorted = False: Me.ListVendorPayeeName.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        Set rsJournal_HD = gconDMIS.Execute("select All_VENDOR.nameofvendor,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.BankCode from AMIS_Journal_HD inner join All_VENDOR on AMIS_Journal_Hd.vendorcode = All_VENDOR.code where Jtype='CDJ' and status = '" & StatusToSearch & "' AND All_VENDOR.nameofvendor like '" & Trim(Me.txtVendorPayeeName) & "%' order by All_VENDOR.nameofvendor asc")
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVendorPayeeName.ListItems, rsJournal_HD
            ListVendorPayeeName.Enabled = True
        Else
            ListVendorPayeeName.Enabled = False
        End If
    End If
End Sub

Private Sub txtBankName_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtBankName.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListBankName.ListItems.Count > 0 And ListBankName.Enabled = True Then: ListBankName.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtBankName_Change()
    If txtBankName = "" Then
        ListBankName.Enabled = False
        Me.ListBankName.Sorted = False: Me.ListBankName.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        Set rsJournal_HD = gconDMIS.Execute("select VoucherNo,Jdate,VendorCode,BankCode from AMIS_Journal_HD where Jtype='APJ' and status = '" & StatusToSearch & "' order by BankCode asc")
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListBankName.ListItems, rsJournal_HD
            ListBankName.Enabled = True
        Else
            ListBankName.Enabled = False
        End If
    Else
        Me.ListBankName.Sorted = False: Me.ListBankName.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        Set rsJournal_HD = gconDMIS.Execute("select VoucherNo,Jdate,VendorCode,BankCode from AMIS_Journal_HD where Jtype='APJ' and status = '" & StatusToSearch & "' AND BankCode like '" & Trim(Me.txtBankName) & "%' order by BankCode asc")
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListBankName.ListItems, rsJournal_HD
            ListBankName.Enabled = True
        Else
            ListBankName.Enabled = False
        End If
    End If
End Sub


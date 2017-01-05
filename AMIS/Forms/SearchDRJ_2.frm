VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISSearchDRJ_2_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Deposited Cash Receipts Journals"
   ClientHeight    =   6450
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
   Icon            =   "SearchDRJ_2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8700
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
      Left            =   5160
      TabIndex        =   2
      Top             =   60
      Width           =   2715
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
      Width           =   2715
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
      Width           =   2295
   End
   Begin XtremeSuiteControls.TabControl SearchTab 
      Height          =   5955
      Left            =   30
      TabIndex        =   3
      Top             =   450
      Width           =   9240
      _Version        =   655364
      _ExtentX        =   16298
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
      SelectedItem    =   3
      Item(0).Caption =   "By &Voucher No"
      Item(0).Tooltip =   "Search Cash Receipts Journals by Voucher Number"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By Customer Name"
      Item(1).Tooltip =   "Search Cash Receipts Journals by Customer Name"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "By OR Number"
      Item(2).Tooltip =   "Search Cash Receipts Journals by OR Number"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Item(3).Caption =   "By Deposit Number"
      Item(3).Tooltip =   "Search Cash Receipts Journals by Deposit Number"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage4"
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   5325
         Left            =   -69970
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   9180
         _Version        =   655364
         _ExtentX        =   16192
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll3 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7530
            TabIndex        =   18
            Top             =   90
            Width           =   1035
         End
         Begin VB.TextBox txtOrNumber 
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
         Begin MSComctlLib.ListView ListOrNumber 
            Height          =   4890
            Left            =   0
            TabIndex        =   7
            Top             =   420
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   8625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
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
            MouseIcon       =   "SearchDRJ_2.frx":000C
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "OR NO."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "BANK NAME"
               Object.Width           =   7056
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
         Width           =   9180
         _Version        =   655364
         _ExtentX        =   16192
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
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
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
            MouseIcon       =   "SearchDRJ_2.frx":016E
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CUSTOMER NAME"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "INVOICE NO."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "VOUCHER NO."
               Object.Width           =   2646
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
         Left            =   -69970
         TabIndex        =   12
         Top             =   30
         Visible         =   0   'False
         Width           =   9180
         _Version        =   655364
         _ExtentX        =   16192
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.CheckBox chkShowAll 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7500
            TabIndex        =   16
            Top             =   90
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
            Width           =   6210
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
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
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
            MouseIcon       =   "SearchDRJ_2.frx":02D0
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
               Text            =   "BANK NAME"
               Object.Width           =   7408
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
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   5325
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   9180
         _Version        =   655364
         _ExtentX        =   16192
         _ExtentY        =   9393
         _StockProps     =   0
         Begin VB.TextBox txtDepositNo 
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
            TabIndex        =   21
            Top             =   45
            Width           =   6210
         End
         Begin VB.CheckBox chkShowAll4 
            Caption         =   "&Show All"
            Height          =   255
            Left            =   7500
            TabIndex        =   20
            Top             =   90
            Width           =   1035
         End
         Begin MSComctlLib.ListView ListDepositNo 
            Height          =   4845
            Left            =   45
            TabIndex        =   22
            Top             =   480
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   8546
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
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
            MouseIcon       =   "SearchDRJ_2.frx":0432
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "DEPOSIT NO."
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "J. DATE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "BANK NAME"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "VOUCHER NO."
               Object.Width           =   2646
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
            TabIndex        =   23
            Top             =   90
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frmAMISSearchDRJ_2_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SJR 082614
'UPDATES:
'       NEW QUERY FOR SEARCHING
'       NEW SEARCH FUNCTION AND TABS

Option Explicit
Dim rsJournal_HD                                            As New ADODB.Recordset
Dim Y, k                                                    As Long
Attribute k.VB_VarUserMemId = 1073938433
Dim StatusToSearch                                          As String

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    StatusToSearch = "P"
    SearchTab.SelectedItem = SEARCH_TAB
    
    If SEARCH_TAB = 0 Then
        txtVoucherNo_Change
        On Error Resume Next
        txtVoucherNo.SetFocus
    End If
    
    If SEARCH_TAB = 1 Then
        txtCustomer_Change
        On Error Resume Next
        txtCustomer.SetFocus
    End If
    
    If SEARCH_TAB = 2 Then
        txtOrNumber_Change
        On Error Resume Next
        txtOrNumber.SetFocus
    End If
    
    If SEARCH_TAB = 3 Then
        txtDepositNo_Change
        On Error Resume Next
        txtDepositNo.SetFocus
    End If
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
            txtOrNumber.SetFocus
        Case 3:
            On Error Resume Next
            txtDepositNo.SetFocus
        End Select
    End If
    
    If Shift = 2 Then
        Select Case KeyCode
        Case vbKeyV: SearchTab.SelectedItem = 0: txtVoucherNo_Change
        Case vbKeyP: SearchTab.SelectedItem = 1: txtCustomer_Change
        Case vbKeyO: SearchTab.SelectedItem = 2: txtOrNumber_Change
        Case vbKeyD: SearchTab.SelectedItem = 3: txtDepositNo_Change
        End Select
        
        On Error Resume Next
        SEARCH_TAB = SearchTab.SelectedItem: SearchTab_SelectedChanged (SearchTab.Selected)
    End If
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
    For Y = 1 To Me.ListOrNumber.ListItems.Count
        If Me.ListOrNumber.ListItems.Count <= 0 Then Exit For
        Me.ListOrNumber.Sorted = False
        Me.ListOrNumber.ListItems.Remove Me.ListOrNumber.SelectedItem.Index
    Next Y
    For Y = 1 To Me.ListDepositNo.ListItems.Count
        If Me.ListDepositNo.ListItems.Count <= 0 Then Exit For
        Me.ListDepositNo.Sorted = False
        Me.ListDepositNo.ListItems.Remove Me.ListDepositNo.SelectedItem.Index
    Next Y
End Sub

Private Sub chkShowAll_Click()
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
End Sub

Private Sub chkShowAll2_Click()
    If SEARCH_TAB = 1 Then txtCustomer_Change
End Sub

Private Sub chkShowAll3_Click()
    If SEARCH_TAB = 2 Then txtOrNumber_Change
End Sub

Private Sub chkShowAll4_Click()
    If SEARCH_TAB = 3 Then txtDepositNo_Change
End Sub

Private Sub optCancelled_Click()
    StatusToSearch = "C"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtOrNumber_Change
    If SEARCH_TAB = 3 Then txtDepositNo_Change
End Sub

Private Sub optCancelled_GotFocus()
    StatusToSearch = "C"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtOrNumber_Change
    If SEARCH_TAB = 3 Then txtDepositNo_Change
End Sub

Private Sub optPosted_Click()
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtOrNumber_Change
    If SEARCH_TAB = 3 Then txtDepositNo_Change
End Sub

Private Sub optPosted_GotFocus()
    StatusToSearch = "P"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtOrNumber_Change
    If SEARCH_TAB = 3 Then txtDepositNo_Change
End Sub

Private Sub optUnPosted_Click()
    StatusToSearch = "N"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtOrNumber_Change
    If SEARCH_TAB = 3 Then txtDepositNo_Change
End Sub

Private Sub optUnPosted_GotFocus()
    StatusToSearch = "N"
    If SEARCH_TAB = 0 Then txtVoucherNo_Change
    If SEARCH_TAB = 1 Then txtCustomer_Change
    If SEARCH_TAB = 2 Then txtOrNumber_Change
    If SEARCH_TAB = 3 Then txtDepositNo_Change
End Sub

Private Sub SearchTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    SEARCH_TAB = Item.Index
    DoEvents
    txtVoucherNo.Enabled = False
    txtCustomer.Enabled = False
    ListVoucherNo.Enabled = False
    ListCustomer.Enabled = False
    ListDepositNo.Enabled = False
    Select Case SEARCH_TAB
    Case 0
        txtVoucherNo.Enabled = True: ListVoucherNo.Enabled = True
        Me.Caption = "Search Item by Voucher Number"
        On Error Resume Next
        txtVoucherNo.SetFocus
        txtVoucherNo_Change
    Case 1
        txtCustomer.Enabled = True: ListCustomer.Enabled = True
        Me.Caption = "Search Item by Customer Name"
        On Error Resume Next
        txtCustomer.SetFocus
        txtCustomer_Change
    Case 2
        txtOrNumber.Enabled = True: ListOrNumber.Enabled = True
        Me.Caption = "Search Item by OR Number"
        On Error Resume Next
        txtOrNumber.SetFocus
        txtOrNumber_Change
    Case 3
        txtDepositNo.Enabled = True: ListDepositNo.Enabled = True
        Me.Caption = "Search Item by Deposit No."
        On Error Resume Next
        txtDepositNo.SetFocus
        txtDepositNo_Change
    End Select
End Sub

'VOUCHER TAB
'SJR
Private Sub txtVoucherNo_Change()
    If txtVoucherNo = "" Then
        ListVoucherNo.Enabled = False
        Me.ListVoucherNo.Sorted = False: Me.ListVoucherNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT DISTINCT AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.JDATE, " & _
                                                "(SELECT BANKNAME FROM ALL_BANKDEPOSITS WHERE BANKCODE =AMIS_JOURNAL_HD.BANKCODE)AS BANKNAME,AMIS_JOURNAL_HD.INVOICENO " & _
                                                "FROM AMIS_JOURNAL_DET " & _
                                                "INNER JOIN AMIS_JOURNAL_HD " & _
                                                "ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JTYPE=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.DEBIT=0 " & _
                                                "ORDER BY AMIS_JOURNAL_HD.VOUCHERNO ASC")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT DISTINCT TOP 18 AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.JDATE, " & _
                                                "(SELECT BANKNAME FROM ALL_BANKDEPOSITS WHERE BANKCODE =AMIS_JOURNAL_HD.BANKCODE)AS BANKNAME,AMIS_JOURNAL_HD.INVOICENO " & _
                                                "FROM AMIS_JOURNAL_DET " & _
                                                "INNER JOIN AMIS_JOURNAL_HD " & _
                                                "ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JTYPE=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.DEBIT=0 " & _
                                                "ORDER BY AMIS_JOURNAL_HD.VOUCHERNO ASC")
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
        If chkShowAll.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT DISTINCT AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.JDATE, " & _
                                                "(SELECT BANKNAME FROM ALL_BANKDEPOSITS WHERE BANKCODE =AMIS_JOURNAL_HD.BANKCODE)AS BANKNAME,AMIS_JOURNAL_HD.INVOICENO " & _
                                                "FROM AMIS_JOURNAL_DET " & _
                                                "INNER JOIN AMIS_JOURNAL_HD " & _
                                                "ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JTYPE=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.DEBIT=0 " & _
                                                "AND AMIS_JOURNAL_HD.VOUCHERNO LIKE '" & txtVoucherNo.Text & "%' " & _
                                                "ORDER BY AMIS_JOURNAL_HD.VOUCHERNO ASC")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT DISTINCT TOP 18 AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.JDATE, " & _
                                                "(SELECT BANKNAME FROM ALL_BANKDEPOSITS WHERE BANKCODE =AMIS_JOURNAL_HD.BANKCODE)AS BANKNAME,AMIS_JOURNAL_HD.INVOICENO " & _
                                                "FROM AMIS_JOURNAL_DET " & _
                                                "INNER JOIN AMIS_JOURNAL_HD " & _
                                                "ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JTYPE=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.DEBIT=0 " & _
                                                "AND AMIS_JOURNAL_HD.VOUCHERNO LIKE '" & txtVoucherNo.Text & "%' " & _
                                                "ORDER BY AMIS_JOURNAL_HD.VOUCHERNO ASC")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
            ListVoucherNo.Enabled = True
        Else
            ListVoucherNo.Enabled = False
        End If
    End If
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
    Call frmAMISJournalEntry_DRJ_2.SearchVoucherNo(Trim(Me.ListVoucherNo.SelectedItem))
    Unload Me
End Sub

Private Sub ListVoucherNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtVoucherNo.SetFocus
        SendKeys "{HOME}+{END}"
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

Private Sub ListVoucherNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call frmAMISJournalEntry_DRJ_2.SearchVoucherNo(Trim(Me.ListVoucherNo.SelectedItem))
        Unload Me
    End If
End Sub

'CUSTOMER TAB
'SJR
Private Sub txtCustomer_Change()
    If txtCustomer = "" Then
        Me.ListCustomer.Sorted = False: Me.ListCustomer.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll2.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT DISTINCT " & _
                                                "(SELECT ACCOUNTNAME FROM " & _
                                                "(SELECT BANKCODE AS CODE, BANKNAME AS ACCOUNTNAME FROM CMIS_BANKS UNION " & _
                                                "SELECT CODE AS CODE, ACCOUNTNAME AS ACCOUNTNAME FROM ALL_ENTITY) AS TABLEENTIY " & _
                                                "WHERE TABLEENTIY.CODE=RIGHT(ENTITY,LEN(ENTITY)-1)) AS ACCOUNTNAME, " & _
                                                "AMIS_JOURNAL_DET.JDATE, " & _
                                                "AMIS_JOURNAL_DET.INVOICENO, " & _
                                                "AMIS_JOURNAL_DET.VOUCHERNO, " & _
                                                "AMIS_JOURNAL_DET.Entity AS ENTITY " & _
                                                "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JTYPE=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.DEBIT=0 " & _
                                                "AND AMIS_JOURNAL_HD.VOUCHERNO LIKE '" & txtVoucherNo.Text & "%' " & _
                                                "ORDER BY ACCOUNTNAME ASC")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT DISTINCT TOP 18 " & _
                                                "(SELECT ACCOUNTNAME FROM " & _
                                                "(SELECT BANKCODE AS CODE, BANKNAME AS ACCOUNTNAME FROM CMIS_BANKS UNION " & _
                                                "SELECT CODE AS CODE, ACCOUNTNAME AS ACCOUNTNAME FROM ALL_ENTITY) AS TABLEENTIY " & _
                                                "WHERE TABLEENTIY.CODE=RIGHT(ENTITY,LEN(ENTITY)-1)) AS ACCOUNTNAME, " & _
                                                "AMIS_JOURNAL_DET.JDATE, " & _
                                                "AMIS_JOURNAL_DET.INVOICENO, " & _
                                                "AMIS_JOURNAL_DET.VOUCHERNO, " & _
                                                "AMIS_JOURNAL_DET.Entity AS ENTITY " & _
                                                "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JTYPE=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.DEBIT=0 " & _
                                                "AND AMIS_JOURNAL_HD.VOUCHERNO LIKE '" & txtVoucherNo.Text & "%' " & _
                                                "ORDER BY ACCOUNTNAME ASC")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListCustomer.ListItems, rsJournal_HD
            ListCustomer.Enabled = True
        Else
            ListCustomer.Enabled = False
        End If
    Else
        Me.ListCustomer.Sorted = False: Me.ListCustomer.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll2.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT DISTINCT " & _
                                                "(SELECT ACCOUNTNAME FROM " & _
                                                "(SELECT BANKCODE AS CODE, BANKNAME AS ACCOUNTNAME FROM CMIS_BANKS UNION " & _
                                                "SELECT CODE AS CODE, ACCOUNTNAME AS ACCOUNTNAME FROM ALL_ENTITY) AS TABLEENTIY " & _
                                                "WHERE TABLEENTIY.CODE=RIGHT(ENTITY,LEN(ENTITY)-1)) AS ACCOUNTNAME, " & _
                                                "AMIS_JOURNAL_DET.JDATE, " & _
                                                "AMIS_JOURNAL_DET.INVOICENO, " & _
                                                "AMIS_JOURNAL_DET.VOUCHERNO, " & _
                                                "AMIS_JOURNAL_DET.Entity AS ENTITY " & _
                                                "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JTYPE=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.DEBIT=0 " & _
                                                "AND AMIS_JOURNAL_HD.VOUCHERNO LIKE '" & txtVoucherNo.Text & "%' " & _
                                                "ORDER BY ACCOUNTNAME ASC")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT DISTINCT TOP 18 " & _
                                                "(SELECT ACCOUNTNAME FROM " & _
                                                "(SELECT BANKCODE AS CODE, BANKNAME AS ACCOUNTNAME FROM CMIS_BANKS UNION " & _
                                                "SELECT CODE AS CODE, ACCOUNTNAME AS ACCOUNTNAME FROM ALL_ENTITY) AS TABLEENTIY " & _
                                                "WHERE TABLEENTIY.CODE=RIGHT(ENTITY,LEN(ENTITY)-1)) AS ACCOUNTNAME, " & _
                                                "AMIS_JOURNAL_DET.JDATE, " & _
                                                "AMIS_JOURNAL_DET.INVOICENO, " & _
                                                "AMIS_JOURNAL_DET.VOUCHERNO, " & _
                                                "AMIS_JOURNAL_DET.Entity AS ENTITY " & _
                                                "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JTYPE=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.DEBIT=0 " & _
                                                "AND AMIS_JOURNAL_HD.VOUCHERNO LIKE '" & txtVoucherNo.Text & "%' " & _
                                                "ORDER BY ACCOUNTNAME ASC")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListCustomer.ListItems, rsJournal_HD
            ListCustomer.Enabled = True
        Else
            ListCustomer.Enabled = False
        End If
    End If
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
    Call frmAMISJournalEntry_DRJ_2.SearchVoucherNo(Trim(Me.ListCustomer.SelectedItem.SubItems(3)))
    Unload Me
End Sub

Private Sub ListCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtCustomer.SetFocus: SendKeys "{HOME}+{END}"
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

Private Sub ListCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call frmAMISJournalEntry_DRJ_2.SearchVoucherNo(Trim(Me.ListCustomer.SelectedItem.SubItems(3)))
        Unload Me
    End If
End Sub

'OR NUMBER TAB
'SJR
Private Sub txtOrNumber_Change()
    If txtOrNumber = "" Then
        ListOrNumber.Enabled = False
        Me.ListOrNumber.Sorted = False: Me.ListOrNumber.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll3.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT AMIS_JOURNAL_DET.INVOICENO,AMIS_JOURNAL_DET.JDATE,(select bankname from ALL_BANKDEPOSITS where BankCode =AMIS_JOURNAL_HD.BankCode)as BANKNAME,AMIS_JOURNAL_DET.VOUCHERNO " & _
                                                "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD " & _
                                                "ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JType=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.Debit=0 " & _
                                                "ORDER BY AMIS_JOURNAL_DET.INVOICENO ASC")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT TOP 18 " & _
                                                "AMIS_JOURNAL_DET.INVOICENO,AMIS_JOURNAL_DET.JDATE,(select bankname from ALL_BANKDEPOSITS where BankCode =AMIS_JOURNAL_HD.BankCode)as BANKNAME,AMIS_JOURNAL_DET.VOUCHERNO " & _
                                                "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD " & _
                                                "ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JType=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.Debit=0 " & _
                                                "ORDER BY AMIS_JOURNAL_DET.INVOICENO ASC")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListOrNumber.ListItems, rsJournal_HD
            ListOrNumber.Enabled = True
        Else
            ListOrNumber.Enabled = False
        End If
    Else
        Me.ListOrNumber.Sorted = False: Me.ListOrNumber.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll3.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT AMIS_JOURNAL_DET.INVOICENO,AMIS_JOURNAL_DET.JDATE,(select bankname from ALL_BANKDEPOSITS where BankCode =AMIS_JOURNAL_HD.BankCode)as BANKNAME,AMIS_JOURNAL_DET.VOUCHERNO " & _
                                                "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD " & _
                                                "ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JType=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.Debit=0 " & _
                                                "AND AMIS_JOURNAL_DET.INVOICENO " & _
                                                "LIKE '" & txtOrNumber.Text & "%' " & _
                                                "ORDER BY AMIS_JOURNAL_DET.INVOICENO ASC")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT TOP 18 " & _
                                                "AMIS_JOURNAL_DET.INVOICENO,AMIS_JOURNAL_DET.JDATE,(select bankname from ALL_BANKDEPOSITS where BankCode =AMIS_JOURNAL_HD.BankCode)as BANKNAME,AMIS_JOURNAL_DET.VOUCHERNO " & _
                                                "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD " & _
                                                "ON AMIS_JOURNAL_DET.REFERENCENO=AMIS_JOURNAL_HD.REFNO " & _
                                                "AND AMIS_JOURNAL_DET.JType=AMIS_JOURNAL_HD.JTYPE " & _
                                                "AND AMIS_JOURNAL_DET.VOUCHERNO=AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "WHERE AMIS_JOURNAL_DET.JTYPE='DRJ' " & _
                                                "AND AMIS_JOURNAL_DET.STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_DET.Debit=0 " & _
                                                "AND AMIS_JOURNAL_DET.INVOICENO " & _
                                                "LIKE '" & txtOrNumber.Text & "%' " & _
                                                "ORDER BY AMIS_JOURNAL_DET.INVOICENO ASC ")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListOrNumber.ListItems, rsJournal_HD
            ListOrNumber.Enabled = True
        Else
            ListOrNumber.Enabled = False
        End If
    End If
End Sub

Private Sub ListOrNumber_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListOrNumber
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub ListOrNumber_DblClick()
    Call frmAMISJournalEntry_DRJ_2.SearchVoucherNo(Trim(Me.ListOrNumber.SelectedItem.SubItems(3)))
    Unload Me
End Sub

Private Sub txtOrNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtOrNumber.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListOrNumber.ListItems.Count > 0 And ListOrNumber.Enabled = True Then: ListOrNumber.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub ListOrNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtOrNumber.SetFocus: SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListOrNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call frmAMISJournalEntry_DRJ_2.SearchVoucherNo(Trim(Me.ListOrNumber.SelectedItem.SubItems(3)))
        Unload Me
    End If
End Sub

'DEPOSIT NO TAB
'SJR
Private Sub txtDepositNo_Change()
    If txtDepositNo = "" Then
        ListDepositNo.Enabled = False
        Me.ListDepositNo.Sorted = False: Me.ListDepositNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll4.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT AMIS_JOURNAL_HD.REFERENCENO , AMIS_JOURNAL_HD.JDATE, ALL_BANKDEPOSITS.BANKNAME,AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "FROM AMIS_JOURNAL_HD INNER JOIN ALL_BANKDEPOSITS " & _
                                                "ON AMIS_JOURNAL_HD.BANKCODE = ALL_BANKDEPOSITS.BANKCODE " & _
                                                "WHERE AMIS_JOURNAL_HD.JTYPE='DRJ' " & _
                                                "AND STATUS = '" & StatusToSearch & "' " & _
                                                "ORDER BY ALL_BANKDEPOSITS.BANKNAME ASC ")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT TOP 18 " & _
                                                "AMIS_JOURNAL_HD.REFERENCENO , AMIS_JOURNAL_HD.JDATE, ALL_BANKDEPOSITS.BANKNAME,AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "FROM AMIS_JOURNAL_HD INNER JOIN ALL_BANKDEPOSITS " & _
                                                "ON AMIS_JOURNAL_HD.BANKCODE = ALL_BANKDEPOSITS.BANKCODE " & _
                                                "WHERE AMIS_JOURNAL_HD.JTYPE='DRJ' " & _
                                                "AND STATUS = '" & StatusToSearch & "' " & _
                                                "ORDER BY ALL_BANKDEPOSITS.BANKNAME ASC ")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListDepositNo.ListItems, rsJournal_HD
            ListDepositNo.Enabled = True
        Else
            ListDepositNo.Enabled = False
        End If
    Else
        Me.ListDepositNo.Sorted = False: Me.ListDepositNo.ListItems.Clear
        Set rsJournal_HD = New ADODB.Recordset
        If chkShowAll4.Value = 1 Then
            Set rsJournal_HD = gconDMIS.Execute("SELECT AMIS_JOURNAL_HD.REFERENCENO , AMIS_JOURNAL_HD.JDATE, ALL_BANKDEPOSITS.BANKNAME,AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "FROM AMIS_JOURNAL_HD INNER JOIN ALL_BANKDEPOSITS " & _
                                                "ON AMIS_JOURNAL_HD.BANKCODE = ALL_BANKDEPOSITS.BANKCODE " & _
                                                "WHERE AMIS_JOURNAL_HD.JTYPE='DRJ' AND STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_HD.REFERENCENO " & _
                                                "like '" & txtDepositNo & "%' " & _
                                                "ORDER BY ALL_BANKDEPOSITS.BANKNAME ASC ")
        Else
            Set rsJournal_HD = gconDMIS.Execute("SELECT TOP 18 " & _
                                                "AMIS_JOURNAL_HD.REFERENCENO , AMIS_JOURNAL_HD.JDATE, ALL_BANKDEPOSITS.BANKNAME,AMIS_JOURNAL_HD.VOUCHERNO " & _
                                                "FROM AMIS_JOURNAL_HD INNER JOIN ALL_BANKDEPOSITS " & _
                                                "ON AMIS_JOURNAL_HD.BANKCODE = ALL_BANKDEPOSITS.BANKCODE " & _
                                                "WHERE AMIS_JOURNAL_HD.JTYPE='DRJ' AND STATUS = '" & StatusToSearch & "' " & _
                                                "AND AMIS_JOURNAL_HD.REFERENCENO " & _
                                                "like '" & txtDepositNo & "%' " & _
                                                "ORDER BY ALL_BANKDEPOSITS.BANKNAME ASC ")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            Listview_Loadval Me.ListDepositNo.ListItems, rsJournal_HD
            ListDepositNo.Enabled = True
        Else
            ListDepositNo.Enabled = False
        End If
    End If
End Sub

Private Sub ListDepositNo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListDepositNo
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub ListDepositNo_DblClick()
    Call frmAMISJournalEntry_DRJ_2.SearchVoucherNo(Trim(Me.ListDepositNo.SelectedItem.SubItems(3)))
    Unload Me
End Sub

Private Sub txtDepositNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtDepositNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListDepositNo.ListItems.Count > 0 And ListOrNumber.Enabled = True Then: ListDepositNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtDepositNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtDepositNo.SetFocus: SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub txtDepositNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call frmAMISJournalEntry_DRJ_2.SearchVoucherNo(Trim(Me.ListDepositNo.SelectedItem.SubItems(3)))
        Unload Me
    End If
End Sub




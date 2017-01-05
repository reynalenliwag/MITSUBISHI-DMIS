VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISSearchGJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search General Journals"
   ClientHeight    =   5970
   ClientLeft      =   2970
   ClientTop       =   3495
   ClientWidth     =   8415
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
   ForeColor       =   &H00FFC0C0&
   Icon            =   "SearchGJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8415
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
      Left            =   1290
      TabIndex        =   5
      Top             =   450
      Width           =   7050
   End
   Begin MSComctlLib.ListView ListVoucherNo 
      Height          =   4995
      Left            =   30
      TabIndex        =   3
      Top             =   900
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   8811
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
      MouseIcon       =   "SearchGJ.frx":000C
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
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "STATUS"
         Object.Width           =   1058
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
      Left            =   90
      TabIndex        =   4
      Top             =   510
      Width           =   1125
   End
End
Attribute VB_Name = "frmAMISSearchGJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As New ADODB.Recordset
Dim Y, k                                          As Long
Attribute k.VB_VarUserMemId = 1073938433
Dim StatusToSearch                                As String
'UPDATED BY: JUN
Dim xJOURNALTYPE                                  As String
Sub LoadJournalType(XXX As String)
    xJOURNALTYPE = XXX
End Sub

Sub clearListView()
    For Y = 1 To Me.ListVoucherNo.ListItems.Count
        If Me.ListVoucherNo.ListItems.Count <= 0 Then Exit For
        Me.ListVoucherNo.Sorted = False
        Me.ListVoucherNo.ListItems.Remove Me.ListVoucherNo.SelectedItem.Index
    Next Y
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then txtVoucherNo.SetFocus
    If Shift = 2 Then
        Select Case KeyCode
        Case vbKeyV: txtVoucherNo_Change
        End Select
    End If
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    StatusToSearch = "P"
    txtVoucherNo_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
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
'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
    If xJOURNALTYPE = "GJ" Then
        frmAMISJournalEntry_GJ.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
    Else
        frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
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
        'frmAMISJournalEntry.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
        frmAMISJournalEntry_GJ.SearchVoucherNo (Trim(Me.ListVoucherNo.SelectedItem))
        Unload Me
    End If
End Sub

Private Sub optCancelled_Click()
    StatusToSearch = "C"
    txtVoucherNo_Change
End Sub

Private Sub optCancelled_GotFocus()
    StatusToSearch = "C"
    txtVoucherNo_Change
End Sub

Private Sub optPosted_Click()
    StatusToSearch = "P"
    txtVoucherNo_Change
End Sub

Private Sub optPosted_GotFocus()
    StatusToSearch = "P"
    txtVoucherNo_Change
End Sub

Private Sub optUnPosted_Click()
    StatusToSearch = "N"
    txtVoucherNo_Change
End Sub

Private Sub optUnPosted_GotFocus()
    StatusToSearch = "N"
    txtVoucherNo_Change
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
        If xJOURNALTYPE = "GJ" Then
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.REMARKS,AMIS_Journal_HD.status from AMIS_Journal_HD where Jtype='GJ' and status = '" & StatusToSearch & "' order by VoucherNo asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.REMARKS,AMIS_Journal_HD.status from AMIS_Journal_HD where Jtype='OPB' and status = '" & StatusToSearch & "' order by VoucherNo asc")
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
        If xJOURNALTYPE = "GJ" Then
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.REMARKS,AMIS_Journal_HD.status from AMIS_Journal_HD Where Jtype='GJ' and status = '" & StatusToSearch & "' and VoucherNo like '" & Format(Replace(Trim(Me.txtVoucherNo), "'", ""), "000000") & "%' order by VoucherNo asc")
        Else
            Set rsJournal_HD = gconDMIS.Execute("select TOP 18 AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.JDATE,AMIS_Journal_Hd.REMARKS,AMIS_Journal_HD.status from AMIS_Journal_HD Where Jtype='OPB' and status = '" & StatusToSearch & "' and VoucherNo like '" & Format(Trim(Me.txtVoucherNo), "000000") & "%' order by VoucherNo asc")
        End If
        If Not (rsJournal_HD.EOF And rsJournal_HD.BOF) Then
            ListVoucherNo.Enabled = True
            Listview_Loadval Me.ListVoucherNo.ListItems, rsJournal_HD
        Else
            ListVoucherNo.Enabled = False
        End If
    End If
End Sub


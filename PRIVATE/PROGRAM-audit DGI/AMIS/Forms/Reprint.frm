VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmReprintTransaction 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4065
   ClientLeft      =   15
   ClientTop       =   90
   ClientWidth     =   5580
   ControlBox      =   0   'False
   DrawWidth       =   10
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView listReason 
      Height          =   2835
      Left            =   60
      TabIndex        =   2
      Top             =   240
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   5001
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Reason"
         Object.Width           =   9172
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose an Existing Reason"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   90
      TabIndex        =   7
      Top             =   3150
      Width           =   4005
      Begin VB.OptionButton ChkAMIS 
         Caption         =   "AMIS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   420
         Width           =   735
      End
      Begin VB.OptionButton ChkPMIS 
         Caption         =   "PMIS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   810
         TabIndex        =   11
         Top             =   420
         Width           =   735
      End
      Begin VB.OptionButton ChkCSMS 
         Caption         =   "CSMS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   420
         Width           =   795
      End
      Begin VB.OptionButton ChkCMIS 
         Caption         =   "CMIS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2340
         TabIndex        =   9
         Top             =   420
         Width           =   795
      End
      Begin VB.OptionButton ChkSMIS 
         Caption         =   "SMIS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   8
         Top             =   420
         Width           =   765
      End
   End
   Begin VB.TextBox txtReason 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   60
      TabIndex        =   3
      Top             =   270
      Width           =   5475
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4830
      MouseIcon       =   "Reprint.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Reprint.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   3240
      Width           =   705
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Confirm"
      Height          =   765
      Left            =   4140
      MouseIcon       =   "Reprint.frx":059D
      MousePointer    =   99  'Custom
      Picture         =   "Reprint.frx":06EF
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Save this Record"
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label LblTransactionNo 
      Caption         =   "transaction"
      Height          =   285
      Left            =   6420
      TabIndex        =   6
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label lblTransaction_type 
      Caption         =   "type"
      Height          =   285
      Left            =   6450
      TabIndex        =   5
      Top             =   390
      Width           =   1515
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   90
      TabIndex        =   4
      Top             =   60
      Width           =   5385
   End
End
Attribute VB_Name = "FrmReprintTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheTransaction_No                                       As String
Dim TheApplication_type                                     As String
Dim xModule                                                 As String

Sub displayReason(nard As String)
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset
    Dim Item                                                As ListItem

    SQL = "SELECT * from All_REPRINT_transaction where module_name=" & nard & ""

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listReason.ListItems.Clear
    Do While Not RS.EOF
        Set Item = listReason.ListItems.Add(, , Null2String(RS!application_type))
        Item.SubItems(1) = Null2String(RS!reason)
        RS.MoveNext
    Loop
    lblState.Caption = "Please Select Your Reason!"
    Set RS = Nothing
End Sub

Sub UpdateReprint(XApplication_type As String, Xtransaction_no As String)
    Dim SQL                                                 As String

    'SQL = "UPDATE all_reprint_transaction set Reason='" & txtReason.Text & "',date_reprint =" & N2Date2Null(LOGDATE) & ",who_reprint='" & LOGNAME & "',Reprint='1' where application_type='" & XApplication_type & "' and transaction_no='" & Xtransaction_no & "'"
    SQL = "INSERT INTO ALL_REPRINT_TRANSACTION (APPLICATION_TYPE,MODULE_NAME,TRANSACTION_NO, REASON,DATE_REPRINT,WHO_REPRINT,REPRINT) " & _
          " VALUES('" & XApplication_type & "','" & MODULENAME & "','" & Xtransaction_no & "','" & txtReason.Text & "','" & LOGDATE & "','" & LOGNAME & "'," & 1 & ")"
    gconDMIS.Execute (SQL)

    ShowSuccessFullyUpdated
End Sub

Private Sub ChkAMIS_Click()
    If ChkAMIS.Value = True Then
        xModule = "'AMIS'"
        listReason.Visible = True
        displayReason xModule
        If Not listReason.ListItems.Count = 0 Then
            lblState.Caption = "Choose a Reason"
            listReason.SetFocus
        Else
            listReason.Visible = False
            lblState.Caption = "Enter a Reason"
            txtReason.SetFocus
        End If

    End If
End Sub

Private Sub ChkCMIS_Click()
    If ChkCMIS.Value = True Then
        xModule = "'CMIS'"
        listReason.Visible = True
        displayReason xModule
        If Not listReason.ListItems.Count = 0 Then
            lblState.Caption = "Choose a Reason"
            listReason.SetFocus
        Else
            listReason.Visible = False
            lblState.Caption = "Enter a Reason"
            txtReason.SetFocus
        End If
    End If
End Sub

Private Sub ChkCSMS_Click()
    If chkCSMS.Value = True Then
        xModule = "'CSMS'"
        listReason.Visible = True
        displayReason xModule
        If Not listReason.ListItems.Count = 0 Then
            lblState.Caption = "Choose a Reason"
            listReason.SetFocus
        Else
            listReason.Visible = False
            lblState.Caption = "Enter a Reason"
            txtReason.SetFocus
        End If

    End If
End Sub

Private Sub ChkPMIS_Click()
    If chkPMIS.Value = True Then
        xModule = "'PMIS'"
        listReason.Visible = True
        displayReason xModule
        If Not listReason.ListItems.Count = 0 Then
            lblState.Caption = "Choose a Reason"
            listReason.SetFocus
        Else
            listReason.Visible = False
            lblState.Caption = "Enter a Reason"
            txtReason.SetFocus
        End If

    End If
End Sub

Private Sub ChkSMIS_Click()
    If chkSMIS.Value = True Then
        xModule = "'SMIS'"
        listReason.Visible = True
        displayReason xModule
        If Not listReason.ListItems.Count = 0 Then
            lblState.Caption = "Choose a Reason"
            listReason.SetFocus
        Else
            listReason.Visible = False
            lblState.Caption = "Enter a Reason"
            txtReason.SetFocus
        End If

    End If
End Sub

Private Sub cmdCancel_Click()
    CANCEL_ANS = "NO"
    Unload Me
End Sub

Private Sub cmdSave_Click()
'On Error Resume Next
    TheTransaction_No = LblTransactionNo
    TheApplication_type = lblTransaction_type
    If txtReason.Text = "" Then
        ShowIsRequiredMsg "Reason Cannot be blank!"
        txtReason.SetFocus
        Exit Sub
    End If
    If Len(txtReason.Text) < 5 Then
        ShowIsRequiredMsg "Please State a valid reason!"
        txtReason.SetFocus
        Exit Sub
    End If

    CANCEL_ANS = "YES"
    UpdateReprint TheApplication_type, TheTransaction_No

    Unload Me

End Sub

Private Sub Form_Load()
'Update BY BTT - 06282008
    Dim TitleApp                                            As String
    Me.Caption = ""
    CenterMe frmMain, Me, 1
    listReason.Visible = False
    If MODULENAME = "AMIS" Then
        'For AMIS
        If JOURNALTYPE = "APJ" Then
            TitleApp = "ACCOUNTS PAYABLE JOURNAL"
        End If
        If JOURNALTYPE = "CDJ" Then
            TitleApp = "CASH DISBURESEMENT JOURNAL"
        End If
        If JOURNALTYPE = "SJ" Then
            TitleApp = "SALES JOURNAL"
        End If
        If JOURNALTYPE = "GJ" Then
            TitleApp = "GENERAL JOURNAL"
        End If
        If JOURNALTYPE = "CRJ" Then
            TitleApp = "UN-DEPOSITED CASH RECEIPTS JOURNAL"
        End If
        If JOURNALTYPE = "DRJ" Then
            TitleApp = "DEPOSITED CASH RECEIPTS JOURNAL"
        End If
        If JOURNALTYPE = "COB" Then
            TitleApp = "CUSTOMER OPENNING BALANCE"
        End If
        If JOURNALTYPE = "VPJ" Then
            TitleApp = "VENDOR OPENNING BALANCE"
        End If
        lblState.Caption = "State Reason Of Re-Printing This Journal"
    End If
    If MODULENAME = "PMIS" Then
        'For PMIS
    End If
    If MODULENAME = "CSMS" Then
        'FOR CSMS
    End If
    If MODULENAME = "CMIS" Then
        'For CMIS
    End If
    If MODULENAME = "SMIS" Then
        'For SMIS
    End If
    Me.Caption = Me.Caption & TitleApp & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub listReason_DblClick()
    If listReason.ListItems.Count = 0 Then Exit Sub
    txtReason.Text = listReason.SelectedItem.SubItems(1)
    listReason.Visible = False
    lblState.Caption = "Enter a Reason"
End Sub

Private Sub listReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If listReason.ListItems.Count = 0 Then Exit Sub
        txtReason.Text = listReason.SelectedItem.SubItems(1)
        listReason.Visible = False
        lblState.Visible = False
    End If
End Sub


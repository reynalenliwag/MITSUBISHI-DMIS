VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form FrmCancelTransaction 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3975
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5655
   ControlBox      =   0   'False
   DrawWidth       =   10
   Icon            =   "CancelTransaction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5655
   Begin MSComctlLib.ListView listReason 
      Height          =   2685
      Left            =   60
      TabIndex        =   7
      Top             =   330
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabIndex        =   6
      Top             =   3060
      Width           =   4005
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
         TabIndex        =   12
         Top             =   420
         Width           =   765
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
         TabIndex        =   11
         Top             =   420
         Width           =   795
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
         Left            =   840
         TabIndex        =   9
         Top             =   420
         Width           =   795
      End
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
         TabIndex        =   8
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.TextBox txtReason 
      Height          =   2685
      Left            =   60
      MaxLength       =   1000
      TabIndex        =   0
      Top             =   330
      Width           =   5475
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4860
      MousePointer    =   99  'Custom
      Picture         =   "CancelTransaction.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   3090
      Width           =   675
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Confirm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4170
      MousePointer    =   99  'Custom
      Picture         =   "CancelTransaction.frx":0457
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save this Record"
      Top             =   3090
      Width           =   705
   End
   Begin VB.Label lblTransaction_type 
      Caption         =   "type"
      Height          =   285
      Left            =   6120
      TabIndex        =   4
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label LblTransactionNo 
      Caption         =   "transaction"
      Height          =   285
      Left            =   6090
      TabIndex        =   3
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   5565
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   5655
      _Version        =   655364
      _ExtentX        =   9975
      _ExtentY        =   503
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColorLight=   12632256
      GradientColorDark=   4210752
   End
End
Attribute VB_Name = "FrmCancelTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheTransaction_No                             As String
Dim TheReason                                     As String
Dim TheApplication_type                           As String
Dim xModule                                       As String
Dim xJOURNALTYPE                                  As String

Sub SaveCancelInformation(XApplication_type As String, Xtransaction_no As String, XReason As String)
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    On Error GoTo RYAN:

    'UPDATE BY : MJP 07252008 12;24 PM
    gconDMIS.Execute "DELETE FROM ALL_reprint_transaction WHERE APPLICATION_TYPE = '" & XApplication_type & "' AND TRANSACTION_NO = '" & Xtransaction_no & "' AND REPRINT = " & 0 & ""
    'UPDATE BY : MJP 07252008 12;24 PM

    gconDMIS.Execute "INSERT INTO ALL_cancel_transaction (Application_type,Module_Name,Transaction_no,Reason,date_cancelled,Who_cancelled)Values('" & XApplication_type & _
                     "','" & MODULENAME & "','" & Xtransaction_no & "','" & XReason & "'," & N2Date2Null(LOGDATE) & ",'" & LOGNAME & "' )"

    MessagePop InfoFriend, "Information Updated", "Transaction Sucessfully Cancelled!", 1000

    Set RS = Nothing
    Exit Sub

RYAN:
    MsgBox Err.Description, "Error"
End Sub

Sub displayReason(nard As String)
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    Dim Item                                      As ListItem

    SQL = "SELECT * from All_cancel_transaction where module_name=" & nard & ""

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
    If ChkCSMS.Value = True Then
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
    If ChkPMIS.Value = True Then
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
    If ChkSMIS.Value = True Then
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
    TheTransaction_No = LblTransactionNo
    TheApplication_type = lblTransaction_type
    TheReason = Trim(txtReason.Text)
    If TheReason = "" Then
        MsgBox "Please state a Reason..", vbInformation, "Information"
        Exit Sub
    End If
    ' Saving Routin
    SaveCancelInformation TheApplication_type, TheTransaction_No, TheReason
    cmdSave.Visible = False
    CANCEL_ANS = "YES"
    'Call cmdCancel_Click
    'cmdCancel.Visible = True
    Unload Me
End Sub

Private Sub Form_Load()
'Update BY BTT - 07212008
    Dim TitleApp                                  As String
    'Me.Caption = ""
    CenterMe frmMain, Me, 1
    listReason.Visible = False
    If MODULENAME = "AMIS" Then
        'For AMIS
        If xJOURNALTYPE = "APJ" Then
            TitleApp = "ACCOUNTS PAYABLE JOURNAL"
        End If
        If xJOURNALTYPE = "CDJ" Then
            TitleApp = "CASH DISBURESEMENT JOURNAL"
        End If
        If xJOURNALTYPE = "SJ" Then
            TitleApp = "SALES JOURNAL"
        End If
        If xJOURNALTYPE = "GJ" Then
            TitleApp = "GENERAL JOURNAL"
        End If
        If xJOURNALTYPE = "CRJ" Then
            TitleApp = "UN-DEPOSITED CASH RECEIPTS JOURNAL"
        End If
        If xJOURNALTYPE = "DRJ" Then
            TitleApp = "DEPOSITED CASH RECEIPTS JOURNAL"
        End If
        If xJOURNALTYPE = "COB" Then
            TitleApp = "CUSTOMER OPENNING BALANCE"
        End If
        If xJOURNALTYPE = "VPJ" Then
            TitleApp = "VENDOR OPENNING BALANCE"
        End If
        lblState.Caption = "State Reason Of Cancelation of This Journal"
    End If
    If MODULENAME = "PMIS" Then
        'For PMIS
    End If
    If MODULENAME = "CSMS" Then
        'For CSMS
    End If
    If MODULENAME = "CMIS" Then
        'For CMIS
    End If
    If MODULENAME = "SMIS" Then
        'For SMIS
    End If
    'Me.Caption = Me.Caption & TitleApp & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Sub LoadJournal(XXX As String)
    xJOURNALTYPE = XXX
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

Private Sub listReason_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        listReason.Visible = False
    End If
End Sub


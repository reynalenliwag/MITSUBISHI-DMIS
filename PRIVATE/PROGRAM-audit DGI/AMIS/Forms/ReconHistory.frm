VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmReconHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Reconcillation History"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   ForeColor       =   &H00E0E0E0&
   Icon            =   "ReconHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvHistory 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10292
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Recon Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account No"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bank Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Bank Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Book Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Adjusted"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmReconHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xList                                                   As ListItem
Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim rsReconHistory                                      As ADODB.Recordset
    Set rsReconHistory = New ADODB.Recordset
    rsReconHistory.Open "SELECT HISTORY.BANKID,HISTORY.RECONDATE,BANKS.BANKACCTNO,BANKS.BANKNAME,HISTORY.BANK,HISTORY.BOOK,HISTORY.ADJUSTED FROM AMIS_RECONHISTORY HISTORY INNER JOIN ALL_BANKS BANKS ON HISTORY.BANKID=BANKS.ID", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsReconHistory.EOF And Not rsReconHistory.BOF Then
        Do While Not rsReconHistory.EOF
            Set xList = lvHistory.ListItems.Add(, , rsReconHistory!ReconDate)
            xList.SubItems(1) = rsReconHistory!BankAcctNo
            xList.SubItems(2) = rsReconHistory!BankName
            xList.SubItems(3) = rsReconHistory!Bank
            xList.SubItems(4) = rsReconHistory!Book
            xList.SubItems(5) = rsReconHistory!Adjusted
            rsReconHistory.MoveNext
        Loop
    End If
End Sub


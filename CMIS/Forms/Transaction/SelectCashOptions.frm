VERSION 5.00
Begin VB.Form frmCASHPOSITIONSelectCashOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select..."
   ClientHeight    =   2040
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   3810
   ForeColor       =   &H00F5F5F5&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   0
      ScaleHeight     =   2025
      ScaleWidth      =   3705
      TabIndex        =   4
      Top             =   0
      Width           =   3705
      Begin VB.CommandButton cmdPettyCashFund 
         Caption         =   "Petty Cash Fund Replenishment..."
         Height          =   345
         Left            =   90
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton cmdPayCashAdvance 
         Caption         =   "Payment of Cash Advances..."
         Height          =   345
         Left            =   90
         TabIndex        =   5
         Top             =   1620
         Width           =   3495
      End
      Begin VB.CommandButton cmdCashierCollection 
         Caption         =   "Cashier Collection..."
         Height          =   345
         Left            =   90
         TabIndex        =   0
         ToolTipText     =   "View Cashier Collection"
         Top             =   60
         Width           =   3495
      End
      Begin VB.CommandButton cmdLTOReg 
         Caption         =   "LTO Registration Replenishment..."
         Height          =   345
         Left            =   90
         TabIndex        =   3
         Top             =   1230
         Width           =   3495
      End
      Begin VB.CommandButton cmdCheckEncash 
         Caption         =   "Check Encashment..."
         Height          =   345
         Left            =   90
         TabIndex        =   1
         ToolTipText     =   "View Check Encashment"
         Top             =   450
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONSelectCashOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCashierCollection_Click()
    Unload Me
    CASH_OPTIONS = "CASH_COL"
    frmCASHPOSITIONCashierCollection.Show vbModal
    
    LogAudit "V", "CASH POSITION - CASHIER COLLECTION"
End Sub

Private Sub cmdCheckEncash_Click()
    Unload Me
    CASH_OPTIONS = "CHECK_EN"
    frmCASHPOSITIONCashierCollection.Show vbModal
    
    LogAudit "V", "CASH POSITION - CHECK ENCASHMENT"
End Sub

Private Sub cmdPettyCashFund_Click()
    Unload Me
    CASH_OPTIONS = "PET_REPL"
    frmCASHPOSITIONCashierCollection.Show vbModal
End Sub

Private Sub cmdLTOReg_Click()
    Unload Me
    CASH_OPTIONS = "LTO_REPL"
    frmCASHPOSITIONCashierCollection.Show vbModal
End Sub

Private Sub cmdPayCashAdvance_Click()
    Unload Me
    CASH_OPTIONS = "PET_ADV"
    frmCASHPOSITIONCashierCollection.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]" '"." & App.Revision & "]"
    If TYPE_ON_HAND = "CARD" Then
        cmdCheckEncash.Enabled = False
        cmdPettyCashFund.Enabled = False
        cmdLTOReg.Enabled = False
        cmdPayCashAdvance.Enabled = False
    ElseIf TYPE_ON_HAND = "CHECK" Then
        cmdCheckEncash.Enabled = True
        cmdPettyCashFund.Enabled = False
        cmdLTOReg.Enabled = False
        cmdPayCashAdvance.Enabled = False
    End If
    Screen.MousePointer = 0
End Sub


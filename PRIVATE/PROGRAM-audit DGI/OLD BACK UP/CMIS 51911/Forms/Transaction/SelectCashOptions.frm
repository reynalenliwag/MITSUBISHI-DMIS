VERSION 5.00
Begin VB.Form frmCASHPOSITIONSelectCashOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select..."
   ClientHeight    =   840
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   3810
   ForeColor       =   &H00F5F5F5&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
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
      Begin VB.CommandButton Command14 
         Caption         =   "Payment of Cash Advances..."
         Height          =   345
         Left            =   90
         TabIndex        =   5
         Top             =   1620
         Width           =   3495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Cashier Collection..."
         Height          =   345
         Left            =   90
         TabIndex        =   0
         ToolTipText     =   "View Cashier Collection"
         Top             =   60
         Width           =   3495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Check Encashment..."
         Height          =   345
         Left            =   90
         TabIndex        =   1
         ToolTipText     =   "View Check Encashment"
         Top             =   450
         Width           =   3495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Petty Cash Fund Replenishment..."
         Height          =   345
         Left            =   90
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command13 
         Caption         =   "LTO Registration Replenishment..."
         Height          =   345
         Left            =   90
         TabIndex        =   3
         Top             =   1230
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONSelectCashOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command10_Click()
    Unload Me
    CASH_OPTIONS = "CASH_COL"
    frmCASHPOSITIONCashierCollection.Show vbModal
    LogAudit "V", "CASH POSITION - CASHIER COLLECTION"
End Sub

Private Sub Command11_Click()
    Unload Me
    CASH_OPTIONS = "CHECK_EN"
    frmCASHPOSITIONCashierCollection.Show vbModal
    LogAudit "V", "CASH POSITION - CHECK ENCASHMENT"
End Sub

Private Sub Command12_Click()
    Unload Me
    CASH_OPTIONS = "PET_REPL"
    frmCASHPOSITIONCashierCollection.Show vbModal
End Sub

Private Sub Command13_Click()
    Unload Me
    CASH_OPTIONS = "LTO_REPL"
    frmCASHPOSITIONCashierCollection.Show vbModal
End Sub

Private Sub Command14_Click()
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
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If TYPE_ON_HAND = "CARD" Then
        Command11.Enabled = False
        Command12.Enabled = False
        Command13.Enabled = False
        Command14.Enabled = False
    End If
    If TYPE_ON_HAND = "CHECK" Then
        Command11.Enabled = True
        Command12.Enabled = False
        Command13.Enabled = False
        Command14.Enabled = False
    End If
    Screen.MousePointer = 0
End Sub


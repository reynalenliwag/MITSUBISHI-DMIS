VERSION 5.00
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "WIZBOX.OCX"
Object = "{205EA659-0BC9-4F44-85D9-FBC10C8940C1}#1.0#0"; "WIZDIGIT.OCX"
Begin VB.Form frmCMISCASHPaymentEntry 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Official Receipt Payment Entry Box"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin wizBox.Box Box1 
      Height          =   1695
      Left            =   30
      Top             =   60
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   2990
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment Module"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   150
      TabIndex        =   13
      Top             =   5280
      Width           =   9015
      Begin VB.TextBox txtAmountTendered 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   360
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Tendered :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   14
         Top             =   390
         Width           =   2085
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   150
      ScaleHeight     =   1215
      ScaleWidth      =   8985
      TabIndex        =   9
      Top             =   3900
      Width           =   9015
      Begin wizDigits.wizDigit wizDigit3 
         Height          =   1215
         Left            =   -420
         TabIndex        =   10
         Top             =   0
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   2143
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9EFE3&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   300
      ScaleHeight     =   1245
      ScaleWidth      =   8865
      TabIndex        =   11
      Top             =   3900
      Width           =   8865
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F88A56&
         Height          =   360
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "P 2,030.00 "
         Top             =   90
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   150
      ScaleHeight     =   1215
      ScaleWidth      =   8985
      TabIndex        =   5
      Top             =   2160
      Width           =   9015
      Begin wizDigits.wizDigit wizDigit2 
         Height          =   1215
         Left            =   -420
         TabIndex        =   6
         Top             =   0
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   2143
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9EFE3&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   300
      ScaleHeight     =   1245
      ScaleWidth      =   8865
      TabIndex        =   7
      Top             =   2160
      Width           =   8865
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F88A56&
         Height          =   360
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "P 2,030.00 "
         Top             =   90
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   150
      ScaleHeight     =   1215
      ScaleWidth      =   8985
      TabIndex        =   1
      Top             =   420
      Width           =   9015
      Begin wizDigits.wizDigit wizDigit1 
         Height          =   1215
         Left            =   -420
         TabIndex        =   2
         Top             =   0
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   2143
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9EFE3&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   300
      ScaleHeight     =   1245
      ScaleWidth      =   8865
      TabIndex        =   3
      Top             =   420
      Width           =   8865
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F88A56&
         Height          =   360
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "P 2,030.00 "
         Top             =   90
         Width           =   1395
      End
   End
   Begin wizBox.Box Box2 
      Height          =   1695
      Left            =   30
      Top             =   1800
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   2990
   End
   Begin wizBox.Box Box3 
      Height          =   1695
      Left            =   30
      Top             =   3540
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   2990
   End
   Begin VB.Label labAmount 
      Caption         =   "0.00"
      Height          =   285
      Left            =   180
      TabIndex        =   18
      Top             =   5670
      Width           =   465
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   180
      TabIndex        =   17
      Top             =   3600
      Width           =   2085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Tendered"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   180
      TabIndex        =   16
      Top             =   1860
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipts Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   180
      TabIndex        =   15
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "frmCMISCASHPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CenterMe frmMain, Me, 1
AMOUNT_TENDERED = 0: CHANGE_DUE = 0
wizDigit1.TextValue = ToDoubleNumber(RECEIPTS_AMOUNT)
wizDigit2.TextValue = ToDoubleNumber(AMOUNT_TENDERED)
wizDigit3.TextValue = ToDoubleNumber(CHANGE_DUE)
End Sub

Private Sub txtAmountTendered_Change()
AMOUNT_TENDERED = NumericVal(txtAmountTendered.Text)
If AMOUNT_TENDERED >= RECEIPTS_AMOUNT Then
   CHANGE_DUE = AMOUNT_TENDERED - NumericVal(RECEIPTS_AMOUNT)
Else
   CHANGE_DUE = "0.00"
End If
wizDigit2.TextValue = ToDoubleNumber(AMOUNT_TENDERED)
wizDigit3.TextValue = ToDoubleNumber(CHANGE_DUE)
End Sub

Private Sub txtAmountTendered_GotFocus()
If NumericVal(txtAmountTendered.Text) > 0 Then txtAmountTendered.Text = NumericVal(txtAmountTendered.Text) Else txtAmountTendered.Text = ""
End Sub

Private Sub txtAmountTendered_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   If AMOUNT_TENDERED >= RECEIPTS_AMOUNT Then
   Else
      MsgBox "Payment Amount does not meet OR Amount...", vbInformation + vbOKOnly, "Message"
   End If
End If
End Sub

Private Sub txtAmountTendered_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtAmountTendered_LostFocus()
txtAmountTendered.Text = ToDoubleNumber(txtAmountTendered.Text)
End Sub

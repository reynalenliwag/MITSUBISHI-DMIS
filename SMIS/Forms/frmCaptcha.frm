VERSION 5.00
Object = "{54AC2DF1-B6CB-406E-BB23-DC06DF6AAD9E}#16.0#0"; "wizCrypto.ocx"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Begin VB.Form frmCaptcha 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picSaves 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   2280
      ScaleHeight     =   645
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   1440
      Width           =   1500
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   555
         Left            =   90
         MouseIcon       =   "frmCaptcha.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmCaptcha.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save Log Call"
         Top             =   65
         Width           =   705
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   555
         Left            =   780
         MouseIcon       =   "frmCaptcha.frx":04A2
         MousePointer    =   99  'Custom
         Picture         =   "frmCaptcha.frx":05F4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancel"
         Top             =   65
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   3555
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.Label lblcapcha 
         Alignment       =   2  'Center
         Caption         =   "uiyuiui"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3615
      Begin VB.TextBox txtcapcha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3375
      End
   End
   Begin wizCrypto.Crypto Crypto1 
      Height          =   465
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   820
   End
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   5280
      Top             =   4920
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
End
Attribute VB_Name = "frmCaptcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function RandomString(cb As Integer) As String

    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next
    RandomString = RandomString & Format(LOGCODE, "000")

End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    With wizVar
        If .EncryptAccess(Trim(lblcapcha.Caption)) = .EncryptAccess(Trim(txtcapcha.Text)) Then
            With frmSMIS_Trans_VehicleInvoice
               Call .unreleaseme
               cmdCancel_Click
            End With
        Else
             MessagePop InfoWarning, "Warning", "Invalid Code!"
        End If
    
    End With
End Sub

Private Sub Form_Load()
    Set CryptVar = frmCaptcha.Crypto1
    Set wizVar = Me.wizEnc1
    lblcapcha.Caption = RandomString(17)
End Sub



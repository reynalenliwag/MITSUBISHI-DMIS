VERSION 5.00
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Begin VB.Form frmUpdatePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Setting"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UpdatePassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   390
      Top             =   5280
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6450
      TabIndex        =   2
      Top             =   5190
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Previous"
      Height          =   375
      Left            =   4290
      TabIndex        =   1
      Top             =   5190
      Width           =   1035
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   5370
      TabIndex        =   0
      Top             =   5190
      Width           =   1035
   End
   Begin VB.PictureBox picStep2 
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   7575
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin wizEncrypt.wizEnc wizEnc1 
         Left            =   4410
         Top             =   1080
         _ExtentX        =   3969
         _ExtentY        =   3969
      End
      Begin VB.ComboBox cboUserName 
         Height          =   330
         Left            =   3480
         TabIndex        =   33
         Top             =   1710
         Width           =   3255
      End
      Begin VB.TextBox txtUserConfirmPassword 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   15
         PasswordChar    =   "l"
         TabIndex        =   9
         Top             =   3000
         Width           =   3225
      End
      Begin VB.TextBox txtUserPassword 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   15
         PasswordChar    =   "l"
         TabIndex        =   5
         Top             =   2340
         Width           =   3225
      End
      Begin VB.Image Image3 
         Height          =   4125
         Left            =   0
         Picture         =   "UpdatePassword.frx":08CA
         Top             =   900
         Width           =   2820
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3480
         TabIndex        =   8
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3480
         TabIndex        =   7
         Top             =   2130
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3480
         TabIndex        =   6
         Top             =   1470
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Your User Name and Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   210
         TabIndex        =   4
         Top             =   240
         Width           =   4185
      End
      Begin VB.Image Image2 
         Height          =   885
         Left            =   0
         Picture         =   "UpdatePassword.frx":2A43
         Top             =   0
         Width           =   7665
      End
   End
   Begin VB.PictureBox picStep3 
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   7575
      TabIndex        =   10
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txtRpt_OSMS 
         Height          =   375
         Left            =   3630
         TabIndex        =   31
         Top             =   3315
         Width           =   3195
      End
      Begin VB.CommandButton Command7 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   30
         Top             =   3330
         Width           =   345
      End
      Begin VB.CommandButton Command9 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   29
         Top             =   3720
         Width           =   345
      End
      Begin VB.CommandButton Command6 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   28
         Top             =   2940
         Width           =   345
      End
      Begin VB.CommandButton Command5 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   27
         Top             =   2520
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   26
         Top             =   2115
         Width           =   345
      End
      Begin VB.CommandButton Command2 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   25
         Top             =   1695
         Width           =   345
      End
      Begin VB.TextBox txtRpt_SMIS 
         Height          =   375
         Left            =   3630
         TabIndex        =   24
         Top             =   3720
         Width           =   3195
      End
      Begin VB.TextBox txtRpt_HRMS 
         Height          =   375
         Left            =   3630
         TabIndex        =   23
         Top             =   2925
         Width           =   3195
      End
      Begin VB.TextBox txtRpt_CSMS 
         Height          =   375
         Left            =   3630
         TabIndex        =   22
         Top             =   2505
         Width           =   3195
      End
      Begin VB.CommandButton Command1 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   18
         Top             =   1215
         Width           =   345
      End
      Begin VB.TextBox txtRpt_CRIS 
         Height          =   375
         Left            =   3630
         TabIndex        =   13
         Top             =   2100
         Width           =   3195
      End
      Begin VB.TextBox txtRpt_AMIS 
         Height          =   375
         Left            =   3630
         TabIndex        =   12
         Top             =   1230
         Width           =   3195
      End
      Begin VB.TextBox txtRpt_CMIS 
         Height          =   375
         Left            =   3630
         TabIndex        =   11
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Image Image4 
         Height          =   4125
         Left            =   0
         Picture         =   "UpdatePassword.frx":3555
         Top             =   900
         Width           =   2820
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OSMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   32
         ToolTipText     =   "Human Resource Management System"
         Top             =   3390
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   21
         ToolTipText     =   "Sales Management Information System"
         Top             =   3780
         Width           =   405
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HRMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   20
         ToolTipText     =   "Human Resource Management System"
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   19
         ToolTipText     =   "Customer Relation Information System"
         Top             =   2175
         Width           =   375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DMIS 2.0 Report Path Configuration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   210
         TabIndex        =   17
         Top             =   240
         Width           =   3360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   16
         ToolTipText     =   "Accounting Management Information System"
         Top             =   1305
         Width           =   435
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CMIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   15
         ToolTipText     =   "Cash Monitoring Information System"
         Top             =   1755
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CSMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   14
         ToolTipText     =   "Car Service Management System"
         Top             =   2580
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   885
         Left            =   0
         Picture         =   "UpdatePassword.frx":ACFB
         Top             =   0
         Width           =   7665
      End
   End
End
Attribute VB_Name = "frmUpdatePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command8_Click()

End Sub

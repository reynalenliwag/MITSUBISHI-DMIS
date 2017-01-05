VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{E6BE8522-29DC-4EDD-813C-BAA34BBA1069}#2.0#0"; "WIZMACFORM.OCX"
Begin VB.Form frmToolsBackup 
   BorderStyle     =   0  'None
   Caption         =   "BACKUP/RESTORE Database"
   ClientHeight    =   2295
   ClientLeft      =   3015
   ClientTop       =   2715
   ClientWidth     =   4050
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ToolBackUpData.frx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   Begin wizMacForm.wizMacApp wizMacApp1 
      Height          =   320
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   556
      MacCaption      =   "Mac Caption"
      Object.ToolTipText     =   "MAC titlebars can even have tooltips"
   End
   Begin VB.TextBox txtCommandLine 
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Text            =   "backup.bat"
      Top             =   1440
      Width           =   3855
   End
   Begin VB.PictureBox ssfrmChoose 
      Height          =   735
      Left            =   60
      Picture         =   "ToolBackUpData.frx":2D3C
      ScaleHeight     =   675
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   360
      Width           =   1875
      Begin VB.OptionButton optWhat 
         Caption         =   "&Restore Database"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   1605
      End
      Begin VB.OptionButton optWhat 
         Caption         =   "&Backup Database"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   1605
      End
   End
   Begin wizButton.cmd cmdOK 
      Height          =   345
      Left            =   720
      TabIndex        =   5
      Top             =   1890
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      TX              =   "&Okey"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "ToolBackUpData.frx":5A78
   End
   Begin wizButton.cmd cmdCancel 
      Height          =   345
      Left            =   2100
      TabIndex        =   6
      Top             =   1890
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "ToolBackUpData.frx":5BDA
   End
   Begin VB.Label Label1 
      Caption         =   "Command Line:"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   1170
      Width           =   1455
   End
End
Attribute VB_Name = "frmToolsBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
UnloadForm Me
End Sub

Private Sub Form_Load()
optWhat(0).Value = True
CenterMe frmMain, Me, 1
DrawXPCtl Me
End Sub

Private Sub optWhat_GotFocus(Index As Integer)
If Index = 0 Then
   txtCommandLine.Text = "BACKUP.BAT"
Else
   txtCommandLine.Text = "RESTORE.BAT"
End If
End Sub

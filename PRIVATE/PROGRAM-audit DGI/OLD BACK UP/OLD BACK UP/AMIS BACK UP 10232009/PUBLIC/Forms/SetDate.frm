VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmSetDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Date"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2640
   ForeColor       =   &H8000000F&
   Icon            =   "SetDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2640
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox txtDate 
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Top             =   90
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm/dd/yyyy"
      PromptChar      =   "_"
   End
   Begin wizButton.cmd cmdOkey 
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   510
      Width           =   915
      _ExtentX        =   1614
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
      MPTR            =   0
      MICON           =   "SetDate.frx":0442
   End
   Begin wizButton.cmd cmdCancel 
      Height          =   345
      Left            =   1350
      TabIndex        =   2
      Top             =   510
      Width           =   915
      _ExtentX        =   1614
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
      MPTR            =   0
      MICON           =   "SetDate.frx":045E
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmSetDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOkey_Click()
On Error GoTo ErrorCode
If IsDate(txtDate.Text) = True Then
   'LOGDATE = txtDate.Text
   'frmMain.BarDate.Caption = Format(LOGDATE, "long date")
   'Unload Me
Else
   MsgBoxXP "Invalid Date!", vbCritical
End If
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Private Sub Form_Load()
CenterMe frmMain, Me, 1
txtDate.Text = LOGDATE
End Sub

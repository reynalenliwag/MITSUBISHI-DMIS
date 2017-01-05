VERSION 5.00
Object = "{A06473E6-73D7-426E-82F2-6CD4F1FA4DBE}#1.0#0"; "WIZMACBUT.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCorpAplForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Application Data Entry for Corporate"
   ClientHeight    =   14160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   ForeColor       =   &H00FFFFFF&
   Icon            =   "LoanCorporate.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   14160
   ScaleWidth      =   11430
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   14085
      Left            =   0
      ScaleHeight     =   14085
      ScaleWidth      =   10635
      TabIndex        =   3
      Top             =   0
      Width           =   10635
      Begin VB.PictureBox picCorp 
         Height          =   14070
         Left            =   -30
         ScaleHeight     =   14010
         ScaleWidth      =   10545
         TabIndex        =   4
         Top             =   -30
         Width           =   10605
      End
   End
   Begin wizMacBut.MacBut cmdCancel 
      Height          =   345
      Left            =   8880
      TabIndex        =   1
      Top             =   6210
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "   Cancel"
   End
   Begin wizMacBut.MacBut cmdSave 
      Height          =   345
      Left            =   7140
      TabIndex        =   2
      Top             =   6210
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "     Save"
   End
   Begin MSForms.ScrollBar ScrollBar1 
      Height          =   6135
      Left            =   10740
      TabIndex        =   0
      Top             =   180
      Width           =   315
      Size            =   "556;10821"
      Max             =   2800
      SmallChange     =   500
      LargeChange     =   500
      Delay           =   0
   End
End
Attribute VB_Name = "frmCorpAplForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Height = 7215
    Picture1.Height = 6195
End Sub

Private Sub ScrollBar1_Change()
    picCorp.Top = 0 - ScrollBar1.Value
End Sub

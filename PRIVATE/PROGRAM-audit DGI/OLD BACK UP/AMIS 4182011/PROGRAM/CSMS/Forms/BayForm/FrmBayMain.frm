VERSION 5.00
Begin VB.Form FrmBayMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Bay Monitoring"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Bay Monitoring"
      Height          =   945
      Left            =   300
      TabIndex        =   1
      Top             =   1200
      Width           =   1035
   End
   Begin VB.CommandButton CmdBay 
      Caption         =   "Bay Master File"
      Height          =   945
      Left            =   300
      Picture         =   "FrmBayMain.frx":0000
      TabIndex        =   0
      Top             =   150
      Width           =   1035
   End
End
Attribute VB_Name = "FrmBayMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBay_Click()
frmBayMasterFile.Show 1
End Sub

Private Sub Command1_Click()
frmBayMonitoring.Show 1
End Sub

Private Sub Command2_Click()
frmCSMSServiceCounter.Show
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call OpenSQLDb
End Sub

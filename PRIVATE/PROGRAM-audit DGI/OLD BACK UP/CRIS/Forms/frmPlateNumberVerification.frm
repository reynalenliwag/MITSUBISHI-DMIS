VERSION 5.00
Begin VB.Form frmPlateNumberVerification 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verify Plate Number"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   Icon            =   "frmPlateNumberVerification.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVPNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   3405
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1740
      MouseIcon       =   "frmPlateNumberVerification.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "frmPlateNumberVerification.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   1020
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1020
      MouseIcon       =   "frmPlateNumberVerification.frx":0D5A
      MousePointer    =   99  'Custom
      Picture         =   "frmPlateNumberVerification.frx":0EAC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ok"
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label lblFlag 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "Please Re-Type the Plate Number:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3585
   End
End
Attribute VB_Name = "frmPlateNumberVerification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    lblFlag.Caption = "Cancel"
    frmPlateNumberVerification.Visible = False
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub txtVPNumber_Change()
    If Trim(UCase$(txtVPNumber.Text)) = frmCSMSNewAppointment.txtPlate_No Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub txtVPNumber_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then: cmdOK_Click
    KeyAscii = UpperAscii(KeyAscii)
End Sub

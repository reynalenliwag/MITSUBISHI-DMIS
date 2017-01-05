VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "CEO-PATS"
   ClientHeight    =   1845
   ClientLeft      =   4620
   ClientTop       =   3285
   ClientWidth     =   5130
   Icon            =   "Password.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1845
   ScaleWidth      =   5130
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "O K"
         Height          =   495
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3450
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   840
         Width           =   2505
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "Password.frx":0442
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
      
      If Text1.Text <> OldPass Then
      'If Trim(UCase(Text1.Text)) <> "" Then
      
      
            MsgBox "ENTRY DENIED, Invalid Password"
            Text1.Text = ""
            Text1.SetFocus
            
      Else
            Unload Me
            frmEditCards.Show vbModal
      
      End If
      
      
      
      
      
End Sub


Private Sub Command2_Click()
    Unload Me
    frmLOGIN.TxtEmpNumber.SetFocus
End Sub


Private Sub Text1_GotFocus()

            Text1.Text = ""
            
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      SendKeys "{TAB}"
   End If
End Sub

VERSION 5.00
Begin VB.Form frmCRIS_CustomerContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Contact Information"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CustomerContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDeleteContact 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2100
      MouseIcon       =   "CustomerContact.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "CustomerContact.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   645
   End
   Begin VB.CommandButton cmdSaveContact 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2790
      MouseIcon       =   "CustomerContact.frx":0D47
      MousePointer    =   99  'Custom
      Picture         =   "CustomerContact.frx":0E99
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   645
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
      Height          =   675
      Left            =   3540
      MouseIcon       =   "CustomerContact.frx":11E9
      MousePointer    =   99  'Custom
      Picture         =   "CustomerContact.frx":133B
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3360
      Width           =   645
   End
   Begin VB.TextBox txtContactAddress 
      Height          =   645
      Left            =   1140
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2520
      Width           =   3045
   End
   Begin VB.TextBox txtContactMobile 
      Height          =   345
      Left            =   1140
      TabIndex        =   5
      Top             =   2115
      Width           =   3045
   End
   Begin VB.TextBox txtContactPhone 
      Height          =   345
      Left            =   1140
      TabIndex        =   4
      Top             =   1725
      Width           =   3045
   End
   Begin VB.TextBox txtContactDepartment 
      Height          =   345
      Left            =   1140
      TabIndex        =   3
      Top             =   1320
      Width           =   3045
   End
   Begin VB.TextBox txtContactPosition 
      Height          =   345
      Left            =   1140
      TabIndex        =   2
      Top             =   915
      Width           =   3045
   End
   Begin VB.ComboBox cboContactRelation 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   345
      ItemData        =   "CustomerContact.frx":1679
      Left            =   1140
      List            =   "CustomerContact.frx":167B
      TabIndex        =   1
      Top             =   525
      Width           =   3045
   End
   Begin VB.TextBox txtContactName 
      Height          =   345
      Left            =   1140
      TabIndex        =   0
      Top             =   120
      Width           =   3045
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   345
      TabIndex        =   13
      Top             =   2700
      Width           =   765
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   510
      TabIndex        =   12
      Top             =   2280
      Width           =   600
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   525
      TabIndex        =   11
      Top             =   1860
      Width           =   585
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Position:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   375
      TabIndex        =   9
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   570
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Relation:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   375
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmCRIS_CustomerContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCloseContactsAE_Click(Index As Integer)
    Unload Me
End Sub


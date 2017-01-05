VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_CustomerChild 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Children Information"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CustomerChild.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2895
      MouseIcon       =   "CustomerChild.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "CustomerChild.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   645
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&OK"
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
      Left            =   2220
      MouseIcon       =   "CustomerChild.frx":0D5A
      MousePointer    =   99  'Custom
      Picture         =   "CustomerChild.frx":0EAC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   645
   End
   Begin VB.OptionButton optUnspecified 
      Caption         =   "Unspecified"
      Height          =   225
      Left            =   2220
      TabIndex        =   6
      Top             =   780
      Width           =   1395
   End
   Begin VB.OptionButton optFemale 
      Caption         =   "Female"
      Height          =   225
      Left            =   1200
      TabIndex        =   5
      Top             =   780
      Width           =   1035
   End
   Begin VB.OptionButton optMale 
      Caption         =   "Male"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   915
   End
   Begin MSComCtl2.DTPicker txtChildDOB 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   39258
   End
   Begin VB.TextBox txtChildName 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Width           =   3075
   End
   Begin VB.Label Label2 
      Caption         =   "Date Of Birth"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Child Name"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
End
Attribute VB_Name = "frmCRIS_CustomerChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub


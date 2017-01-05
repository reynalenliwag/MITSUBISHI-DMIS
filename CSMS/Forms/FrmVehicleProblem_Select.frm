VERSION 5.00
Begin VB.Form frmCSMSVehicleProblem_Select 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selected Vehicle Problem / Notes"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmVehicleProblem_Select.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9105
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "[ Printed Description ]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3945
      Left            =   390
      TabIndex        =   4
      Top             =   750
      Width           =   8325
      Begin VB.TextBox txtDesc 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Text            =   "FrmVehicleProblem_Select.frx":014A
         Top             =   930
         Width           =   8055
      End
      Begin VB.TextBox txtCustomertext 
         Height          =   345
         Left            =   1530
         TabIndex        =   1
         Text            =   "Customer state that"
         Top             =   510
         Width           =   4125
      End
      Begin VB.Label Label2 
         Caption         =   "Lead in text"
         Height          =   255
         Left            =   270
         TabIndex        =   5
         Top             =   540
         Width           =   1365
      End
   End
   Begin VB.TextBox txtProblem 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   150
      Width           =   5925
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
      Left            =   8040
      MouseIcon       =   "FrmVehicleProblem_Select.frx":0150
      MousePointer    =   99  'Custom
      Picture         =   "FrmVehicleProblem_Select.frx":02A2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancel"
      Top             =   4740
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
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
      Height          =   795
      Left            =   7320
      MouseIcon       =   "FrmVehicleProblem_Select.frx":05E0
      MousePointer    =   99  'Custom
      Picture         =   "FrmVehicleProblem_Select.frx":0732
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ok"
      Top             =   4740
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Problem  : "
      Height          =   315
      Left            =   390
      TabIndex        =   3
      Top             =   240
      Width           =   1125
   End
End
Attribute VB_Name = "frmCSMSVehicleProblem_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SendKeys "{end}"
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8490
      Left            =   0
      ScaleHeight     =   8460
      ScaleWidth      =   8025
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.OptionButton optSelect 
         Caption         =   "Last Name"
         CausesValidation=   0   'False
         Height          =   345
         Index           =   1
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   615
         Value           =   -1  'True
         Width           =   1830
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Account Name"
         CausesValidation=   0   'False
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   1830
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1950
         TabIndex        =   1
         Top             =   1050
         Width           =   2235
      End
      Begin MSComctlLib.ListView lvSearch 
         CausesValidation=   0   'False
         Height          =   4500
         Left            =   1980
         TabIndex        =   4
         Top             =   1380
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   7938
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form2.frx":0000
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4092
         EndProperty
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSIddleTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change reason for Idle Time"
   ClientHeight    =   3930
   ClientLeft      =   7290
   ClientTop       =   1200
   ClientWidth     =   6210
   ForeColor       =   &H8000000F&
   Icon            =   "FrmiddleTime.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6210
   Begin MSComctlLib.ListView lstIdle 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   510
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmiddleTime.frx":01CA
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Current Reason for Clocking Out of Open Job"
         Object.Width           =   11465
      EndProperty
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
      Height          =   735
      Left            =   5400
      MouseIcon       =   "FrmiddleTime.frx":032C
      MousePointer    =   99  'Custom
      Picture         =   "FrmiddleTime.frx":047E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      MouseIcon       =   "FrmiddleTime.frx":07BC
      MousePointer    =   99  'Custom
      Picture         =   "FrmiddleTime.frx":090E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Select"
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Please choose the primary reason for going to Idle Time :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   5385
   End
End
Attribute VB_Name = "frmCSMSIddleTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim rsUpload                                                      As ADODB.Recordset
    Set rsUpload = New ADODB.Recordset
    lstIdle.Enabled = False
    lstIdle.Sorted = False: lstIdle.ListItems.Clear
    lstIdle.Enabled = True
End Sub
Private Sub lstIdle_DblClick()
    cmdSelect.Value = True
End Sub

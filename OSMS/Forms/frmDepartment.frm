VERSION 5.00
Begin VB.Form frmDepartment 
   Caption         =   "DEPARTMENT"
   ClientHeight    =   4005
   ClientLeft      =   2775
   ClientTop       =   2715
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   465
      Left            =   5340
      TabIndex        =   5
      Top             =   2790
      Width           =   1485
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   465
      Left            =   7740
      TabIndex        =   8
      Top             =   3450
      Width           =   795
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   465
      Left            =   300
      TabIndex        =   7
      Top             =   3450
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "DEPARTMENT DATA ENTRY"
      Height          =   2235
      Left            =   300
      TabIndex        =   9
      Top             =   300
      Width           =   8175
      Begin VB.TextBox txtDeptCode 
         Height          =   405
         Left            =   3720
         TabIndex        =   0
         Top             =   660
         Width           =   1605
      End
      Begin VB.TextBox txtDeptName 
         Height          =   405
         Left            =   3660
         TabIndex        =   1
         Top             =   1395
         Width           =   2145
      End
      Begin VB.Label Label1 
         Caption         =   "Department code:"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   11
         Top             =   690
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Department name:"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   1440
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   465
      Left            =   7020
      TabIndex        =   6
      Top             =   2790
      Width           =   1485
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   465
      Left            =   3660
      TabIndex        =   4
      Top             =   2790
      Width           =   1485
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   465
      Left            =   1980
      TabIndex        =   3
      Top             =   2790
      Width           =   1485
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   465
      Left            =   300
      TabIndex        =   2
      Top             =   2790
      Width           =   1485
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gconDepartment As ADODB.Connection

Private Sub Form_Load()
    makeConnection
End Sub

Sub makeConnection()
  gconDepartment = New ADODB.Connection
  gconDepartment.Open
End Sub


Private Sub cmdAdd_Click()
    txtDeptCode.Text = ""
    txtDeptName.Text = ""
    txtDeptCode.SetFocus
End Sub

Private Sub cmdEdit_Click()
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub


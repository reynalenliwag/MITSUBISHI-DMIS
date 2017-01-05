VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCMISCancelOREntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelled Official Receipt"
   ClientHeight    =   4080
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   9000
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CancelOR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   9000
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   -30
      Width           =   8865
      Begin VB.TextBox txtTranno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txtCheckNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3390
         TabIndex        =   1
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Cancel From :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -390
         TabIndex        =   6
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2940
         TabIndex        =   3
         Top             =   270
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   60
      TabIndex        =   7
      Top             =   570
      Width           =   8865
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3135
         Left            =   60
         TabIndex        =   8
         Top             =   180
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         FormatString    =   "  Date Cancel    |       Time       |    OR No.           |   OR Date         | Cancelled By               "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   2910
      TabIndex        =   5
      Top             =   1650
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   3240
      TabIndex        =   4
      Top             =   1620
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCMISCancelOREntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
End Sub


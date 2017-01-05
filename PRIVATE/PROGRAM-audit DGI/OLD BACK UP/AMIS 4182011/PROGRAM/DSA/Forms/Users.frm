VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Begin VB.Form frmRAM_User 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " User Administration"
   ClientHeight    =   7245
   ClientLeft      =   1560
   ClientTop       =   825
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Users.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   7665
      TabIndex        =   44
      Top             =   0
      Width           =   7695
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " User Administration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   60
         TabIndex        =   45
         Top             =   60
         Width           =   3885
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   -30
      ScaleHeight     =   3435
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   2910
      Width           =   7785
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   43
         Top             =   0
         Width           =   5685
      End
      Begin MSComctlLib.ListView lvwUsers 
         Height          =   2985
         Left            =   60
         TabIndex        =   1
         Top             =   420
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   5265
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Users.frx":08CA
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Username"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "UserGroup"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "  "
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   60
         Width           =   2475
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   60
      ScaleHeight     =   2475
      ScaleWidth      =   7575
      TabIndex        =   3
      Top             =   480
      Width           =   7575
      Begin VB.ComboBox cboEmployeeName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Users.frx":0A2C
         Left            =   1800
         List            =   "Users.frx":0A2E
         TabIndex        =   47
         Text            =   "Combo1"
         Top             =   1980
         Width           =   2805
      End
      Begin VB.Frame Frame1 
         Caption         =   "Users Module"
         Height          =   2175
         Left            =   4770
         TabIndex        =   32
         Top             =   60
         Width           =   2715
         Begin VB.CheckBox chkPMIS 
            Caption         =   "PMIS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1290
            TabIndex        =   40
            Top             =   1500
            Width           =   885
         End
         Begin VB.CheckBox chkOSMS 
            Caption         =   "OSMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1290
            TabIndex        =   39
            Top             =   1140
            Width           =   885
         End
         Begin VB.CheckBox chkSMIS 
            Caption         =   "SMIS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1290
            TabIndex        =   38
            Top             =   720
            Width           =   885
         End
         Begin VB.CheckBox chkHRMS 
            Caption         =   "HRMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1290
            TabIndex        =   37
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox chkCSMS 
            Caption         =   "CSMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   210
            TabIndex        =   36
            Top             =   1530
            Width           =   885
         End
         Begin VB.CheckBox chkCRIS 
            Caption         =   "CRIS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   210
            TabIndex        =   35
            Top             =   1110
            Width           =   885
         End
         Begin VB.CheckBox chkCMIS 
            Caption         =   "CMIS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   210
            TabIndex        =   34
            Top             =   690
            Width           =   885
         End
         Begin VB.CheckBox chkAMIS 
            Caption         =   "AMIS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   210
            TabIndex        =   33
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.TextBox txtUser_Code 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1815
         MaxLength       =   3
         TabIndex        =   16
         Top             =   90
         Width           =   2745
      End
      Begin VB.TextBox txtUser_Pass2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1815
         MaxLength       =   20
         PasswordChar    =   "l"
         TabIndex        =   7
         Top             =   1230
         Width           =   2745
      End
      Begin VB.TextBox txtUser_Pass1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1815
         MaxLength       =   20
         PasswordChar    =   "l"
         TabIndex        =   6
         Top             =   855
         Width           =   2745
      End
      Begin VB.TextBox txtUser_Name 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1815
         MaxLength       =   20
         TabIndex        =   5
         Top             =   480
         Width           =   2745
      End
      Begin VB.ComboBox cboUser_Groups 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Users.frx":0A30
         Left            =   1815
         List            =   "Users.frx":0A32
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1605
         Width           =   2805
      End
      Begin wizEncrypt.wizEnc wizEnc1 
         Left            =   2130
         Top             =   -2430
         _ExtentX        =   3969
         _ExtentY        =   3969
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   360
         TabIndex        =   46
         Top             =   2010
         Width           =   1365
      End
      Begin VB.Label olduser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3600
         TabIndex        =   31
         Top             =   150
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label labID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3660
         TabIndex        =   30
         Top             =   150
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   825
         TabIndex        =   18
         Top             =   150
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   570
         TabIndex        =   17
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   570
         TabIndex        =   15
         Top             =   1620
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   1260
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   630
         TabIndex        =   13
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   780
         TabIndex        =   11
         Top             =   1620
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   1260
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   870
         TabIndex        =   9
         Top             =   915
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   855
         TabIndex        =   8
         Top             =   510
         Width           =   885
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6210
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   27
      Top             =   6330
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   735
         MouseIcon       =   "Users.frx":0A34
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":0B86
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   45
         MouseIcon       =   "Users.frx":0EC4
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":1016
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   1620
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   19
      Top             =   6330
      Width           =   6075
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   5310
         MouseIcon       =   "Users.frx":1366
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":14B8
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "Enable"
         Height          =   795
         Left            =   4620
         MouseIcon       =   "Users.frx":181E
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":1970
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Disable"
         Height          =   795
         Left            =   4620
         MouseIcon       =   "Users.frx":216F
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":22C1
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3930
         MouseIcon       =   "Users.frx":25EC
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":273E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   3240
         MouseIcon       =   "Users.frx":2A9A
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":2BEC
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   2550
         MouseIcon       =   "Users.frx":2EFF
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":3051
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   1860
         MouseIcon       =   "Users.frx":33A9
         MousePointer    =   99  'Custom
         Picture         =   "Users.frx":34FB
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdAddMod 
         Caption         =   "Assign Modules"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1170
         MouseIcon       =   "Users.frx":385A
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdCopySetting 
         Caption         =   "Copy Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   480
         MouseIcon       =   "Users.frx":39AC
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   45
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmRAM_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'FUNCTION / FEATURE :cmdSave_Click:ADDED VALIDATION OF USERCODE AND USER NAME
'DATE STARTED       :5/31/200715:44
'LAST UPDATED       :5/31/200715:44
'DATABASE UPDATES   :
'WHO UPDATED        :AXP  5/31/2007
'UDPATING CODE    :AXP-531200715:44
'==========================================================================================
Option Explicit
Dim rsUser                                                            As ADODB.Recordset
Dim ModuleID                                                          As Long
Dim AddorEdit                                                         As String

Private Sub cboEmployeeName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboEmployeeName_LostFocus()
On Error Resume Next
cboEmployeeName.ListIndex = SelectCombo(cboEmployeeName, cboEmployeeName.Text)
If cboEmployeeName.ListIndex = -1 Then
cboEmployeeName.SetFocus
End If
End Sub

Private Sub cboUser_Groups_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    picAdds.Visible = False
    picSaves.Visible = True
    Picture1.Enabled = True
    Picture2.Enabled = False
    InitMemVar
    txtUser_Name.SetFocus
End Sub

'Upating Code       : AXP-0713200715:22
Private Sub cmdAddMod_Click()
    On Error GoTo Errorcode:

    If lvwUsers.SelectedItem Is Nothing Then
        '        MessagePop RecNotFound, "No Record", "There are No Record", 1000
        Exit Sub
    End If

    frmAcessManagement.UserID = labID
    frmAcessManagement.Username = lvwUsers.SelectedItem
    frmAcessManagement.Show 1
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picSaves.Visible = False
    picAdds.Visible = True
    Picture1.Enabled = False
    Picture2.Enabled = True
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If lvwUsers.ListItems.Count = 0 Then Exit Sub
    If lvwUsers.SelectedItem.Text = LOGNAME Then
        MsgBox "Cannot Disable Your Own User Credentials", vbInformation
        Exit Sub
    End If


    If MsgBox("Are you sure you want to Disable user " & lvwUsers.SelectedItem & "?", vbExclamation + vbYesNo, "Remove User") = vbYes Then
        If lvwUsers.ListItems.Count = 1 Then
            MsgBox "Sorry, can't remove selected user.", vbCritical, "Access denied!"
        Else
            gconDMIS.Execute ("UPDATE ALL_Rams_Users SET LOCK=1 where username = '" & lvwUsers.SelectedItem & "' ")
            rsUser.Requery
            rsUser.Find ("USERID=" & labID)
            StoreMemvars
        End If
    End If
End Sub

'Upating Code       : AXP-0713200715:22
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:

    AddorEdit = "EDIT"
    picAdds.Visible = False
    picSaves.Visible = True
    Picture1.Enabled = True
    Picture2.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200715:22
Private Sub cmdEnable_Click()

    On Error GoTo Errorcode:

    If lvwUsers.ListItems.Count = 0 Then Exit Sub



    If MsgBox("Are you sure you want to Enable user " & lvwUsers.SelectedItem & "?", vbExclamation + vbYesNo, "Remove User") = vbYes Then

        gconDMIS.Execute ("UPDATE ALL_Rams_Users SET LOCK=0 where username = '" & lvwUsers.SelectedItem & "' ")
        rsUser.Requery
        rsUser.Find ("USERID=" & labID)
        StoreMemvars


    End If
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200715:22
Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    rsUser.MoveNext
    If rsUser.EOF Then
        rsUser.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError

End Sub

'Upating Code       : AXP-0713200715:22
Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rsUser.MovePrevious
    If rsUser.BOF Then
        rsUser.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdSave_Click()
'UDPATING COCODE    :AXP-531200715:44
    Dim SQL                                                           As String
    Dim lng                                                           As Integer
    Dim empno                                                         As String
    Dim TEMPRS                                                        As ADODB.Recordset
    '  On Error GoTo cmdSave_Click_Error

    If RTrim(LTrim(txtUser_Code)) = "" Then
        ShowIsRequiredMsg "User Name"
        txtUser_Code.SetFocus
        Exit Sub
    End If



    If RTrim(LTrim(txtUser_Name)) = "" Then
        ShowIsRequiredMsg "User Name"
        txtUser_Name.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtUser_Pass1)) = "" Then
        ShowIsRequiredMsg "Password"
        txtUser_Pass1.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtUser_Pass2)) = "" Then
        ShowIsRequiredMsg "Confirm Password"
        txtUser_Pass2.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtUser_Pass2)) <> RTrim(LTrim(txtUser_Pass2)) Then
        MsgSpeechBox "Passwords do not match!"
        txtUser_Pass1.SetFocus
        Exit Sub
    End If
    If cboUser_Groups = "" Then
        ShowIsRequiredMsg "Level"
        cboUser_Groups.SetFocus
        Exit Sub
    End If
    empno = ""
    
     If cboEmployeeName.ListIndex <> -1 Then
        empno = GETEMPNO(cboEmployeeName.ItemData(cboEmployeeName.ListIndex))
     End If
    ''''''USER CODE
    lng = gconDMIS.Execute("select Count(*) from ALL_RAMS_USERs WHERE USERCODE=" & N2Str2Null(txtUser_Code)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            txtUser_Code.SetFocus
            Exit Sub
        End If
    Else
        If lng >= 1 And LTrim(RTrim(UCase(Null2String(rsUser!USERCODE)))) <> LTrim(RTrim(UCase(txtUser_Code))) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            txtUser_Code.SetFocus
            Exit Sub
        End If
    End If
    ''''''USER NAME
    lng = gconDMIS.Execute("select Count(*) from ALL_RAMS_USERs WHERE USERNAME=" & N2Str2Null(txtUser_Name)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "User Name Already Exist"
            txtUser_Name.SetFocus
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsUser!Username)) <> UCase(txtUser_Name) Then
            MessagePop RecSaveWarning, "Duplicate Record", "User Name Already Exist"
            txtUser_Name.SetFocus
            Exit Sub
        End If
    End If
    Dim currentid                                                     As Long
    If AddorEdit = "ADD" Then
        SQL = "INSERT INTO ALL_RAMS_USERS (EMPNO, USERCODE,USERNAME, PASSWORD, USERGROUP, LOCK)  VALUES ("
        SQL = SQL & N2Str2Null(empno) & " ," & vbCrLf
        SQL = SQL & N2Str2Null(txtUser_Code) & " ," & vbCrLf
        SQL = SQL & N2Str2Null(txtUser_Name) & " ," & vbCrLf
        SQL = SQL & N2Str2Null(wizEnc1.EncryptAccess(txtUser_Pass1)) & " ,"
        SQL = SQL & N2Str2Null(cboUser_Groups) & " ,0)"
        gconDMIS.Execute SQL
        currentid = gconDMIS.Execute("select max(USERID) FROM all_rams_users").Fields(0).Value
    Else

        SQL = "Update ALL_RAMS_USERS SET " & vbCrLf
        SQL = SQL & " USERCODE=" & N2Str2Null(txtUser_Code) & ", " & vbCrLf
        SQL = SQL & " PASSWORD=" & N2Str2Null(wizEnc1.EncryptAccess(txtUser_Pass1)) & ", " & vbCrLf
        SQL = SQL & " username=" & N2Str2Null(txtUser_Name) & ", " & vbCrLf
        SQL = SQL & " usergroup=" & N2Str2Null(cboUser_Groups) & ", " & vbCrLf
        SQL = SQL & " empno=" & N2Str2Null(empno) & ", " & vbCrLf
        SQL = SQL & " lock=0" & vbCrLf
        SQL = SQL & " where userid=" & labID
        currentid = labID
        gconDMIS.Execute SQL
    End If

    If UCase(cboUser_Groups) <> "SDM" Then
        gconDMIS.Execute ("DELETE FROM ALL_RAMS_USER_MODULES WHERE USERID=" & currentid)
        If chkAMIS.Value = 1 Then: gconDMIS.Execute ("Insert Into ALL_RAMS_USER_MODULES (USERID,MAINMODULENAME) VALUES(" & currentid & ", 'AMIS')")
        If chkCMIS.Value = 1 Then: gconDMIS.Execute ("Insert Into ALL_RAMS_USER_MODULES (USERID,MAINMODULENAME) VALUES(" & currentid & ", 'CMIS')")
        If chkCRIS.Value = 1 Then: gconDMIS.Execute ("Insert Into ALL_RAMS_USER_MODULES (USERID,MAINMODULENAME) VALUES(" & currentid & ", 'CRIS')")
        If chkCSMS.Value = 1 Then: gconDMIS.Execute ("Insert Into ALL_RAMS_USER_MODULES (USERID,MAINMODULENAME) VALUES(" & currentid & ", 'CSMS')")
        If chkHRMS.Value = 1 Then: gconDMIS.Execute ("Insert Into ALL_RAMS_USER_MODULES (USERID,MAINMODULENAME) VALUES(" & currentid & ", 'HRMS')")
        If chkOSMS.Value = 1 Then: gconDMIS.Execute ("Insert Into ALL_RAMS_USER_MODULES (USERID,MAINMODULENAME) VALUES(" & currentid & ", 'OSMS')")
        If chkSMIS.Value = 1 Then: gconDMIS.Execute ("Insert Into ALL_RAMS_USER_MODULES (USERID,MAINMODULENAME) VALUES(" & currentid & ", 'SMIS')")
        If chkPMIS.Value = 1 Then: gconDMIS.Execute ("Insert Into ALL_RAMS_USER_MODULES (USERID,MAINMODULENAME) VALUES(" & currentid & ", 'PMIS')")
    End If

    If AddorEdit = "ADD" Then
        MessagePop RecSaveOk, "New User", "User " & Trim(txtUser_Name) & " successfully Added!"
    Else
        MessagePop RecSaveOk, "User Info Updated ", "User " & Trim(txtUser_Name) & " successfully Updated!"
    End If
    rsRefresh
    If AddorEdit = "EDIT" Then
        rsUser.Find ("USERID=" & labID)
    End If
    cmdCancel.Value = True
    FillSearchGrid
    Exit Sub
    'adder:
    'MsgBox Err.Description
    'Err.Clear
    '  On Error GoTo 0
    '  Exit Sub
    'cmdSave_Click_Error:
    '   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSave_Click of Form frmRAM_User"
End Sub

Private Sub FillSearchGrid()
    Dim RS                                                            As ADODB.Recordset
    If LTrim(RTrim(Text1)) = "" Then

        Set RS = gconDMIS.Execute("Select username, usergroup, userid from ALL_RAMS_USERS order by username")
    Else
        Set RS = gconDMIS.Execute("Select username, usergroup, userid from ALL_RAMS_USERS where username like '%" & Repleys(Text1) & "%' order by username")
    End If
    If Not (RS.BOF And RS.EOF) Then
        Listview_Loadval lvwUsers.ListItems, RS
        lvwUsers.Enabled = True
    Else
        lvwUsers.Enabled = False
        lvwUsers.ListItems.Clear
    End If
    Set RS = Nothing
End Sub

'Upating Code       : AXP-0713200715:22
Private Sub cmdCopySetting_Click()
    On Error GoTo Errorcode:

    frmCopySetting.Show 1

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        KeyCode = 0
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_Load()

    InitMemVar


    picSaves.Visible = False
    picAdds.Visible = True
    Picture1.Enabled = False
    Picture2.Enabled = True

    Call Combo_Loadval(cboUser_Groups, gconDMIS.Execute("Select code,groupname from ALL_RAMS_USERGROUPS"))
    lvwUsers.ColumnHeaders(1).Width = 0.5 * lvwUsers.Width
    lvwUsers.ColumnHeaders(2).Width = 0.5 * lvwUsers.Width
    Call FillCombo("SELECT ID, upper( LASTNAME + isnull(' ,' + FIRSTNAME,'') + ISNULL('.' + LEFT(MIDDLENAME,1),'')) FROM HRMS_EMPINFO ORDER BY 1", 0, 1, cboEmployeeName)


    Call FillSearchGrid
    rsRefresh
    StoreMemvars

    If MODULENAME = "" Then Exit Sub
    On Error GoTo adder

    ModuleID = gconDMIS.Execute("SELECT ID FROM ALL_Profile WHERE MODULENAME='" & MODULENAME & "'").Fields(0).Value

    Exit Sub
adder:
    MsgBox "ERROR"
End Sub

Private Sub InitMemVar()
    txtUser_Name = vbNullString
    txtUser_Code = vbNullString
    txtUser_Pass1 = vbNullString
    txtUser_Pass2 = vbNullString
    cboUser_Groups.ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AddorEdit = ""
End Sub

Private Sub lvwUsers_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lvwUsers_ItemClick(ByVal item As MSComctlLib.ListItem)
    labID = item.SubItems(2)
    rsUser.MoveFirst
    rsUser.Find "USERID=" & labID
    StoreMemvars
End Sub

Sub rsRefresh()
    Set rsUser = New ADODB.Recordset
    Call rsUser.Open("SELECT * FROM ALL_Rams_Users order by userid desc", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Private Sub StoreMemvars()
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim TempModuleID                                                  As String
    If Not (rsUser.EOF Or rsUser.BOF) Then
        labID = Null2String(rsUser!UserID)
        txtUser_Name = Null2String(rsUser!Username)
        txtUser_Code = Null2String(rsUser!USERCODE)
        txtUser_Pass1 = wizEnc1.DecryptAccess(rsUser!Password)
        txtUser_Pass2 = txtUser_Pass1
        cboUser_Groups = Null2String(rsUser!userGroup)

        chkAMIS.Value = 0
        chkCMIS.Value = 0
        chkCRIS.Value = 0
        chkCSMS.Value = 0
        chkPMIS.Value = 0
        chkHRMS.Value = 0
        chkOSMS.Value = 0
        chkSMIS.Value = 0
        chkPMIS.Value = 0
        If rsUser!Lock = True Then
            cmdDelete.Visible = False
            cmdEnable.Visible = True
            cmdCopySetting.Enabled = False
        Else
            cmdCopySetting.Enabled = True
            cmdDelete.Visible = True
            cmdEnable.Visible = False
        End If

        Set TEMPRS = gconDMIS.Execute("SELECT * FROM ALL_RAMS_USER_MODULES WHERE USERID=" & labID)
        If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
            While Not TEMPRS.EOF
                TempModuleID = Null2String(TEMPRS!MAINMODULENAME)
                If TempModuleID = "AMIS" Then
                    chkAMIS.Value = 1
                End If
                If TempModuleID = "CMIS" Then
                    chkCMIS.Value = 1
                End If
                If TempModuleID = "CRIS" Then
                    chkCRIS.Value = 1
                End If
                If TempModuleID = "CSMS" Then
                    chkCSMS.Value = 1
                End If
                If TempModuleID = "HRMS" Then
                    chkHRMS.Value = 1
                End If
                If TempModuleID = "SMIS" Then
                    chkSMIS.Value = 1
                End If
                If TempModuleID = "OSMS" Then
                    chkOSMS.Value = 1
                End If
                If TempModuleID = "PMIS" Then
                    chkPMIS.Value = 1
                End If
                TEMPRS.MoveNext
            Wend
        End If
        Set TEMPRS = Nothing

        If Null2String(rsUser!userGroup) = "SDM" Then
            cmdAddMod.Enabled = False
            cmdCopySetting.Enabled = False
            Frame1.Enabled = False
        Else
            cmdAddMod.Enabled = True
            cmdCopySetting.Enabled = True
            Frame1.Enabled = True
        End If

        cboEmployeeName = SETEMPNO(Null2String(rsUser!empno))

    Else
        ShowNoRecord
        cmdCancel.Value = True
    End If

End Sub

Private Sub Text1_Change()
    FillSearchGrid
End Sub

Private Sub txtUser_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtUser_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtUser_Pass1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtUser_Pass2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub



Function GETEMPNO(xxx)
    Dim TEMPRS                                                        As ADODB.Recordset
    GETEMPNO = ""
    Set TEMPRS = gconDMIS.Execute("Select empno from hrms_empinfo where id=" & xxx)
    If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
        GETEMPNO = Null2String(TEMPRS!empno)
    End If
    Set TEMPRS = Nothing
End Function

Function SETEMPNO(xxx)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT upper(LASTNAME + isnull(' ,' + FIRSTNAME,'') + ISNULL('.' + LEFT(MIDDLENAME,1),'')) AS FULLNAME FROM HRMS_EMPINFO where EMPNO='" & xxx & "'")
    If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
        SETEMPNO = Null2String(TEMPRS!FULLNAME)
    End If
    Set TEMPRS = Nothing
End Function


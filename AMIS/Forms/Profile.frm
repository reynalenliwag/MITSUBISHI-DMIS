VERSION 5.00
Begin VB.Form frmAMISProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Setup"
   ClientHeight    =   7350
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   7215
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Profile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   7215
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   90
      ScaleHeight     =   465
      ScaleWidth      =   5175
      TabIndex        =   45
      Top             =   6660
      Width           =   5205
      Begin VB.TextBox txtCUT_OFF_DATE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         TabIndex        =   47
         Top             =   30
         Width           =   3225
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   46
         Top             =   90
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   5340
      Picture         =   "Profile.frx":0E42
      ScaleHeight     =   2820
      ScaleWidth      =   1785
      TabIndex        =   27
      Top             =   3450
      Width           =   1815
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   5340
      ScaleHeight     =   975
      ScaleWidth      =   1785
      TabIndex        =   21
      Top             =   2220
      Width           =   1815
      Begin VB.CommandButton cmdClosing 
         Caption         =   "Journal Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   60
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "Set - Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   60
         TabIndex        =   43
         Top             =   90
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdShowBooks 
         Caption         =   "Show Open Books"
         Height          =   375
         Left            =   60
         TabIndex        =   41
         Top             =   90
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtPeriodYear 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   720
         TabIndex        =   24
         Top             =   510
         Width           =   975
      End
      Begin VB.TextBox txtPeriodMonth 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   22
         Top             =   90
         Width           =   975
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   60
         TabIndex        =   25
         Top             =   570
         Width           =   1725
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   60
         TabIndex        =   23
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   90
      ScaleHeight     =   2055
      ScaleWidth      =   5175
      TabIndex        =   28
      Top             =   4320
      Width           =   5205
      Begin VB.TextBox txtPreparedBy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         TabIndex        =   30
         Top             =   60
         Width           =   3225
      End
      Begin VB.TextBox txtCheckedBy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         TabIndex        =   32
         Top             =   450
         Width           =   3225
      End
      Begin VB.TextBox txtApprovedBy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         TabIndex        =   34
         Top             =   840
         Width           =   3225
      End
      Begin VB.TextBox txtNotedBy1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         TabIndex        =   36
         Top             =   1230
         Width           =   3225
      End
      Begin VB.TextBox txtNotedBy2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         TabIndex        =   38
         Top             =   1620
         Width           =   3225
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   29
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Checked By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   31
         Top             =   510
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   33
         Top             =   900
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "1st Noted By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   35
         Top             =   1260
         Width           =   1725
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Noted By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   37
         Top             =   1650
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   90
      ScaleHeight     =   1785
      ScaleWidth      =   5175
      TabIndex        =   12
      Top             =   2220
      Width           =   5205
      Begin VB.TextBox txtGeneralManager 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1890
         TabIndex        =   13
         Top             =   90
         Width           =   3225
      End
      Begin VB.TextBox txtSBManager 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1890
         TabIndex        =   18
         Top             =   900
         Width           =   3225
      End
      Begin VB.TextBox txtAccountNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         TabIndex        =   16
         Top             =   510
         Width           =   3225
      End
      Begin VB.TextBox txtCorpSec 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1890
         TabIndex        =   20
         Top             =   1320
         Width           =   3225
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "General Manager"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Manager"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   570
         Width           =   1725
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Secretary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   19
         Top             =   1380
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   90
      ScaleHeight     =   1665
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   270
      Width           =   7065
      Begin VB.TextBox txtCompanyTINNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   7
         Top             =   840
         Width           =   2265
      End
      Begin VB.TextBox txtCompanyAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   5
         Top             =   450
         Width           =   5055
      End
      Begin VB.TextBox txtCompanyName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   3
         Top             =   60
         Width           =   5055
      End
      Begin VB.TextBox txtCompanySSSNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1230
         Width           =   2265
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company TIN No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   870
         Width           =   1785
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   4
         Top             =   510
         Width           =   1785
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   1785
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company SSS No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   1260
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   6450
      MouseIcon       =   "Profile.frx":28AA
      MousePointer    =   99  'Custom
      Picture         =   "Profile.frx":29FC
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Close Window"
      Top             =   6360
      Width           =   705
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
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
      Left            =   5760
      MouseIcon       =   "Profile.frx":2D3A
      MousePointer    =   99  'Custom
      Picture         =   "Profile.frx":2E8C
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Save Changes"
      Top             =   6360
      Width           =   705
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cut-Off Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   90
      TabIndex        =   44
      Top             =   6420
      Width           =   3345
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Period"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5370
      TabIndex        =   11
      Top             =   1980
      Width           =   1785
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Signatories:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   90
      TabIndex        =   26
      Top             =   4080
      Width           =   3345
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Officers Information:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   90
      TabIndex        =   10
      Top             =   1980
      Width           =   3345
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Information:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   3345
   End
End
Attribute VB_Name = "frmAMISProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsProfile                                               As ADODB.Recordset

Sub StoreMemVars()
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        txtCompanyName.Text = Null2String(rsProfile!CompanyName)
        txtCompanyAddress.Text = Null2String(rsProfile!Companyaddress)
        txtCompanyTINNo.Text = Null2String(rsProfile!companytinno)
        txtCompanySSSNo.Text = Null2String(rsProfile!companysssno)
        txtPreparedBy.Text = Null2String(rsProfile!PreparedBy)
        txtCheckedBy.Text = Null2String(rsProfile!CheckedBy)
        txtApprovedBy.Text = Null2String(rsProfile!ApprovedBy)
        txtNotedBy1.Text = Null2String(rsProfile!notedby1)
        txtNotedBy2.Text = Null2String(rsProfile!notedby2)
        txtGeneralManager.Text = Null2String(rsProfile!GeneralManager)
        txtAccountNo.Text = Null2String(rsProfile!ACCOUNTNO)
        txtSBManager.Text = Null2String(rsProfile!bankmanager)
        txtCorpSec.Text = Null2String(rsProfile!SECRETARY)
        txtPeriodMonth.Text = Null2String(rsProfile!PERIODMONTH)
        txtPeriodYear.Text = Null2String(rsProfile!PERIODYEAR)
        txtCUT_OFF_DATE.Text = Null2String(rsProfile!Cut_Off_Date)
    End If
End Sub

Sub rsRefresh()
    Set rsProfile = New ADODB.Recordset
    rsProfile.Open "select * from ALL_Profile Where ModuleName = 'AMIS'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub cmdClosing_Click()
    If Module_Access(LOGID, "ACCOUNTING PERIOD CLOSING", "SYSTEM") = False Then Exit Sub
    FormExistsShow frmAccountingPeriodClosing
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    If txtCompanyName.Text = "" Then
        MessagePop RecSaveError, "Company Must have a Name", "Invalid Name"
        Exit Sub
    End If
    If txtCompanyAddress.Text = "" Then
        MessagePop RecSaveError, "Company Must have a Address", "Invalid Address", 2500
        Exit Sub
    End If
    If txtCompanyTINNo.Text = "" Then
        If MsgBox("TIN No. Omitted. Continue Any way?", vbYesNo + vbQuestion, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If
    If txtCompanySSSNo.Text = "" Then
        If MsgBox("SSS No. Omitted. Continue Any way?", vbYesNo + vbQuestion, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If

    If txtCUT_OFF_DATE.Text = "" Then
        MessagePop RecSaveError, "SYSTEM INFORMATION", "Cut-Off Date Must have a value."
        Exit Sub
    End If

    gconDMIS.Execute "update ALL_PROFILE set" & _
                     " CompanyName = " & N2Str2Null(txtCompanyName.Text) & "," & _
                     " CompanyAddress = " & N2Str2Null(txtCompanyAddress.Text) & "," & _
                     " CompanyTINNo = " & N2Str2Null(txtCompanyTINNo.Text) & "," & _
                     " CompanySSSNo = " & N2Str2Null(txtCompanySSSNo.Text) & "," & _
                     " preparedby = " & N2Str2Null(txtPreparedBy.Text) & "," & _
                     " checkedby = " & N2Str2Null(txtCheckedBy.Text) & "," & _
                     " approvedby = " & N2Str2Null(txtApprovedBy.Text) & "," & _
                     " notedby1 = " & N2Str2Null(txtNotedBy1.Text) & "," & _
                     " notedby2 = " & N2Str2Null(txtNotedBy2.Text) & "," & _
                     " generalmanager = " & N2Str2Null(txtGeneralManager.Text) & "," & _
                     " accountno = " & N2Str2Null(txtAccountNo.Text) & "," & _
                     " Bankmanager = " & N2Str2Null(txtSBManager.Text) & "," & _
                     " periodMonth = " & N2Str2Zero(txtPeriodMonth.Text) & "," & _
                     " periodYear = " & N2Str2Zero(txtPeriodYear.Text) & "," & _
                     " usercode = " & N2Str2Null(LOGCODE) & "," & _
                     " lastupdate = " & N2Str2Null(LOGDATE) & "," & _
                     " Secretary = " & N2Str2Null(txtCorpSec.Text) & "," & _
                     " CUT_OFF_DATE = " & N2Date2Null(txtCUT_OFF_DATE.Text) & _
                     " Where ModuleName = 'AMIS'"
    ShowSuccessFullyUpdated
    LogAudit "E", "COMPANY PROFILE", txtPeriodMonth & "-" & txtPeriodYear
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSetup_Click()
    If Module_Access(LOGID, "ACCOUNTING PERIOD SETUP", "SYSTEM") = False Then Exit Sub
    FormExistsShow frmAccountingPeriod
End Sub

Private Sub cmdShowBooks_Click()
'COMMENTED BY: ACL
'frmAMISAccountingPeriod.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    StoreMemVars
    Screen.MousePointer = 0

    cmdSetup.Visible = True
    cmdClosing.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAMISProfile = Nothing
End Sub


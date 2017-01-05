VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmForm2316 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form 2316"
   ClientHeight    =   7575
   ClientLeft      =   1620
   ClientTop       =   630
   ClientWidth     =   8865
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8865
   Begin Crystal.CrystalReport rptForm2316 
      Left            =   2790
      Top             =   450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Form 2316"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2745
      Left            =   60
      TabIndex        =   82
      Top             =   4830
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4842
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Present"
      TabPicture(0)   =   "Form2316.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label26"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label27"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label28"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label30"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPresentZipCode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPresentRegisteredAddress"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPresentEmployersName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPresentTIN"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Prev - 1"
      TabPicture(1)   =   "Form2316.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPrev1TIN"
      Tab(1).Control(1)=   "txtPrev1EmployersName"
      Tab(1).Control(2)=   "txtPrev1ZipCode"
      Tab(1).Control(3)=   "txtPrev1RegisteredAddress"
      Tab(1).Control(4)=   "Label34"
      Tab(1).Control(5)=   "Label33"
      Tab(1).Control(6)=   "Label32"
      Tab(1).Control(7)=   "Label31"
      Tab(1).Control(8)=   "Label29"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Prev - 2"
      TabPicture(2)   =   "Form2316.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPrev2TIN"
      Tab(2).Control(1)=   "txtPrev2EmployersName"
      Tab(2).Control(2)=   "txtPrev2RegisteredAddress"
      Tab(2).Control(3)=   "txtPrev2ZipCode"
      Tab(2).Control(4)=   "Label39"
      Tab(2).Control(5)=   "Label38"
      Tab(2).Control(6)=   "Label37"
      Tab(2).Control(7)=   "Label36"
      Tab(2).Control(8)=   "Label35"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Prev - 3"
      TabPicture(3)   =   "Form2316.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtPrev3TIN"
      Tab(3).Control(1)=   "txtPrev3EmployersName"
      Tab(3).Control(2)=   "txtPrev3RegisteredAddress"
      Tab(3).Control(3)=   "txtPrev3ZipCode"
      Tab(3).Control(4)=   "Label44"
      Tab(3).Control(5)=   "Label43"
      Tab(3).Control(6)=   "Label42"
      Tab(3).Control(7)=   "Label41"
      Tab(3).Control(8)=   "Label40"
      Tab(3).ControlCount=   9
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   88
         Top             =   2400
         Width           =   3825
         Begin VB.OptionButton optMainEmployer 
            Caption         =   "main employer"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   30
            TabIndex        =   22
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optSecondaryEmployer 
            Caption         =   "secondary employer"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1770
            TabIndex        =   23
            Top             =   0
            Width           =   2055
         End
      End
      Begin MSMask.MaskEdBox txtPresentTIN 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1020
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPresentEmployersName 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPresentRegisteredAddress 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2100
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPresentZipCode 
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   2100
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev1TIN 
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   1020
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev1EmployersName 
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   1560
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev1ZipCode 
         Height          =   255
         Left            =   -71880
         TabIndex        =   27
         Top             =   2100
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev2TIN 
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   1020
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev2EmployersName 
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   1560
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev2RegisteredAddress 
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   2100
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev2ZipCode 
         Height          =   255
         Left            =   -71880
         TabIndex        =   31
         Top             =   2100
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev3TIN 
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1020
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev3EmployersName 
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   1560
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev3RegisteredAddress 
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   2100
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev3ZipCode 
         Height          =   255
         Left            =   -71880
         TabIndex        =   35
         Top             =   2100
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrev1RegisteredAddress 
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   2100
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label44 
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71880
         TabIndex        =   103
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label43 
         Caption         =   "Registered Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   102
         Top             =   1830
         Width           =   2505
      End
      Begin VB.Label Label42 
         Caption         =   "Employer's Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   101
         Top             =   1290
         Width           =   2955
      End
      Begin VB.Label Label41 
         Caption         =   "Taxpayer Identification Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   100
         Top             =   750
         Width           =   2955
      End
      Begin VB.Label Label40 
         Caption         =   "Employer Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   99
         Top             =   450
         Width           =   3795
      End
      Begin VB.Label Label39 
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71880
         TabIndex        =   98
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label38 
         Caption         =   "Registered Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   97
         Top             =   1830
         Width           =   2505
      End
      Begin VB.Label Label37 
         Caption         =   "Employer's Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   96
         Top             =   1290
         Width           =   2955
      End
      Begin VB.Label Label36 
         Caption         =   "Taxpayer Identification Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   95
         Top             =   750
         Width           =   2955
      End
      Begin VB.Label Label35 
         Caption         =   "Employer Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   94
         Top             =   450
         Width           =   3795
      End
      Begin VB.Label Label34 
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71880
         TabIndex        =   93
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label33 
         Caption         =   "Registered Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   1830
         Width           =   2505
      End
      Begin VB.Label Label32 
         Caption         =   "Employer's Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   91
         Top             =   1290
         Width           =   2955
      End
      Begin VB.Label Label31 
         Caption         =   "Taxpayer Identification Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   90
         Top             =   750
         Width           =   2955
      End
      Begin VB.Label Label29 
         Caption         =   "Employer Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   450
         Width           =   3795
      End
      Begin VB.Label Label30 
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   87
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label28 
         Caption         =   "Registered Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1830
         Width           =   2505
      End
      Begin VB.Label Label27 
         Caption         =   "Employer's Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1290
         Width           =   2955
      End
      Begin VB.Label Label26 
         Caption         =   "Taxpayer Identification Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   750
         Width           =   2955
      End
      Begin VB.Label Label25 
         Caption         =   "Employer Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   450
         Width           =   3795
      End
   End
   Begin VB.ComboBox cboYear 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   60
      Width           =   1185
   End
   Begin VB.CommandButton cmdSelectYear 
      Caption         =   "Recompute"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2790
      TabIndex        =   57
      ToolTipText     =   "Recompute"
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print 2316 Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6990
      TabIndex        =   56
      ToolTipText     =   "Print 2316 Form"
      Top             =   60
      Width           =   1815
   End
   Begin MSMask.MaskEdBox txtTax4 
      Height          =   255
      Left            =   7590
      TabIndex        =   42
      Top             =   3090
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtSum1 
      Height          =   255
      Left            =   7590
      TabIndex        =   45
      Top             =   4440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTax1 
      Height          =   255
      Left            =   7590
      TabIndex        =   39
      Top             =   2190
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtNonTax2 
      Height          =   255
      Left            =   7590
      TabIndex        =   37
      Top             =   1170
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtSum3 
      Height          =   255
      Left            =   7590
      TabIndex        =   47
      Top             =   5160
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtSum4 
      Height          =   255
      Left            =   7590
      TabIndex        =   48
      Top             =   5490
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtSum5 
      Height          =   255
      Left            =   7590
      TabIndex        =   49
      Top             =   5910
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtNonTax1 
      Height          =   255
      Left            =   7590
      TabIndex        =   36
      Top             =   870
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTax6 
      Height          =   255
      Left            =   7590
      TabIndex        =   44
      Top             =   3720
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTax2 
      Height          =   255
      Left            =   7590
      TabIndex        =   40
      Top             =   2490
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtSum7 
      Height          =   255
      Left            =   7590
      TabIndex        =   51
      Top             =   6630
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtEmployeeName 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1050
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtRegisteredAddress 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1590
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtZipCode 
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   1590
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDateOfBirth 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2130
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTelephoneNumber 
      Height          =   255
      Left            =   2190
      TabIndex        =   5
      Top             =   2130
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtExemptionStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2670
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtQDependent1 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtQDependent2 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3990
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtQDependent3 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4260
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtQDependent4 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4530
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3240
      TabIndex        =   81
      Top             =   2970
      Width           =   675
      Begin VB.OptionButton optNoIsTheWife 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optYesIstheWife 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   675
      End
   End
   Begin MSMask.MaskEdBox txtSum2 
      Height          =   255
      Left            =   7590
      TabIndex        =   46
      Top             =   4740
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtNonTax3 
      Height          =   255
      Left            =   7590
      TabIndex        =   38
      Top             =   1470
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTax3 
      Height          =   255
      Left            =   7590
      TabIndex        =   41
      Top             =   2790
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTax5 
      Height          =   255
      Left            =   7590
      TabIndex        =   43
      Top             =   3390
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtSum6 
      Height          =   255
      Left            =   7590
      TabIndex        =   50
      Top             =   6300
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   4140
      TabIndex        =   104
      Top             =   6840
      Width           =   4695
      Begin MSMask.MaskEdBox txtAuthorizedAgent 
         Height          =   255
         Left            =   60
         TabIndex        =   52
         Top             =   390
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label45 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Present Employer / Authorized Agent"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   60
         TabIndex        =   105
         Top             =   150
         Width           =   3315
      End
   End
   Begin MSMask.MaskEdBox txtQDepBday1 
      Height          =   255
      Left            =   2820
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtQDepBday2 
      Height          =   255
      Left            =   2820
      TabIndex        =   12
      Top             =   3990
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtQDepBday3 
      Height          =   255
      Left            =   2820
      TabIndex        =   14
      Top             =   4260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtQDepBday4 
      Height          =   255
      Left            =   2820
      TabIndex        =   17
      Top             =   4530
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin VB.Label Label46 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable 13th Month Pay and Other Benefits"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4800
      TabIndex        =   107
      Top             =   3330
      Width           =   2505
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Housing Allowance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   106
      Top             =   2790
      Width           =   2505
   End
   Begin VB.Label Label24 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   80
      Top             =   480
      Width           =   3795
   End
   Begin VB.Label Label23 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Qualified Dependent Children"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   79
      Top             =   3450
      Width           =   3795
   End
   Begin VB.Label Label22 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Is the wife claiming the additional exemption for qualified children?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   78
      Top             =   2970
      Width           =   2955
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Exemption Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   2400
      Width           =   2505
   End
   Begin VB.Label Label20 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2190
      TabIndex        =   76
      Top             =   1860
      Width           =   1725
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   75
      Top             =   1860
      Width           =   1725
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   74
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   1320
      Width           =   2505
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Employees Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   72
      Top             =   780
      Width           =   2505
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Taxes Withheld"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4800
      TabIndex        =   71
      Top             =   6630
      Width           =   2715
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount of Taxes Withheld Present Employer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4800
      TabIndex        =   70
      Top             =   6150
      Width           =   2715
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount of Taxes Withheld Previous Employer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4800
      TabIndex        =   69
      Top             =   5700
      Width           =   2715
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Due"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4800
      TabIndex        =   68
      Top             =   5460
      Width           =   2715
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Less: Total Exemptions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4800
      TabIndex        =   67
      Top             =   5160
      Width           =   2715
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Add: Taxable Compensation from Previous Employer(s)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   4800
      TabIndex        =   66
      Top             =   4680
      Width           =   2715
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Compensation Income"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4800
      TabIndex        =   65
      Top             =   4440
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4170
      TabIndex        =   64
      Top             =   4140
      Width           =   4635
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "B. Taxable Compensation Income"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4170
      TabIndex        =   63
      Top             =   1920
      Width           =   4635
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "A. Non-Taxable/Exempt Compensation Income"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4170
      TabIndex        =   62
      Top             =   480
      Width           =   4635
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Taxable Compensation Income"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4800
      TabIndex        =   61
      Top             =   3720
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost of Living Allowance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   60
      Top             =   2490
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Non-Taxable/Exempt Compensation Income"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   4800
      TabIndex        =   59
      Top             =   1440
      Width           =   2505
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT YEAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   58
      Top             =   90
      Width           =   1485
   End
   Begin VB.Label Label58 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "13th Month Pay and Other Benefits"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   4800
      TabIndex        =   55
      Top             =   750
      Width           =   2505
   End
   Begin VB.Label Label53 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "SSS/PHIC/Pag-ibig"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   54
      Top             =   1170
      Width           =   2505
   End
   Begin VB.Label Label55 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Salary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   53
      Top             =   2190
      Width           =   2505
   End
   Begin VB.Label Label60 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   3090
      Width           =   2505
   End
End
Attribute VB_Name = "frmForm2316"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo                                                         As ADODB.Recordset
Dim rsYTDDETAILS                                                      As ADODB.Recordset
Dim rsDEPENDENTS                                                      As ADODB.Recordset
Dim RSPAYROLL                                                         As ADODB.Recordset
Dim rsSalaryGrade                                                     As ADODB.Recordset
Dim rsCommission                                                      As ADODB.Recordset
Dim EMPLIVIL                                                          As String

Function SetSalary(SalCode As String) As Double
    Set rsSalaryGrade = New ADODB.Recordset
    'rsSalaryGrade.Open "select code,salary from HRMS_SalaryGrade where code = '" & SalCode & "'", gconDMIS
    rsSalaryGrade.Open "SELECT CODE,SALARY FROM HRMS_SALARYGRADE WHERE CODE = '" & SalCode & "'", gconDMIS
    If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
        SetSalary = N2Str2Zero(rsSalaryGrade!SALARY)
    End If
End Function

Function SetDailyRate(SalCode As String) As Double
    Set rsSalaryGrade = New ADODB.Recordset
    'rsSalaryGrade.Open "select code,dailyrate from HRMS_SalaryGrade where code = '" & SalCode & "'", gconDMIS
    rsSalaryGrade.Open "SELECT CODE,DAILYRATE FROM HRMS_SALARYGRADE WHERE CODE = '" & SalCode & "'", gconDMIS
    If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
        SetDailyRate = N2Str2Zero(rsSalaryGrade!DailyRate)
    End If
End Function

Sub rsrefresh()
    'If EMPINFOSHOW = True Then
    '   Set rsEMPINFO = New ADODB.Recordset
    '       rsEMPINFO.Open "select * from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & IMPNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'ElseIf HEADEMPINFOSHOW = True Then
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & IMPNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'Else
    '   Set rsEMPINFO = New ADODB.Recordset
    '       rsEMPINFO.Open "select * from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " order by lastname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'End If
End Sub

Sub StoreMemVars()
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        txtEmployeeName.Text = Null2String(rsEmpInfo!lastname) & ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
        txtRegisteredAddress.Text = Null2String(rsEmpInfo!ADDRESS)
        txtDateOfBirth.Text = Null2String(rsEmpInfo!BIRTHDATE)
        txtTelephoneNumber.Text = Null2String(rsEmpInfo!TELEPHONE)
        txtExemptionStatus.Text = Null2String(rsEmpInfo!STATUS)
        Set rsYTDDETAILS = New ADODB.Recordset
        rsYTDDETAILS.Open "select * from HRMS_ytddetails where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboYear.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
            If Null2String(rsYTDDETAILS!IsTheWife) = "Y" Then
                optYesIstheWife.Value = True
            Else
                optNoIsTheWife.Value = True
            End If
            If Null2String(rsYTDDETAILS!IsMainSecond) = "Y" Then
                optMainEmployer.Value = True
            Else
                optSecondaryEmployer.Value = True
            End If
            txtQDependent1.Text = Null2String(rsYTDDETAILS!QDependent1)
            txtQDependent2.Text = Null2String(rsYTDDETAILS!QDependent2)
            txtQDependent3.Text = Null2String(rsYTDDETAILS!QDependent3)
            txtQDependent4.Text = Null2String(rsYTDDETAILS!QDependent4)
            txtQDepBday1.Text = Null2String(rsYTDDETAILS!QDepbday1)
            txtQDepBday2.Text = Null2String(rsYTDDETAILS!QDepbday2)
            txtQDepBday3.Text = Null2String(rsYTDDETAILS!QDepbday3)
            txtQDepBday4.Text = Null2String(rsYTDDETAILS!QDepbday4)

            txtNonTax1.Text = N2Str2Zero(rsYTDDETAILS!NonTax1)
            txtNonTax2.Text = N2Str2Zero(rsYTDDETAILS!NonTax2)

            txtTax1.Text = N2Str2Zero(rsYTDDETAILS!tax1)
            txtTax2.Text = N2Str2Zero(rsYTDDETAILS!Tax2)
            txtTax3.Text = N2Str2Zero(rsYTDDETAILS!tax3)
            txtTax4.Text = N2Str2Zero(rsYTDDETAILS!tax4)
            txtTax5.Text = N2Str2Zero(rsYTDDETAILS!tax5)

            txtSum1.Text = N2Str2Zero(rsYTDDETAILS!sum1)
            txtSum2.Text = N2Str2Zero(rsYTDDETAILS!sum2)
            txtSum3.Text = N2Str2Zero(rsYTDDETAILS!Sum3)
            txtSum4.Text = N2Str2Zero(rsYTDDETAILS!Sum4)
            txtSum5.Text = N2Str2Zero(rsYTDDETAILS!sum5)
            txtSum6.Text = N2Str2Zero(rsYTDDETAILS!sum6)
            txtAuthorizedAgent.Text = Null2String(rsYTDDETAILS!AuthorizedAgent)
        End If
    End If
End Sub

Sub InitMemvars()
    txtEmployeeName.Text = ""
    txtRegisteredAddress.Text = ""
    txtDateOfBirth.Text = ""
    txtTelephoneNumber.Text = ""
    txtExemptionStatus.Text = ""

    optYesIstheWife.Value = False
    optNoIsTheWife.Value = False

    txtQDependent1.Text = ""
    txtQDependent2.Text = ""
    txtQDependent3.Text = ""
    txtQDependent4.Text = ""

    txtPresentTIN.Text = COMPANY_TIN
    txtPresentEmployersName.Text = COMPANY_NAME
    txtPresentRegisteredAddress.Text = COMPANY_ADDRESS
    txtPresentZipCode.Text = "4400"

    optMainEmployer.Value = False
    optSecondaryEmployer.Value = False

    txtPrev1TIN.Text = ""
    txtPrev1EmployersName.Text = ""
    txtPrev1RegisteredAddress.Text = ""
    txtPrev1ZipCode.Text = ""

    txtPrev2TIN.Text = ""
    txtPrev2EmployersName.Text = ""
    txtPrev2RegisteredAddress.Text = ""
    txtPrev2ZipCode.Text = ""

    txtPrev3TIN.Text = ""
    txtPrev3EmployersName.Text = ""
    txtPrev3RegisteredAddress.Text = ""
    txtPrev3ZipCode.Text = ""

    txtNonTax1.Text = "0.00"
    txtNonTax2.Text = "0.00"
    txtNonTax3.Text = "0.00"

    txtTax1.Text = "0.00"
    txtTax2.Text = "0.00"
    txtTax3.Text = "0.00"
    txtTax4.Text = "0.00"

    txtSum1.Text = "0.00"
    txtSum2.Text = "0.00"
    txtSum3.Text = "0.00"
    txtSum4.Text = "0.00"
    txtSum5.Text = "0.00"
    txtSum6.Text = "0.00"
    txtAuthorizedAgent.Text = ""
End Sub

Sub RecomputeYTD()
    Dim VEMPNO, VTaxCode                                              As String
    Dim VYTDBasicPay, VYTDUTLate, VYTDAbsent                          As Double
    Dim VCommission, VCommissionTax, VDecCommissionTax                As Double
    Dim VOvertime, VTaxableAdj, VNonTaxableAdj                        As Double
    Dim Vsss, Vphic, Vpagibig, VCOLA, VYTDIncome                      As Double
    Dim VRemSal, VRemCOLA, VAccSalary, VRemOT                         As Double
    Dim VRemWTax, VCURRemSal, VCURRemOT                               As Double
    Dim VCURRemWTax, VMidYear, V13thMonth                             As Double
    Dim VPersonalEx, VYTDTaxable, VYTDNonTaxable                      As Double
    Dim VNetTaxable, VNetTax, VDecNetTax                              As Double
    Dim VTaxDue                                                       As Double

    Dim VTOTYTDBasicPay, VTOTCommission, VTOTCommissionTax            As Double
    Dim VDECTOTCommissionTax, VTOTOvertime, VTOTTaxableAdj            As Double
    Dim VTOTNonTaxableAdj, VTOTsss, VTOTphic                          As Double
    Dim VTOTpagibig, VTOTcola, VTOTYTDIncome, VTOTPersonalEx          As Double
    Dim VTOTYTDTaxable, VTOTYTDNonTaxable, VTOTNetTaxable             As Double
    Dim VTOTNetTax, VDECTOTNetTax, VTOTTaxDue, VBONUS                 As Double

    Dim VPAYempno, VPAYtaxcode                                        As String
    Dim VPAYrate, VPAYdailyrate, VPAYovertime                         As Double
    Dim VPAYcommission, VPAYcommissionTax, VDECPAYcommissionTax       As Double
    Dim VPAYtaxableadj, VPAYnontaxableadj, VPAYgross                  As Double
    Dim VPAYundertime, VPAYsss, VPAYphilhealth                        As Double
    Dim VPAYpagibig As Double, VPAYcola As Double, VPAYtin As Double, VDECPAYtin As Double
    Dim VPAYabsent                                                    As Double
    Dim VTOTPAYempno As String, VTOTPAYtaxcode                        As String
    Dim VTOTPAYrate As Double, VTOTPAYdailyrate As Double, VTOTPAYovertime As Double
    Dim VTOTPAYcommission As Double, VTOTPAYcommissionTax As Double, VDECTOTPAYcommissionTax As Double
    Dim VTOTPAYtaxableadj As Double, VTOTPAYnontaxableadj As Double, VTOTPAYgross As Double
    Dim VTOTPAYundertime As Double, VTOTPAYsss As Double, VTOTPAYphilhealth As Double
    Dim VTOTPAYpagibig As Double, VTOTPAYcola As Double, VTOTPAYtin As Double, VDECTOTPAYtin As Double

    Dim VTOTPAYabsent                                                 As Double

    Dim VNOVDECBASICPAY                                               As Double
    Dim VARYEER                                                       As String
    Dim NoMonths, manths, manths2                                     As Integer
    Dim VARYTDINCOME, VARPERSONALEX, BULANAN                          As Double

    Dim CutOffDate                                                    As String
    CutOffDate = ""
    Set rsEmpInfo = New ADODB.Recordset
    'rsEmpInfo.Open "select * from HRMS_EmpInfo where emplevel = " & EMPLIVIL & " and empno = '" & IMPNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & IMPNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        VARYEER = cboYear.Text
        VEMPNO = "": VTaxCode = "": NoMonths = 0
        VYTDBasicPay = 0: VCommission = 0: VOvertime = 0
        VTaxableAdj = 0: VNonTaxableAdj = 0: VCommissionTax = 0: VDecCommissionTax = 0
        VYTDIncome = 0: VYTDTaxable = 0: VYTDNonTaxable = 0
        VNetTaxable = 0: VNetTax = 0: VDecNetTax = 0: VAccSalary = 0
        VRemSal = 0: V13thMonth = 0: VRemOT = 0: VRemWTax = 0
        Vsss = 0: Vphic = 0: Vpagibig = 0: VCOLA = 0

        VPAYrate = 0: VPAYdailyrate = 0: VPAYcommissionTax = 0: VDECPAYcommissionTax = 0
        VPAYovertime = 0: VPAYcommission = 0: VPAYtaxableadj = 0: VDECPAYtin = 0: VPAYtin = 0
        VPAYnontaxableadj = 0: VPAYundertime = 0: VPAYsss = 0
        VPAYphilhealth = 0: VPAYpagibig = 0: VPAYcola = 0: VPAYabsent = 0

        VTOTPAYrate = 0: VTOTPAYdailyrate = 0: VTOTPAYtin = 0: VDECTOTPAYtin = 0
        VTOTPAYovertime = 0: VTOTPAYtaxableadj = 0: VMidYear = 0
        VTOTPAYnontaxableadj = 0: VTOTPAYundertime = 0: VTOTPAYsss = 0
        VTOTPAYphilhealth = 0: VTOTPAYpagibig = 0: VTOTPAYcola = 0: VTOTPAYabsent = 0
        VTOTPAYcommissionTax = 0: VDECTOTPAYcommissionTax = 0: VTOTPAYcommission = 0:

        VPAYempno = Null2String(rsEmpInfo!EMPNO)
        VNOVDECBASICPAY = 0
        Set RSPAYROLL = New ADODB.Recordset
        RSPAYROLL.Open "select * from HRMS_Payroll where (EMPLEVEL = " & N2Str2Null(rsEmpInfo!EMPLEVEL) & ") AND year(paydateto) = " & cboYear.Text & " AND empno =" & N2Str2Null(rsEmpInfo!EMPNO) & " order by paydateto desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPAYROLL.EOF And Not RSPAYROLL.BOF Then
            RSPAYROLL.MoveFirst
            Do While Not RSPAYROLL.EOF
                VPAYrate = 0: VPAYdailyrate = 0: VPAYtin = 0: VDECPAYtin = 0
                VPAYovertime = 0: VPAYtaxableadj = 0
                VPAYnontaxableadj = 0: VPAYundertime = 0: VPAYsss = 0
                VPAYphilhealth = 0: VPAYpagibig = 0: VPAYcola = 0: VPAYabsent = 0
                VPAYcommissionTax = 0: VDECPAYcommissionTax = 0: VPAYcommission = 0

                manths = MONTH(RSPAYROLL!paydatefrom)
                If manths2 <> manths Then
                    manths2 = manths
                    NoMonths = NoMonths + 1
                End If
                VPAYtaxcode = Null2String(rsEmpInfo!ExStatus)
                VPAYrate = N2Str2Zero(RSPAYROLL!Rate)
                VPAYdailyrate = N2Str2Zero(RSPAYROLL!DailyRate)
                VPAYovertime = NumericVal(N2Str2Zero(RSPAYROLL!OVERTIME)) + NumericVal(N2Str2Zero(RSPAYROLL!HOLIDAY))
                If MONTH(RSPAYROLL!paydateto) = 12 Then
                    VDECPAYtin = N2Str2Zero(RSPAYROLL!TAX)
                    VDECTOTPAYtin = VDECTOTPAYtin + VDECPAYtin
                Else
                    VPAYtin = N2Str2Zero(RSPAYROLL!TAX)
                    VTOTPAYtin = VTOTPAYtin + VPAYtin
                End If
                If MONTH(RSPAYROLL!paydateto) > 10 Then
                    VNOVDECBASICPAY = VNOVDECBASICPAY + VPAYrate
                End If
                VPAYtaxableadj = N2Str2Zero(RSPAYROLL!TAXABLEADJ)
                VPAYnontaxableadj = N2Str2Zero(RSPAYROLL!NONTAXABLEADJ)
                VPAYundertime = N2Str2Zero(RSPAYROLL!UNDERTIME)
                VPAYsss = N2Str2Zero(RSPAYROLL!SSSE)
                VPAYphilhealth = N2Str2Zero(RSPAYROLL!PHILHEALTHE)
                VPAYpagibig = N2Str2Zero(RSPAYROLL!PAGIBIG)
                VPAYcola = N2Str2Zero(RSPAYROLL!cola)
                VPAYabsent = N2Str2Zero(RSPAYROLL!ABSENT)

                VTOTPAYrate = VTOTPAYrate + VPAYrate
                VTOTPAYdailyrate = VTOTPAYdailyrate + VPAYdailyrate
                VTOTPAYovertime = VTOTPAYovertime + VPAYovertime

                VTOTPAYtaxableadj = VTOTPAYtaxableadj + VPAYtaxableadj
                VTOTPAYnontaxableadj = VTOTPAYnontaxableadj + VPAYnontaxableadj
                VTOTPAYundertime = VTOTPAYundertime + VPAYundertime
                VTOTPAYsss = VTOTPAYsss + VPAYsss
                VTOTPAYphilhealth = VTOTPAYphilhealth + VPAYphilhealth
                VTOTPAYpagibig = VTOTPAYpagibig + VPAYpagibig
                VTOTPAYcola = VTOTPAYcola + VPAYcola
                VTOTPAYabsent = VTOTPAYabsent + VPAYabsent
                RSPAYROLL.MoveNext
            Loop
        End If
        Set rsCommission = New ADODB.Recordset
        rsCommission.Open "select * from HRMS_Commission where (EMPLEVEL = " & N2Str2Null(rsEmpInfo!EMPLEVEL) & ") AND year(deyt) = " & cboYear.Text & " AND empno ='" & rsEmpInfo!EMPNO & "' order by deyt asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCommission.EOF And Not rsCommission.BOF Then
            rsCommission.MoveFirst
            VPAYcommissionTax = 0: VDECPAYcommissionTax = 0: VPAYcommission = 0
            Do While Not rsCommission.EOF
                VPAYcommission = N2Str2Zero(rsCommission!AMOUNT)
                If MONTH(rsCommission!DEYT) = 12 Then
                    VTOTPAYcommission = VTOTPAYcommission + VPAYcommission
                    VDECPAYcommissionTax = N2Str2Zero(rsCommission!TAX)
                    VDECTOTPAYcommissionTax = VDECTOTPAYcommissionTax + VDECPAYcommissionTax
                Else
                    VTOTPAYcommission = VTOTPAYcommission + VPAYcommission
                    VPAYcommissionTax = N2Str2Zero(rsCommission!TAX)
                    VTOTPAYcommissionTax = VTOTPAYcommissionTax + VPAYcommissionTax
                End If
                rsCommission.MoveNext
            Loop
        End If
        VEMPNO = VPAYempno
        VTaxCode = VPAYtaxcode
        VYTDBasicPay = (VTOTPAYrate) - (VTOTPAYundertime + VTOTPAYabsent)
        VYTDUTLate = VTOTPAYundertime
        VYTDAbsent = VTOTPAYabsent
        VCommission = VTOTPAYcommission
        VCommissionTax = VTOTPAYcommissionTax
        VDecCommissionTax = VDECTOTPAYcommissionTax
        VOvertime = VTOTPAYovertime
        VTaxableAdj = VTOTPAYtaxableadj
        VNonTaxableAdj = VTOTPAYnontaxableadj
        VARYTDINCOME = (VTOTPAYrate + VTOTPAYcola + VTaxableAdj + VCommission + VTOTPAYovertime) - (VTOTPAYundertime + VTOTPAYabsent)
        VARPERSONALEX = Personal_EX(VPAYtaxcode)
        Vsss = VTOTPAYsss
        Vphic = VTOTPAYphilhealth
        Vpagibig = VTOTPAYpagibig
        VCOLA = VTOTPAYcola

        If Null2String(rsEmpInfo!EMPSTATUS) = "M" Then
            BULANAN = SetSalary(Null2String(rsEmpInfo!SalaryCode))
        Else
            BULANAN = (SetDailyRate(Null2String(rsEmpInfo!SalaryCode)) * 314) / 12
        End If


        Set rsYTDDETAILS = New ADODB.Recordset
        rsYTDDETAILS.Open "select * from HRMS_ytddetails where (EMPLEVEL = " & N2Str2Null(rsEmpInfo!EMPLEVEL) & ") AND empno = '" & VEMPNO & "' and yeer = '" & VARYEER & "'", gconDMIS, adOpenKeyset, adLockOptimistic
        If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
            VMidYear = N2Str2Zero(rsYTDDETAILS!midyear)
            VCURRemSal = N2Str2Zero(rsYTDDETAILS!remsal)
            VCURRemOT = N2Str2Zero(rsYTDDETAILS!remot)
            VCURRemWTax = N2Str2Zero(rsYTDDETAILS!remwtax)
            VBONUS = N2Str2Zero(rsYTDDETAILS!bonus)
            VRemCOLA = N2Str2Zero(rsYTDDETAILS!RemCOLA)
            VRemSal = VCURRemSal
            VRemOT = VRemOT + (VCURRemOT - VRemOT)
            VRemWTax = VRemWTax + (VCURRemWTax - VRemWTax)
            If Null2String(rsEmpInfo!RESIGNED) = "" Then
                VAccSalary = 0
                V13thMonth = (((VYTDBasicPay) + VAccSalary)) / 12
                VBONUS = V13thMonth / 4
                V13thMonth = V13thMonth - VMidYear
            Else
                V13thMonth = (((VYTDBasicPay) + VRemSal)) / 12
            End If
            VMidYear = N2Str2Zero(rsYTDDETAILS!midyear)
            V13thMonth = N2Str2Zero(rsYTDDETAILS!t13thmonth)
            VBONUS = N2Str2Zero(rsYTDDETAILS!bonus)
            If V13thMonth + VMidYear + VBONUS > 30000 Then VARYTDINCOME = VARYTDINCOME + ((V13thMonth + VMidYear + VBONUS) - 30000)
            VYTDTaxable = (VARYTDINCOME + VRemCOLA + VRemOT + VRemSal)
            VYTDNonTaxable = VTOTPAYsss + VTOTPAYphilhealth + VTOTPAYpagibig
            VYTDBasicPay = (VYTDBasicPay + VOvertime) - VYTDNonTaxable
            VNetTaxable = VYTDTaxable - (VYTDNonTaxable + VARPERSONALEX)
            VNetTax = VTOTPAYtin + VRemWTax
            VDecNetTax = VDECTOTPAYtin
            VTaxDue = Tax_Due(VNetTaxable)
            VRemSal = VAccSalary
            gconDMIS.Execute "UPDATE HRMS_YTDDETAILS SET " & _
                             "TAXCODE = '" & VTaxCode & "', " & _
                             "YTDGROSS = " & VTOTPAYrate & ", YTDBASICPAY = " & VYTDBasicPay & ", " & _
                             "YTDUTLATE = " & VYTDUTLate & ", " & _
                             "YTDABSENT = " & VYTDAbsent & ", " & _
                             "COMMISSION = " & VCommission & ", " & _
                             "COMMISSIONTAX = " & VCommissionTax & ", DECCOMMISSIONTAX = " & VDecCommissionTax & ", " & _
                             "OVERTIME =" & VOvertime & ", " & _
                             "TAXABLEADJ = " & VTaxableAdj & ", " & _
                             "NONTAXABLEADJ = " & VNonTaxableAdj & ", " & _
                             "YTDSSS = " & Vsss & ", " & _
                             "YTDPHIC = " & Vphic & ", ytdpagibig = " & Vpagibig & ", ytdcola = " & VCOLA & ", " & _
                             "YTDINCOME = " & VARYTDINCOME + VRemSal + VRemOT & ", " & _
                             "PERSONALEX = " & VARPERSONALEX & ", " & _
                             "YTDTABLE = " & VYTDTaxable & ", " & _
                             "NONTAXABLE = " & VYTDNonTaxable & ", " & _
                             "NETTAXABLE = " & VNetTaxable & ", " & _
                             "YTDTAX = " & VNetTax & ", DECYTDTAX = " & VDecNetTax & ", " & _
                             "REMCOLA = " & VRemCOLA & ", REMSAL = " & VRemSal & ", REMOT = " & VRemOT & ", REMWTAX = " & VRemWTax & ", " & _
                             "DATEHIRED = " & N2Date2Null(rsEmpInfo!DateHired) & ", " & _
                             "YTDCUTOFFDATE = " & N2Date2Null(CutOffDate) & ", " & _
                             "YTDGENERATE = " & N2Date2Null(GENTO) & ", " & _
                             "TAXDUE = " & VTaxDue & _
                           " WHERE (EMPLEVEL = " & N2Str2Null(rsEmpInfo!EMPLEVEL) & ") AND EMPNO = '" & VEMPNO & "' AND YEER = '" & VARYEER & "'"
        End If
    End If
End Sub

Private Sub cmdSelectYear_Click()

    On Error GoTo Errorcode:

    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        txtEmployeeName.Text = Null2String(rsEmpInfo!lastname) & ", " & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME)
        txtRegisteredAddress.Text = Null2String(rsEmpInfo!ADDRESS)
        txtDateOfBirth.Text = Null2String(rsEmpInfo!BIRTHDATE)
        txtTelephoneNumber.Text = Null2String(rsEmpInfo!TELEPHONE)
        If Left(Null2String(rsEmpInfo!ExStatus), 2) = "ME" Then
            txtExemptionStatus.Text = "Married"
        ElseIf Left(Null2String(rsEmpInfo!ExStatus), 2) = "HF" Then
            txtExemptionStatus.Text = "Head of the Family"
        Else
            txtExemptionStatus.Text = "Single"
        End If
        Set rsDEPENDENTS = New ADODB.Recordset
        'Set rsDEPENDENTS = gconDMIS.Execute("select FullName,Birthday,Relation,TaxClaim,ID from Dependents where taxclaim = 'Y' and EMPLEVEL = " & EMPLIVIL & " AND empno = " & Null2String(rsEmpInfo!EMPNO))
        Set rsDEPENDENTS = gconDMIS.Execute("SELECT FULLNAME,BIRTHDAY,RELATION,TAXCLAIM,ID FROM HRMS_DEPENDENTS WHERE TAXCLAIM = 'Y' AND EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & Null2String(rsEmpInfo!EMPNO))
        If Not (rsDEPENDENTS.EOF And rsDEPENDENTS.BOF) Then
            Do While Not rsDEPENDENTS.EOF
                txtQDependent1.Text = Null2String(rsDEPENDENTS!FULLNAME)
                txtQDepBday1.Text = Null2String(rsDEPENDENTS!BIRTHDAY)
                rsDEPENDENTS.MoveNext
                If rsDEPENDENTS.EOF = True Then Exit Do
                txtQDependent2.Text = Null2String(rsDEPENDENTS!FULLNAME)
                txtQDepBday2.Text = Null2String(rsDEPENDENTS!BIRTHDAY)
                rsDEPENDENTS.MoveNext
                If rsDEPENDENTS.EOF = True Then Exit Do
                txtQDependent3.Text = Null2String(rsDEPENDENTS!FULLNAME)
                txtQDepBday3.Text = Null2String(rsDEPENDENTS!BIRTHDAY)
                rsDEPENDENTS.MoveNext
                If rsDEPENDENTS.EOF = True Then Exit Do
                txtQDependent4.Text = Null2String(rsDEPENDENTS!FULLNAME)
                txtQDepBday4.Text = Null2String(rsDEPENDENTS!BIRTHDAY)
                Exit Do
            Loop
        End If
        Set rsYTDDETAILS = New ADODB.Recordset
        'rsYTDDETAILS.Open "select * from HRMS_ytddetails where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboYear.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        rsYTDDETAILS.Open "SELECT * FROM HRMS_YTDDETAILS WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & rsEmpInfo!EMPNO & "' AND YEER = '" & cboYear.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
            'txtPresentTIN.Text = Null2String(rsYTDDETAILS!PresentTIN)
            'txtPresentEmployersName.Text = Null2String(rsYTDDETAILS!PresentEmployersName)
            'txtPresentRegisteredAddress.Text = Null2String(rsYTDDETAILS!PresentRegisteredAddress)
            'txtPresentZipCode.Text = Null2String(rsYTDDETAILS!PresentZipCode)
            RecomputeYTD
            optMainEmployer.Value = False
            optSecondaryEmployer.Value = False

            txtPrev1TIN.Text = Null2String(rsYTDDETAILS!Prev1TIN)
            txtPrev1EmployersName.Text = Null2String(rsYTDDETAILS!Prev1EmployersName)
            txtPrev1RegisteredAddress.Text = Null2String(rsYTDDETAILS!Prev1RegisteredAddress)
            txtPrev1ZipCode.Text = Null2String(rsYTDDETAILS!Prev1ZipCode)

            txtPrev2TIN.Text = Null2String(rsYTDDETAILS!Prev2TIN)
            txtPrev2EmployersName.Text = Null2String(rsYTDDETAILS!Prev2EmployersName)
            txtPrev2RegisteredAddress.Text = Null2String(rsYTDDETAILS!Prev2RegisteredAddress)
            txtPrev2ZipCode.Text = Null2String(rsYTDDETAILS!Prev2ZipCode)

            txtPrev3TIN.Text = Null2String(rsYTDDETAILS!Prev3TIN)
            txtPrev3EmployersName.Text = Null2String(rsYTDDETAILS!Prev3EmployersName)
            txtPrev3RegisteredAddress.Text = Null2String(rsYTDDETAILS!Prev3RegisteredAddress)
            txtPrev3ZipCode.Text = Null2String(rsYTDDETAILS!Prev3ZipCode)

            If N2Str2Zero(rsYTDDETAILS!t13thmonth) + N2Str2Zero(rsYTDDETAILS!midyear) + N2Str2Zero(rsYTDDETAILS!bonus) > 30000 Then
                txtNonTax1.Text = 30000
                txtTax5.Text = (N2Str2Zero(rsYTDDETAILS!t13thmonth) + N2Str2Zero(rsYTDDETAILS!midyear) + N2Str2Zero(rsYTDDETAILS!bonus)) - 30000
            Else
                txtNonTax1.Text = N2Str2Zero(rsYTDDETAILS!t13thmonth) + N2Str2Zero(rsYTDDETAILS!midyear) + N2Str2Zero(rsYTDDETAILS!bonus)
                txtTax5.Text = "0.00"
            End If
            txtNonTax2.Text = N2Str2Zero(rsYTDDETAILS!NONTAXABLE)
            txtNonTax3.Text = NumericVal(txtNonTax2.Text) + NumericVal(txtNonTax1.Text)

            txtTax1.Text = N2Str2Zero(rsYTDDETAILS!ytdbasicpay) + N2Str2Zero(rsYTDDETAILS!TAXABLEADJ)
            txtTax2.Text = N2Str2Zero(rsYTDDETAILS!ytdcola)
            txtTax3.Text = N2Str2Zero(rsYTDDETAILS!tax3)
            txtTax4.Text = N2Str2Zero(rsYTDDETAILS!commission)
            txtTax6.Text = NumericVal(txtTax1.Text) + NumericVal(txtTax2.Text) + NumericVal(txtTax3.Text) + NumericVal(txtTax4.Text) + NumericVal(txtTax5.Text)

            txtSum1.Text = NumericVal(txtTax6.Text)
            txtSum2.Text = N2Str2Zero(rsYTDDETAILS!sum2)
            txtSum3.Text = N2Str2Zero(rsYTDDETAILS!PersonalEx)
            txtSum4.Text = N2Str2Zero(rsYTDDETAILS!Taxdue)
            txtSum5.Text = N2Str2Zero(rsYTDDETAILS!sum5)
            txtSum6.Text = N2Str2Zero(rsYTDDETAILS!ytdtax) + N2Str2Zero(rsYTDDETAILS!commissiontax) + N2Str2Zero(rsYTDDETAILS!decytdtax) + N2Str2Zero(rsYTDDETAILS!deccommissiontax)
            txtSum7.Text = NumericVal(txtSum5.Text) + NumericVal(txtSum6.Text)
        End If
    End If

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    If MsgBox("Save this data?", vbYesNo + vbQuestion, "Update System") = vbYes Then
        Dim vIsMainSecond, vIsTheWife                                 As String
        If optMainEmployer.Value = True Then vIsMainSecond = "'Y'" Else vIsMainSecond = "'N'"
        If optYesIstheWife.Value = True Then vIsTheWife = "'Y'" Else vIsTheWife = "'N'"
        '        gconDMIS.Execute "UPDATE HRMS_YTDDETAILS SET " & _
                 '                         "QDependent1 = " & N2Str2Null(txtQDependent1.Text) & ", QDependent2 = " & N2Str2Null(txtQDependent2.Text) & "," & _
                 '                         "QDependent3 = " & N2Str2Null(txtQDependent3.Text) & ", QDependent4 = " & N2Str2Null(txtQDependent4.Text) & "," & _
                 '                         "QDepBday1 = " & N2Str2Null(txtQDepBday1.Text) & ", QDepBday2 = " & N2Str2Null(txtQDepBday2.Text) & "," & _
                 '                         "QDepBday3 = " & N2Str2Null(txtQDepBday3.Text) & ", QDepBday4 = " & N2Str2Null(txtQDepBday4.Text) & "," & _
                 '                         "IsMainSecond = " & vIsMainSecond & ", IsTheWife = " & vIsTheWife & "," & _
                 '                         "PresentTIN = " & N2Str2Null(txtPresentTIN.Text) & ", PresentEmployersName = " & N2Str2Null(txtPresentEmployersName.Text) & "," & _
                 '                         "PresentRegisteredAddress = " & N2Str2Null(txtPresentRegisteredAddress.Text) & ", PresentZipCode = " & N2Str2Null(txtPresentZipCode.Text) & "," & _
                 '                         "Prev1TIN = " & N2Str2Null(txtPrev1TIN.Text) & ", Prev1EmployersName = " & N2Str2Null(txtPrev1EmployersName.Text) & "," & _
                 '                         "Prev1RegisteredAddress = " & N2Str2Null(txtPrev1RegisteredAddress.Text) & ", Prev1ZipCode = " & N2Str2Null(txtPrev1ZipCode.Text) & "," & _
                 '                         "Prev2TIN = " & N2Str2Null(txtPrev2TIN.Text) & ", Prev2EmployersName = " & N2Str2Null(txtPrev2EmployersName.Text) & "," & _
                 '                         "Prev2RegisteredAddress = " & N2Str2Null(txtPrev2RegisteredAddress.Text) & ", Prev2ZipCode = " & N2Str2Null(txtPrev2ZipCode.Text) & "," & _
                 '                         "Prev3TIN = " & N2Str2Null(txtPrev3TIN.Text) & ", Prev3EmployersName = " & N2Str2Null(txtPrev3EmployersName.Text) & "," & _
                 '                         "Prev3RegisteredAddress = " & N2Str2Null(txtPrev3RegisteredAddress.Text) & ", Prev3ZipCode = " & N2Str2Null(txtPrev3ZipCode.Text) & "," & _
                 '                         "NonTax1 = " & NumericVal(txtNonTax1.Text) & ", NonTax2 = " & NumericVal(txtNonTax2.Text) & ", NonTax3 = " & NumericVal(txtNonTax3.Text) & "," & _
                 '                         "Tax1 = " & NumericVal(txtTax1.Text) & ", Tax2 = " & NumericVal(txtTax2.Text) & ", Tax3 = " & NumericVal(txtTax3.Text) & ", Tax4 = " & NumericVal(txtTax4.Text) & ", Tax5 = " & NumericVal(txtTax5.Text) & ", Tax6 = " & NumericVal(txtTax6.Text) & "," & _
                 '                         "Sum1 = " & NumericVal(txtSum1.Text) & ", Sum2 = " & NumericVal(txtSum2.Text) & ", Sum3 = " & NumericVal(txtSum3.Text) & ", Sum4 = " & NumericVal(txtSum4.Text) & ", Sum5 = " & NumericVal(txtSum5.Text) & ", Sum6 = " & NumericVal(txtSum6.Text) & ", Sum7 = " & NumericVal(txtSum7.Text) & "," & _
                 '                         "AuthorizedAgent = " & N2Str2Null(txtAuthorizedAgent.Text) & _
                 '                       " WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND YEER = '" & cboYear.Text & "'"
        '
        gconDMIS.Execute "UPDATE HRMS_YTDDETAILS SET " & _
                         "QDEPENDENT1 = " & N2Str2Null(txtQDependent1.Text) & ", QDEPENDENT2 = " & N2Str2Null(txtQDependent2.Text) & "," & _
                         "QDEPENDENT3 = " & N2Str2Null(txtQDependent3.Text) & ", QDEPENDENT4 = " & N2Str2Null(txtQDependent4.Text) & "," & _
                         "QDEPBDAY1 = " & N2Str2Null(txtQDepBday1.Text) & ", QDEPBDAY2 = " & N2Str2Null(txtQDepBday2.Text) & "," & _
                         "QDEPBDAY3 = " & N2Str2Null(txtQDepBday3.Text) & ", QDEPBDAY4 = " & N2Str2Null(txtQDepBday4.Text) & "," & _
                         "ISMAINSECOND = " & vIsMainSecond & ", ISTHEWIFE = " & vIsTheWife & "," & _
                         "PRESENTTIN = " & N2Str2Null(txtPresentTIN.Text) & ", PRESENTEMPLOYERSNAME = " & N2Str2Null(txtPresentEmployersName.Text) & "," & _
                         "PRESENTREGISTEREDADDRESS = " & N2Str2Null(txtPresentRegisteredAddress.Text) & ", PRESENTZIPCODE = " & N2Str2Null(txtPresentZipCode.Text) & "," & _
                         "PREV1TIN = " & N2Str2Null(txtPrev1TIN.Text) & ", PREV1EMPLOYERSNAME = " & N2Str2Null(txtPrev1EmployersName.Text) & "," & _
                         "PREV1REGISTEREDADDRESS = " & N2Str2Null(txtPrev1RegisteredAddress.Text) & ", PREV1ZIPCODE = " & N2Str2Null(txtPrev1ZipCode.Text) & "," & _
                         "PREV2TIN = " & N2Str2Null(txtPrev2TIN.Text) & ", PREV2EMPLOYERSNAME = " & N2Str2Null(txtPrev2EmployersName.Text) & "," & _
                         "PREV2REGISTEREDADDRESS = " & N2Str2Null(txtPrev2RegisteredAddress.Text) & ", PREV2ZIPCODE = " & N2Str2Null(txtPrev2ZipCode.Text) & "," & _
                         "PREV3TIN = " & N2Str2Null(txtPrev3TIN.Text) & ", PREV3EMPLOYERSNAME = " & N2Str2Null(txtPrev3EmployersName.Text) & "," & _
                         "PREV3REGISTEREDADDRESS = " & N2Str2Null(txtPrev3RegisteredAddress.Text) & ", PREV3ZIPCODE = " & N2Str2Null(txtPrev3ZipCode.Text) & "," & _
                         "NONTAX1 = " & NumericVal(txtNonTax1.Text) & ", NONTAX2 = " & NumericVal(txtNonTax2.Text) & ", NONTAX3 = " & NumericVal(txtNonTax3.Text) & "," & _
                         "TAX1 = " & NumericVal(txtTax1.Text) & ", TAX2 = " & NumericVal(txtTax2.Text) & ", TAX3 = " & NumericVal(txtTax3.Text) & ", TAX4 = " & NumericVal(txtTax4.Text) & ", TAX5 = " & NumericVal(txtTax5.Text) & ", TAX6 = " & NumericVal(txtTax6.Text) & "," & _
                         "SUM1 = " & NumericVal(txtSum1.Text) & ", SUM2 = " & NumericVal(txtSum2.Text) & ", SUM3 = " & NumericVal(txtSum3.Text) & ", SUM4 = " & NumericVal(txtSum4.Text) & ", SUM5 = " & NumericVal(txtSum5.Text) & ", SUM6 = " & NumericVal(txtSum6.Text) & ", SUM7 = " & NumericVal(txtSum7.Text) & "," & _
                         "AUTHORIZEDAGENT = " & N2Str2Null(txtAuthorizedAgent.Text) & _
                       " WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND YEER = '" & cboYear.Text & "'"
    End If
    Screen.MousePointer = 11
    PrintSQLReport rptForm2316, HRMS_REPORT_PATH & "Form2316.rpt", "{EmpInfo.EmpNo} = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND {EmpInfo.EmpLEVEL} = " & EMPLIVIL & " AND {YTDDetails.YEER} = '" & cboYear.Text & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CenterMe Me, Me, 0
    InitMemvars
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    cboYear.Text = frmHRMSLedger.cboYear.Text
    rsrefresh
    StoreMemVars
End Sub

Private Sub txtNonTax1_Change()
    txtNonTax3.Text = ToDoubleNumber(NumericVal(txtNonTax1.Text) + NumericVal(txtNonTax2.Text))
End Sub

Private Sub txtNonTax2_Change()
    txtNonTax3.Text = ToDoubleNumber(NumericVal(txtNonTax1.Text) + NumericVal(txtNonTax2.Text))
End Sub

Private Sub txtSum2_Change()
    txtSum4.Text = Tax_Due((NumericVal(txtSum1.Text) + NumericVal(txtSum2.Text)) - NumericVal(txtSum3.Text))
End Sub

Private Sub txtSum5_Change()
    txtSum7.Text = NumericVal(txtSum5.Text) + NumericVal(txtSum6.Text)
End Sub

Private Sub txtSum6_Change()
    txtSum7.Text = NumericVal(txtSum5.Text) + NumericVal(txtSum6.Text)
End Sub

Private Sub txtTax1_Change()
    txtTax6.Text = ToDoubleNumber(NumericVal(txtTax1.Text) + NumericVal(txtTax2.Text) + NumericVal(txtTax3.Text) + NumericVal(txtTax4.Text) + NumericVal(txtTax5.Text))
End Sub

Private Sub txtTax2_Change()
    txtTax6.Text = ToDoubleNumber(NumericVal(txtTax1.Text) + NumericVal(txtTax2.Text) + NumericVal(txtTax3.Text) + NumericVal(txtTax4.Text) + NumericVal(txtTax5.Text))
End Sub

Private Sub txtTax3_Change()
    txtTax6.Text = ToDoubleNumber(NumericVal(txtTax1.Text) + NumericVal(txtTax2.Text) + NumericVal(txtTax3.Text) + NumericVal(txtTax4.Text) + NumericVal(txtTax5.Text))
End Sub

Private Sub txtTax4_Change()
    txtTax6.Text = ToDoubleNumber(NumericVal(txtTax1.Text) + NumericVal(txtTax2.Text) + NumericVal(txtTax3.Text) + NumericVal(txtTax4.Text) + NumericVal(txtTax5.Text))
End Sub

Private Sub txtTax5_Change()
    txtTax6.Text = ToDoubleNumber(NumericVal(txtTax1.Text) + NumericVal(txtTax2.Text) + NumericVal(txtTax3.Text) + NumericVal(txtTax4.Text) + NumericVal(txtTax5.Text))
End Sub

Private Sub txtTax6_Change()
    txtSum1.Text = NumericVal(txtTax6.Text)
    txtSum4.Text = Tax_Due((NumericVal(txtSum1.Text) + NumericVal(txtSum2.Text)) - NumericVal(txtSum3.Text))
End Sub


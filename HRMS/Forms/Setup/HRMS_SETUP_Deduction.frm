VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSETUP_Deduction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deduction Setup"
   ClientHeight    =   5895
   ClientLeft      =   315
   ClientTop       =   600
   ClientWidth     =   11220
   Icon            =   "HRMS_SETUP_Deduction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   11220
   Begin VB.PictureBox picSave 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   5130
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   58
      Top             =   4920
      Width           =   1440
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
         Left            =   720
         MouseIcon       =   "HRMS_SETUP_Deduction.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "HRMS_SETUP_Deduction.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
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
         Height          =   795
         Left            =   30
         MouseIcon       =   "HRMS_SETUP_Deduction.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "HRMS_SETUP_Deduction.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame6 
      Enabled         =   0   'False
      Height          =   1845
      Left            =   240
      TabIndex        =   68
      Top             =   7590
      Width           =   3255
      Begin VB.OptionButton chkLoan1 
         Caption         =   "1st Cut-Off Pay Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   450
         TabIndex        =   71
         Top             =   840
         Width           =   2595
      End
      Begin VB.OptionButton chkLoan2 
         Caption         =   "2nd Cut-Off Pay Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   450
         TabIndex        =   70
         Top             =   1140
         Width           =   2595
      End
      Begin VB.OptionButton chkLoan3 
         Caption         =   "1st and 2nd Cut-Off Pay Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   450
         TabIndex        =   69
         Top             =   1440
         Width           =   2595
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   180
         Picture         =   "HRMS_SETUP_Deduction.frx":11FC
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label5 
         Caption         =   "Loans Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   990
         TabIndex        =   72
         Top             =   330
         Width           =   2085
      End
   End
   Begin VB.Frame Frame7 
      Enabled         =   0   'False
      Height          =   1845
      Left            =   3570
      TabIndex        =   63
      Top             =   7590
      Width           =   3255
      Begin VB.OptionButton chkOther1 
         Caption         =   "1st Cut-Off Pay Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   66
         Top             =   810
         Width           =   2595
      End
      Begin VB.OptionButton chkOther2 
         Caption         =   "2nd Cut-Off Pay Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   65
         Top             =   1110
         Width           =   2595
      End
      Begin VB.OptionButton chkOther3 
         Caption         =   "1st and 2nd Cut-Off Pay Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   64
         Top             =   1410
         Width           =   2595
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   240
         Picture         =   "HRMS_SETUP_Deduction.frx":1506
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label16 
         Caption         =   "Other Deductions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   990
         TabIndex        =   67
         Top             =   330
         Width           =   2085
      End
   End
   Begin VB.PictureBox picAdd 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   3420
      ScaleHeight     =   915
      ScaleWidth      =   3240
      TabIndex        =   57
      Top             =   4920
      Width           =   3240
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
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
         Left            =   2430
         MouseIcon       =   "HRMS_SETUP_Deduction.frx":1810
         MousePointer    =   99  'Custom
         Picture         =   "HRMS_SETUP_Deduction.frx":1962
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
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
         Height          =   795
         Left            =   1740
         MouseIcon       =   "HRMS_SETUP_Deduction.frx":1CC8
         MousePointer    =   99  'Custom
         Picture         =   "HRMS_SETUP_Deduction.frx":1E1A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   1050
         MouseIcon       =   "HRMS_SETUP_Deduction.frx":2145
         MousePointer    =   99  'Custom
         Picture         =   "HRMS_SETUP_Deduction.frx":2297
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   360
         MouseIcon       =   "HRMS_SETUP_Deduction.frx":25F3
         MousePointer    =   99  'Custom
         Picture         =   "HRMS_SETUP_Deduction.frx":2745
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   11235
      TabIndex        =   32
      Top             =   60
      Width           =   11235
      Begin VB.TextBox txtNewSet 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   5280
         TabIndex        =   62
         Text            =   "1"
         Top             =   60
         Width           =   855
      End
      Begin VB.ComboBox cboSet 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   60
         Width           =   840
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1845
         Left            =   90
         TabIndex        =   34
         Top             =   870
         Width           =   3255
         Begin VB.OptionButton chkSSS3 
            Caption         =   "1st and 2nd Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   2
            Top             =   1470
            Width           =   2595
         End
         Begin VB.OptionButton chkSSS2 
            Caption         =   "2nd Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   1
            Top             =   1170
            Width           =   2595
         End
         Begin VB.OptionButton chkSSS1 
            Caption         =   "1st Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   0
            Top             =   870
            Width           =   2595
         End
         Begin VB.Label Label1 
            Caption         =   "Social Security System"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   990
            TabIndex        =   39
            Top             =   330
            Width           =   2085
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   90
            Picture         =   "HRMS_SETUP_Deduction.frx":2A58
            Top             =   210
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   2115
         Left            =   120
         TabIndex        =   33
         Top             =   2700
         Width           =   3255
         Begin VB.CheckBox Check1 
            Caption         =   "Use User-defined Values"
            Height          =   255
            Left            =   510
            TabIndex        =   73
            Top             =   1770
            Width           =   2475
         End
         Begin VB.OptionButton chkPagIbig3 
            Caption         =   "1st and 2nd Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   510
            TabIndex        =   5
            Top             =   1470
            Width           =   2595
         End
         Begin VB.OptionButton chkPagIbig2 
            Caption         =   "2nd Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   510
            TabIndex        =   4
            Top             =   1170
            Width           =   2595
         End
         Begin VB.OptionButton chkPagIbig1 
            Caption         =   "1st Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   510
            TabIndex        =   3
            Top             =   870
            Width           =   2595
         End
         Begin VB.Label Label2 
            Caption         =   "Pag-Ibig Fund"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   990
            TabIndex        =   40
            Top             =   360
            Width           =   2085
         End
         Begin VB.Image Image2 
            Height          =   615
            Left            =   60
            Picture         =   "HRMS_SETUP_Deduction.frx":458A
            Stretch         =   -1  'True
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   2115
         Left            =   3450
         TabIndex        =   35
         Top             =   2700
         Width           =   3255
         Begin VB.OptionButton chkTax4 
            Caption         =   "Non Taxable"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   540
            TabIndex        =   74
            Top             =   1800
            Width           =   2595
         End
         Begin VB.OptionButton chkTax3 
            Caption         =   "1st and 2nd Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   540
            TabIndex        =   11
            Top             =   1500
            Width           =   2595
         End
         Begin VB.OptionButton chkTax2 
            Caption         =   "2nd Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   540
            TabIndex        =   10
            Top             =   1200
            Width           =   2595
         End
         Begin VB.OptionButton chkTax1 
            Caption         =   "1st Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   540
            TabIndex        =   9
            Top             =   900
            Width           =   2595
         End
         Begin VB.Label Label4 
            Caption         =   "Withholding Tax"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1020
            TabIndex        =   42
            Top             =   360
            Width           =   2085
         End
         Begin VB.Image Image4 
            Height          =   615
            Left            =   90
            Picture         =   "HRMS_SETUP_Deduction.frx":8FA8
            Stretch         =   -1  'True
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1845
         Left            =   3450
         TabIndex        =   37
         Top             =   870
         Width           =   3255
         Begin VB.OptionButton chkPhilH3 
            Caption         =   "1st and 2nd Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   1440
            Width           =   2595
         End
         Begin VB.OptionButton chkPhilH2 
            Caption         =   "2nd Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   7
            Top             =   1140
            Width           =   2595
         End
         Begin VB.OptionButton chkPhilH1 
            Caption         =   "1st Cut-Off Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   840
            Width           =   2595
         End
         Begin VB.Label Label3 
            Caption         =   "Philippine Health"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   930
            TabIndex        =   41
            Top             =   330
            Width           =   2085
         End
         Begin VB.Image Image3 
            Height          =   615
            Left            =   90
            Picture         =   "HRMS_SETUP_Deduction.frx":B4B2
            Stretch         =   -1  'True
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         Height          =   5745
         Left            =   6750
         TabIndex        =   36
         Top             =   -60
         Width           =   4365
         Begin VB.ComboBox cboTaxComp 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2970
            Width           =   2235
         End
         Begin VB.ComboBox cboTAXBasis 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1290
            Width           =   2235
         End
         Begin VB.ComboBox cboPHICBasis 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   900
            Width           =   2235
         End
         Begin VB.ComboBox cboPIFBasis 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   510
            Width           =   2235
         End
         Begin VB.ComboBox cboSSSBasis 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   120
            Width           =   2235
         End
         Begin VB.TextBox txtHours 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   2790
            MaxLength       =   3
            TabIndex        =   24
            Top             =   4980
            Width           =   615
         End
         Begin VB.TextBox txtAverage 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   2790
            MaxLength       =   3
            TabIndex        =   23
            Top             =   4590
            Width           =   615
         End
         Begin VB.TextBox txtComp 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   20
            Top             =   2970
            Width           =   315
         End
         Begin VB.TextBox txtDays 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   2790
            MaxLength       =   3
            TabIndex        =   22
            Top             =   4200
            Width           =   615
         End
         Begin VB.Frame Frame9 
            Height          =   735
            Left            =   90
            TabIndex        =   53
            Top             =   3300
            Width           =   4125
            Begin VB.Label Label18 
               Caption         =   "[1] Monthly Net Taxable"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   55
               Top             =   180
               Width           =   2475
            End
            Begin VB.Label Label17 
               Caption         =   "[2] Annualized Computation"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   54
               Top             =   420
               Width           =   2475
            End
         End
         Begin VB.TextBox txtTax 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   18
            Top             =   1290
            Width           =   315
         End
         Begin VB.TextBox txtPhilH 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   16
            Top             =   900
            Width           =   315
         End
         Begin VB.TextBox txtPagI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   14
            Top             =   510
            Width           =   315
         End
         Begin VB.TextBox txtSSS 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   12
            Top             =   120
            Width           =   315
         End
         Begin VB.Frame Frame8 
            Height          =   1215
            Left            =   90
            TabIndex        =   47
            Top             =   1650
            Width           =   4125
            Begin VB.Label Label13 
               Caption         =   "[4] for Basic plus OT less Lates and Absences"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   900
               Width           =   3795
            End
            Begin VB.Label Label12 
               Caption         =   "[3] for Basic less Lates and Absences"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   660
               Width           =   3285
            End
            Begin VB.Label Label11 
               Caption         =   "[2] for Basic"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   49
               Top             =   420
               Width           =   1515
            End
            Begin VB.Label Label10 
               Caption         =   "[1] for Gross"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   48
               Top             =   180
               Width           =   1515
            End
         End
         Begin VB.Label Label20 
            Caption         =   "Working Hours per Day"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   60
            Top             =   5100
            Width           =   2745
         End
         Begin VB.Label Label19 
            Caption         =   "Average No. of Days in a Month"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   59
            Top             =   4710
            Width           =   2745
         End
         Begin VB.Label Label15 
            Caption         =   "Working Days in a year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   56
            Top             =   4320
            Width           =   1995
         End
         Begin VB.Label Label14 
            Caption         =   "Tax Computed By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   150
            TabIndex        =   52
            Top             =   3030
            Width           =   1995
         End
         Begin VB.Label Label9 
            Caption         =   "Tax Basis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   150
            TabIndex        =   46
            Top             =   1290
            Width           =   1515
         End
         Begin VB.Label Label8 
            Caption         =   "Philhealth Basis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   150
            TabIndex        =   45
            Top             =   900
            Width           =   1515
         End
         Begin VB.Label Label7 
            Caption         =   "Pag-Ibig Basis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   150
            TabIndex        =   44
            Top             =   540
            Width           =   1515
         End
         Begin VB.Label Label6 
            Caption         =   "SSS Basis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   150
            TabIndex        =   43
            Top             =   210
            Width           =   1515
         End
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   930
         TabIndex        =   75
         Top             =   5070
         Visible         =   0   'False
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   51707905
         CurrentDate     =   40452
      End
      Begin MSComCtl2.DTPicker dtptoDate 
         Height          =   315
         Left            =   900
         TabIndex        =   76
         Top             =   5400
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   51707905
         CurrentDate     =   40452
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -210
         TabIndex        =   79
         Top             =   5430
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -150
         TabIndex        =   78
         Top             =   5100
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label22 
         Caption         =   "Deduction Date  Range"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   77
         Top             =   4860
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label labDeductionSet 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   90
         TabIndex        =   61
         Top             =   540
         Width           =   6615
      End
      Begin VB.Label Label73 
         Caption         =   "Employee Deduction Set"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1140
         TabIndex        =   38
         Top             =   120
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmSETUP_Deduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADD_EDIT                                                          As String
Dim rsSETUP                                                           As ADODB.Recordset

Function GenerateDeductionSetup() As Integer
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DEDUCTION_SET FROM HRMS_SETUPDEDUCTION ORDER BY DEDUCTION_SET DESC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GenerateDeductionSetup = NumericVal(RSTMP!Deduction_set) + 1
    Else
        GenerateDeductionSetup = 1
    End If
    Set RSTMP = Nothing
End Function

Function SetLevel(XXX As Integer) As String
    If XXX = 1 Then SetLevel = "FIRST"
    If XXX = 2 Then SetLevel = "SECOND"
    If XXX = 3 Then SetLevel = "THIRD"
    If XXX = 4 Then SetLevel = "FOURTH"
    If XXX = 5 Then SetLevel = "FIFTH"
    If XXX = 6 Then SetLevel = "SIXTH"
    If XXX = 7 Then SetLevel = "SEVENTH"
    If XXX = 8 Then SetLevel = "EIGHT"
    If XXX = 9 Then SetLevel = "NINTH"
    If XXX = 10 Then SetLevel = "TENTH"
End Function

Sub ChangePic(COND As Boolean)
    If COND = True Then
        picAdd.Visible = False
        picSave.Visible = True
        cboSet.Enabled = False
        Frame1.Enabled = True
        Frame3.Enabled = True
        Frame5.Enabled = True
        Frame2.Enabled = True
        Frame4.Enabled = True
        Frame6.Enabled = True
        Frame7.Enabled = True
    Else
        picAdd.Visible = True
        picSave.Visible = False
        cboSet.Enabled = True
        Frame1.Enabled = False
        Frame3.Enabled = False
        Frame5.Enabled = False
        Frame2.Enabled = False
        Frame4.Enabled = False
        Frame6.Enabled = False
        Frame7.Enabled = False
    End If
End Sub

Sub InitMemvars()
    chkTax1.Value = False
    chkTax2.Value = False
    chkTax3.Value = False

    chkLoan1.Value = False
    chkLoan2.Value = False
    chkLoan3.Value = False

    chkPagIbig1.Value = False
    chkPagIbig2.Value = False
    chkPagIbig3.Value = False

    chkOther1.Value = False
    chkOther2.Value = False
    chkOther3.Value = False

    chkPhilH1.Value = False
    chkPhilH2.Value = False
    chkPhilH3.Value = False

    chkSSS2.Value = False
    chkSSS2.Value = False
    chkSSS3.Value = False

    txtSSS.Text = ""
    txtPagI.Text = ""
    txtPhilH.Text = ""
    txtTax.Text = ""
    txtDays.Text = ""
    txtComp.Text = ""
    txtHours.Text = ""
    txtAverage.Text = ""
    InitComboDed
End Sub

Sub InitComboDed()
    cboSSSBasis.Clear
    cboSSSBasis.AddItem "[1] for Gross"
    cboSSSBasis.AddItem "[2] for Basic"
    cboSSSBasis.AddItem "[3] for Basic less Lates and Absences"
    cboSSSBasis.AddItem "[4] for Basic plus OT less Lates and Absences"
    cboSSSBasis.ListIndex = -1

    cboPIFBasis.Clear
    cboPIFBasis.AddItem "[1] for Gross"
    cboPIFBasis.AddItem "[2] for Basic"
    cboPIFBasis.AddItem "[3] for Basic less Lates and Absences"
    cboPIFBasis.AddItem "[4] for Basic plus OT less Lates and Absences"
    cboPIFBasis.ListIndex = -1

    cboPHICBasis.Clear
    cboPHICBasis.AddItem "[1] for Gross"
    cboPHICBasis.AddItem "[2] for Basic"
    cboPHICBasis.AddItem "[3] for Basic less Lates and Absences"
    cboPHICBasis.AddItem "[4] for Basic plus OT less Lates and Absences"
    cboPHICBasis.ListIndex = -1

    cboTAXBasis.Clear
    cboTAXBasis.AddItem "[1] for Gross"
    cboTAXBasis.AddItem "[2] for Basic"
    cboTAXBasis.AddItem "[3] for Basic less Lates and Absences"
    cboTAXBasis.AddItem "[4] for Basic plus OT less Lates and Absences"
    cboTAXBasis.ListIndex = -1

    cboTaxComp.Clear
    cboTaxComp.AddItem "[1] Monthly Net Taxable"
    cboTaxComp.AddItem "[2] Annualized Computation"
    cboTaxComp.ListIndex = -1
    'DoEvents
End Sub

Sub SetComboDed(Combo As ComboBox, KIM As Integer)
    If KIM = 1 Then Combo.Text = "[1] for Gross"
    If KIM = 2 Then Combo.Text = "[2] for Basic"
    If KIM = 3 Then Combo.Text = "[3] for Basic less Lates and Absences"
    If KIM = 4 Then Combo.Text = "[4] for Basic plus OT less Lates and Absences"
End Sub

Sub SetTextDed(Text As TextBox, KIM As String)
    If KIM = "[1] for Gross" Then Text.Text = 1
    If KIM = "[2] for Basic" Then Text.Text = 2
    If KIM = "[3] for Basic less Lates and Absences" Then Text.Text = 3
    If KIM = "[4] for Basic plus OT less Lates and Absences" Then Text.Text = 4
End Sub

Sub FillCombo()
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT adj_from,adj_to,DEDUCTION_SET FROM HRMS_SETUPDEDUCTION ORDER BY DEDUCTION_SET")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        cboSet.AddItem Null2String(RSTMP!Deduction_set)
        cboSet.Text = Null2String(RSTMP!Deduction_set)
        
        dtpFromDate.Value = Null2String(RSTMP!adj_from)
        dtpToDate.Value = Null2String(RSTMP!adj_to)
        
        RSTMP.MoveNext
        Do While Not RSTMP.EOF
            cboSet.AddItem Null2String(RSTMP!Deduction_set)
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub StoreMemVars(XXX As Integer)
    Set rsSETUP = New ADODB.Recordset
    rsSETUP.Open "SELECT * FROM HRMS_SETUPDEDUCTION WHERE DEDUCTION_SET = " & XXX & " ORDER BY DEDUCTION_SET", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not (rsSETUP.BOF And rsSETUP.EOF) Then
        If cboSet.Text <> XXX Then cboSet.Text = XXX
        labDeductionSet.Caption = " " & SetLevel(XXX) & " DEDUCTION SET OF PAYROLL"
        If NumericVal(rsSETUP!SSS) = 1 Then chkSSS1.Value = True:
        If NumericVal(rsSETUP!SSS) = 2 Then chkSSS2.Value = True:
        If NumericVal(rsSETUP!SSS) = 3 Then chkSSS3.Value = True:

        If NumericVal(rsSETUP!PAGIBIG) = 1 Then chkPagIbig1.Value = True:
        If NumericVal(rsSETUP!PAGIBIG) = 2 Then chkPagIbig2.Value = True:
        If NumericVal(rsSETUP!PAGIBIG) = 3 Then chkPagIbig3.Value = True:

        If NumericVal(rsSETUP!PHILHEALTH) = 1 Then chkPhilH1.Value = True:
        If NumericVal(rsSETUP!PHILHEALTH) = 2 Then chkPhilH2.Value = True:
        If NumericVal(rsSETUP!PHILHEALTH) = 3 Then chkPhilH3.Value = True:

        If NumericVal(rsSETUP!TAX) = 1 Then chkTax1.Value = True:
        If NumericVal(rsSETUP!TAX) = 2 Then chkTax2.Value = True:
        If NumericVal(rsSETUP!TAX) = 3 Then chkTax3.Value = True:
        If NumericVal(rsSETUP!TAX) = 4 Then chkTax4.Value = True:

        If NumericVal(rsSETUP!LOAN) = 1 Then chkLoan1.Value = True:
        If NumericVal(rsSETUP!LOAN) = 2 Then chkLoan2.Value = True:
        If NumericVal(rsSETUP!LOAN) = 3 Then chkLoan3.Value = True:

        If NumericVal(rsSETUP!Others) = 1 Then chkOther1.Value = True:
        If NumericVal(rsSETUP!Others) = 2 Then chkOther2.Value = True:
        If NumericVal(rsSETUP!Others) = 3 Then chkOther3.Value = True:

        txtSSS.Text = NumericVal(rsSETUP!SSS_BASIS)
        txtPagI.Text = NumericVal(rsSETUP!PAGIBIG_BASIS)
        txtPhilH.Text = NumericVal(rsSETUP!PHILHEALTH_BASIS)
        txtTax.Text = NumericVal(rsSETUP!TAX_BASIS)

        SetComboDed cboSSSBasis, txtSSS
        SetComboDed cboPIFBasis, txtPagI
        SetComboDed cboPHICBasis, txtPhilH
        SetComboDed cboTAXBasis, txtTax

        txtComp.Text = NumericVal(rsSETUP!TAX_COMPUTED)
        If txtComp.Text = 1 Then
            cboTaxComp.Text = "[1] Monthly Net Taxable"
        Else
            cboTaxComp.Text = "[2] Annualized Computation"
        End If
        txtDays.Text = NumericVal(rsSETUP!WORKING_DAY)
        txtAverage.Text = NumericVal(rsSETUP!AVERAGE_MONTH)
        txtHours.Text = NumericVal(rsSETUP!WORKING_HOURS)

        If N2Str2Zero(rsSETUP!PAGIBIG_SET) = 1 Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If

    Else
        ShowNoRecord
        Call cmdAdd_Click
    End If
End Sub

Private Sub cboPHICBasis_Click()
    SetTextDed txtPhilH, cboPHICBasis
End Sub

Private Sub cboPIFBasis_Click()
    SetTextDed txtPagI, cboPIFBasis
End Sub

Private Sub cboSet_Change()
    StoreMemVars cboSet
End Sub

Private Sub cboSet_Click()
    StoreMemVars cboSet
End Sub

Private Sub cboSet_LostFocus()
    StoreMemVars cboSet
End Sub

Private Sub cboSSSBasis_Click()
    SetTextDed txtSSS, cboSSSBasis
End Sub

Private Sub cboTAXBasis_Click()
    SetTextDed txtTax, cboTAXBasis
End Sub

Private Sub cboTaxComp_Click()
    If cboTaxComp.Text = "[1] Monthly Net Taxable" Then
        txtComp.Text = 1
    Else
        txtComp.Text = 2
    End If
End Sub

Private Sub cmdAdd_Click()
    ADD_EDIT = "ADD"
    Call ChangePic(True)

    txtNewSet.Visible = True
    txtNewSet.Text = GenerateDeductionSetup
    InitMemvars
End Sub

Private Sub cmdCancel_Click()
    ADD_EDIT = ""
    txtNewSet.Text = ""
    txtNewSet.Visible = False
    Call ChangePic(False)
    StoreMemVars NumericVal(cboSet)
End Sub

Private Sub cmdDelete_Click()
    Dim RSTMP                                                         As New ADODB.Recordset

    If Not cboSet.Text = "" Then
        Set RSTMP = gconDMIS.Execute("Select PayrollGroup From HRMS_EmpInfo Where PayrollGroup = " & cboSet.Text & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            MsgBox "Deduction Seup # " & cboSet.Text & " cannot be deleted, setp is used in the system.", vbInformation, "HRMS"
        Else
            If MsgBox("delete this Deduction Setup", vbQuestion + vbYesNo, "Are you Sure") = vbYes Then
                gconDMIS.Execute ("Delete from HRMS_SETUPDEDUCTION Where Deduction_SET = '" & cboSet.Text & "'")

                Call FillCombo
                StoreMemVars cboSet
            End If
        End If

        Set RSTMP = Nothing
    End If
    LogAudit "X", "DELETE SETUP DEDUCTION", cboSet.Text
End Sub

Private Sub cmdEdit_Click()
    If Not cboSet.Text = "" Then
        Call ChangePic(True)
        ADD_EDIT = "EDIT"
        chkSSS1.SetFocus
    Else
        MsgBox "Choose Deduction Setup to Edit", vbInformation, "HRMS"
        cboSet.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim Vsss                                                          As Integer
    Dim vPAGI                                                         As Integer
    Dim vPHIL                                                         As Integer
    Dim vTAX                                                          As Integer
    Dim vLOAN                                                         As Integer
    Dim vOTHER                                                        As Integer
    Dim vSSSB                                                         As Integer
    Dim vTAXB                                                         As Integer
    Dim vPHILB                                                        As Integer
    Dim vPAGIB                                                        As Integer
    Dim vCOMP                                                         As Integer
    Dim vDAYS                                                         As Integer
    Dim vHOURS                                                        As Integer
    Dim vAVEMONTHS                                                    As Integer

    Dim vPAGIBSET                                                     As Integer

    Dim ActiveSet                                                     As Integer

    If txtSSS.Text = "" Then
        ShowIsRequiredMsg ("SSS Basis Field cannot be Blank")
        txtSSS.SetFocus
        Exit Sub
    End If
    If txtTax.Text = "" Then
        ShowIsRequiredMsg ("Tax Basis Field cannot be Blank")
        txtTax.SetFocus
        Exit Sub
    End If
    If txtPhilH.Text = "" Then
        ShowIsRequiredMsg ("PhilHealth Basis Field cannot be Blank")
        txtPhilH.SetFocus
        Exit Sub
    End If
    If txtPagI.Text = "" Then
        ShowIsRequiredMsg ("Pagibig Basis Field cannot be Blank")
        txtPagI.SetFocus
        Exit Sub
    End If
    If txtComp.Text = "" Then
        ShowIsRequiredMsg ("Tax COmputed By Field cannot be Blank")
        txtComp.SetFocus
        Exit Sub
    End If
    If txtDays.Text = "" Then
        ShowIsRequiredMsg ("Working Days Field cannot be Blank")
        txtDays.SetFocus
        Exit Sub
    End If


    If chkSSS1.Value = True Then Vsss = 1
    If chkSSS2.Value = True Then Vsss = 2
    If chkSSS3.Value = True Then Vsss = 3

    If chkPagIbig1.Value = True Then vPAGI = 1
    If chkPagIbig2.Value = True Then vPAGI = 2
    If chkPagIbig3.Value = True Then vPAGI = 3

    If chkPhilH1.Value = True Then vPHIL = 1
    If chkPhilH2.Value = True Then vPHIL = 2
    If chkPhilH3.Value = True Then vPHIL = 3

    If chkTax1.Value = True Then vTAX = 1
    If chkTax2.Value = True Then vTAX = 2
    If chkTax3.Value = True Then vTAX = 3
    If chkTax4.Value = True Then vTAX = 4

    If chkLoan1.Value = True Then vLOAN = 1
    If chkLoan2.Value = True Then vLOAN = 2
    If chkLoan3.Value = True Then vLOAN = 3

    If chkOther1.Value = True Then vOTHER = 1
    If chkOther2.Value = True Then vOTHER = 2
    If chkOther3.Value = True Then vOTHER = 3

    If Check1.Value = 1 Then
        vPAGIBSET = 1
    Else
        vPAGIBSET = 0
    End If

    vTAXB = NumericVal(txtTax.Text)
    vSSSB = NumericVal(txtSSS.Text)
    vPAGIB = NumericVal(txtPagI.Text)
    vPHILB = NumericVal(txtPhilH.Text)
    vDAYS = NumericVal(txtDays.Text)
    vCOMP = NumericVal(txtComp.Text)
    vHOURS = NumericVal(txtHours.Text)
    vAVEMONTHS = NumericVal(txtAverage.Text)

    If ADD_EDIT = "ADD" Then
        ActiveSet = NumericVal(txtNewSet.Text)
        gconDMIS.Execute ("INSERT INTO HRMS_SETUPDEDUCTION (DEDUCTION_SET,SSS,PAGIBIG,PHILHEALTH,TAX,LOAN,OTHERS,SSS_BASIS,PAGIBIG_BASIS,PHILHEALTH_BASIS,TAX_BASIS,TAX_COMPUTED,WORKING_DAY,WORKING_HOURS,AVERAGE_MONTH, PAGIBIG_SET) " & _
            " VALUES(" & ActiveSet & "," & Vsss & "," & vPAGI & "," & vPHIL & _
            "," & vTAX & "," & vLOAN & "," & vOTHER & "," & vSSSB & _
            "," & vPAGIB & "," & vPAGIB & "," & vTAXB & "," & vCOMP & _
            "," & vDAYS & "," & vHOURS & "," & vAVEMONTHS & "," & vPAGIBSET & ")")
        
        Call ShowSuccessFullyAdded
        LogAudit "A", "ADD SETUP DEDUCTION", cboSet.Text
    Else
        ActiveSet = NumericVal(cboSet.Text)
        
        '**** update kang my date range pa *******

'        gconDMIS.Execute ("UPDATE HRMS_SETUPDEDUCTION SET " & _
'            " SSS = " & Vsss & _
'            ", PAGIBIG = " & vPAGI & _
'            ", PHILHEALTH = " & vPHIL & _
'            ", TAX = " & vTAX & _
'            ", LOAN = " & vLOAN & _
'            ", OTHERS = " & vOTHER & _
'            ", SSS_BASIS = " & vSSSB & _
'            ", PAGIBIG_BASIS = " & vPAGIB & _
'            ", PHILHEALTH_BASIS = " & vPHILB & _
'            ", TAX_BASIS = " & vTAXB & _
'            ", TAX_COMPUTED = " & vCOMP & _
'            ", WORKING_DAY = " & vDAYS & _
'            ", WORKING_HOURS = " & vHOURS & _
'            ", AVERAGE_MONTH = " & vAVEMONTHS & _
'            ", PAGIBIG_SET = " & vPAGIBSET & _
'            ", ADJ_FROM = " & N2Str2Null(dtpFromDate) & _
'            ", ADJ_TO = " & N2Str2Null(dtpToDate) & _
'            " WHERE DEDUCTION_SET = " & NumericVal(cboSet.Text))
        
        '******************************************
            
            
            
             gconDMIS.Execute ("UPDATE HRMS_SETUPDEDUCTION SET " & _
            " SSS = " & Vsss & _
            ", PAGIBIG = " & vPAGI & _
            ", PHILHEALTH = " & vPHIL & _
            ", TAX = " & vTAX & _
            ", LOAN = " & vLOAN & _
            ", OTHERS = " & vOTHER & _
            ", SSS_BASIS = " & vSSSB & _
            ", PAGIBIG_BASIS = " & vPAGIB & _
            ", PHILHEALTH_BASIS = " & vPHILB & _
            ", TAX_BASIS = " & vTAXB & _
            ", TAX_COMPUTED = " & vCOMP & _
            ", WORKING_DAY = " & vDAYS & _
            ", WORKING_HOURS = " & vHOURS & _
            ", AVERAGE_MONTH = " & vAVEMONTHS & _
            ", PAGIBIG_SET = " & vPAGIBSET & _
            " WHERE DEDUCTION_SET = " & NumericVal(cboSet.Text))
        
        Call ShowSuccessFullyUpdated
        LogAudit "E", "EDIT SETUP DEDUCTION", cboSet.Text
    End If
    
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select Deduction_Set from HRMS_setupdeduction Order By Deduction_Set")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveFirst: cboSet.Clear
        Do While Not RSTMP.EOF
            cboSet.AddItem Null2String(RSTMP!Deduction_set)
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
    
    cboSet.Text = ActiveSet
    'StoreMemVars ActiveSet
    Call cmdCancel_Click
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemvars
    Call FillCombo
    StoreMemVars NumericVal(cboSet.Text)
    cmdCancel.Value = True
End Sub


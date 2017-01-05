VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmHRMS_ATM_Summary 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATM Summary"
   ClientHeight    =   7860
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "ATM_Summary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   10920
   Visible         =   0   'False
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5D8BC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   90
      ScaleHeight     =   975
      ScaleWidth      =   5115
      TabIndex        =   76
      Top             =   1020
      Width           =   5145
      Begin VB.TextBox txtBATCHNO 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   3510
         TabIndex        =   81
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
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
         Left            =   4290
         MouseIcon       =   "ATM_Summary.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   -375
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1740
         Picture         =   "ATM_Summary.frx":09D6
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   77
         Top             =   -660
         Width           =   435
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   0
            MaxLength       =   3
            TabIndex        =   78
            Top             =   60
            Width           =   525
         End
      End
      Begin MSComCtl2.DTPicker txtCREDITDATE 
         Height          =   345
         Left            =   1080
         TabIndex        =   82
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         Format          =   54329345
         CurrentDate     =   39683
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   84
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Batch No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2640
         TabIndex        =   83
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5D8BC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   90
      ScaleHeight     =   555
      ScaleWidth      =   10755
      TabIndex        =   58
      Top             =   6330
      Width           =   10785
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   1350
         TabIndex        =   65
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   90
         Width           =   2385
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   4620
         TabIndex        =   63
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   90
         Width           =   2385
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   8280
         TabIndex        =   62
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   90
         Width           =   2385
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1740
         Picture         =   "ATM_Summary.frx":3712
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   60
         Top             =   -660
         Width           =   435
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   0
            MaxLength       =   3
            TabIndex        =   61
            Top             =   60
            Width           =   525
         End
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   -375
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Pay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   3840
         TabIndex        =   67
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Acct. No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   66
         Top             =   150
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Hash Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   7080
         TabIndex        =   64
         Top             =   180
         Width           =   1140
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   6405
      Left            =   90
      ScaleHeight     =   6405
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   -30
      Width           =   10875
      Begin FlexCell.Grid Grid1 
         Height          =   4335
         Left            =   0
         TabIndex        =   87
         Top             =   2040
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7646
         BackColor2      =   12907725
         BackColorBkg    =   -2147483645
         Cols            =   5
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         GridColor       =   12632256
         Rows            =   30
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5D8BC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   5100
         ScaleHeight     =   975
         ScaleWidth      =   5670
         TabIndex        =   51
         Top             =   1050
         Width           =   5700
         Begin VB.ComboBox Combo5 
            Height          =   345
            Left            =   3780
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   450
            Width           =   975
         End
         Begin VB.ComboBox Combo4 
            Height          =   345
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   450
            Width           =   1245
         End
         Begin VB.ComboBox Combo2 
            Height          =   345
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   450
            Width           =   1485
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Import"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   4860
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "ATM_Summary.frx":644E
            MousePointer    =   99  'Custom
            Picture         =   "ATM_Summary.frx":65A0
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "View Options"
            Top             =   60
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   5550
            TabIndex        =   55
            Text            =   "cboChargeTo"
            ToolTipText     =   "Select option from list."
            Top             =   -405
            Width           =   1785
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00DEDFDE&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   1740
            Picture         =   "ATM_Summary.frx":6995
            ScaleHeight     =   405
            ScaleWidth      =   435
            TabIndex        =   53
            Top             =   -660
            Width           =   435
            Begin VB.TextBox Text10 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   0
               MaxLength       =   3
               TabIndex        =   54
               Top             =   60
               Width           =   525
            End
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5640
            MaxLength       =   3
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   -375
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   8
            Left            =   3930
            TabIndex        =   75
            Top             =   90
            Width           =   390
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Month"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   7
            Left            =   2760
            TabIndex        =   74
            Top             =   90
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Cut-Off"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   5
            Left            =   1320
            TabIndex        =   73
            Top             =   90
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Period"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   69
            Top             =   510
            Width           =   915
         End
      End
      Begin VB.PictureBox picHdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5D8BC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   10785
         TabIndex        =   1
         Top             =   60
         Width           =   10815
         Begin VB.TextBox txtPRESENTINGOFFICE 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   360
            Left            =   1710
            TabIndex        =   49
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   525
            Width           =   2385
         End
         Begin VB.TextBox txtCOMPANYCODE 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   360
            Left            =   1710
            TabIndex        =   46
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   120
            Width           =   2385
         End
         Begin VB.TextBox txtCEILINGAMOUNT 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   360
            Left            =   7350
            TabIndex        =   45
            ToolTipText     =   "Input customer code (e.g. S01163)"
            Top             =   510
            Width           =   2355
         End
         Begin VB.TextBox txtACCOUNTNO 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   360
            Left            =   7350
            TabIndex        =   44
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   90
            Width           =   2355
         End
         Begin VB.TextBox txtChargeTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5640
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   -375
            Width           =   495
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00DEDFDE&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   1740
            Picture         =   "ATM_Summary.frx":96D1
            ScaleHeight     =   405
            ScaleWidth      =   435
            TabIndex        =   2
            Top             =   -660
            Width           =   435
            Begin VB.TextBox txtTranType 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   0
               MaxLength       =   3
               TabIndex        =   3
               Top             =   60
               Width           =   525
            End
         End
         Begin VB.ComboBox cboChargeTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   5550
            TabIndex        =   4
            Text            =   "cboChargeTo"
            ToolTipText     =   "Select option from list."
            Top             =   -405
            Width           =   1785
         End
         Begin VB.CommandButton Command2 
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
            Height          =   825
            Left            =   9960
            MouseIcon       =   "ATM_Summary.frx":C40D
            MousePointer    =   99  'Custom
            Picture         =   "ATM_Summary.frx":C55F
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Edit Selected Record"
            Top             =   90
            Width           =   735
         End
         Begin VB.CommandButton Command3 
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
            Height          =   825
            Left            =   9960
            MouseIcon       =   "ATM_Summary.frx":C8BB
            MousePointer    =   99  'Custom
            Picture         =   "ATM_Summary.frx":CA0D
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Save this Record"
            Top             =   90
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Presenting Office"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   585
            Width           =   1485
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   17
            Left            =   330
            TabIndex        =   48
            Top             =   150
            Width           =   1275
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Ceiling Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   5970
            TabIndex        =   47
            Top             =   570
            Width           =   1275
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Acct. No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   5640
            TabIndex        =   6
            Top             =   150
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox fraAddTran 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   2400
      ScaleHeight     =   4095
      ScaleWidth      =   6855
      TabIndex        =   15
      Top             =   1620
      Visible         =   0   'False
      Width           =   6885
      Begin VB.TextBox txtTranSONO 
         Height          =   330
         Left            =   120
         MaxLength       =   12
         TabIndex        =   43
         Top             =   1050
         Width           =   3645
      End
      Begin VB.CommandButton Command1 
         Caption         =   "::"
         Height          =   315
         Left            =   2550
         TabIndex        =   42
         Top             =   1680
         Width           =   315
      End
      Begin VB.TextBox txtPartID 
         BackColor       =   &H00FF0000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2760
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   30
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.CommandButton cmdTranDelete 
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
         Left            =   6000
         MouseIcon       =   "ATM_Summary.frx":CD5D
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":CEAF
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Delete Entry"
         Top             =   3060
         Width           =   735
      End
      Begin VB.ComboBox cboTranPartNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         ItemData        =   "ATM_Summary.frx":D1DA
         Left            =   120
         List            =   "ATM_Summary.frx":D1DC
         Sorted          =   -1  'True
         TabIndex        =   28
         Text            =   "Combo1"
         ToolTipText     =   "Select Part Number from the list."
         Top             =   1650
         Width           =   2415
      End
      Begin VB.TextBox txtTranItemNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   27
         ToolTipText     =   "Type item number (e.g. 0001)"
         Top             =   540
         Width           =   1125
      End
      Begin VB.TextBox txtTranUPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   25
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   2880
         Width           =   1995
      End
      Begin VB.TextBox txtTranTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   150
         MaxLength       =   10
         TabIndex        =   24
         Top             =   3510
         Width           =   3645
      End
      Begin VB.TextBox txtTranDescription 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2250
         Width           =   3675
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Parts Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   1815
         Left            =   3900
         TabIndex        =   17
         Top             =   570
         Width           =   2865
         Begin VB.CheckBox chkAvailableOnStock 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Available on Stock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   120
            TabIndex        =   20
            Top             =   270
            Width           =   2595
         End
         Begin VB.TextBox txtCurrentOnhand 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   405
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "0"
            Top             =   690
            Width           =   855
         End
         Begin VB.TextBox txtCurrentAllocated 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   405
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0"
            Top             =   1170
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Current Onhand"
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
            Index           =   2
            Left            =   150
            TabIndex        =   22
            Top             =   750
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Total Allocated"
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
            Index           =   1
            Left            =   150
            TabIndex        =   21
            Top             =   1260
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboDNPRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         ItemData        =   "ATM_Summary.frx":D1DE
         Left            =   2880
         List            =   "ATM_Summary.frx":D1E0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Select Part Number from the list."
         Top             =   1650
         Width           =   915
      End
      Begin VB.TextBox txtTranQty 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         MaxLength       =   10
         TabIndex        =   26
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   2880
         Width           =   1605
      End
      Begin VB.CommandButton cmdTranCancel 
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
         Left            =   5280
         MouseIcon       =   "ATM_Summary.frx":D1E2
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":D334
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Cancel Entry"
         Top             =   3060
         Width           =   735
      End
      Begin VB.CommandButton cmdTranSave 
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
         Left            =   4560
         MouseIcon       =   "ATM_Summary.frx":D672
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":D7C4
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Save Entry"
         Top             =   3060
         Width           =   735
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   6885
         _Version        =   655364
         _ExtentX        =   12144
         _ExtentY        =   661
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "DNP Rate"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   2880
         TabIndex        =   40
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   36
         Top             =   2010
         Width           =   1275
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Line Item No."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1410
         TabIndex        =   35
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   2610
         Width           =   660
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "DNP"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1830
         TabIndex        =   32
         Top             =   2640
         Width           =   390
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Line Amount"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   31
         Top             =   3270
         Width           =   1035
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SO NO"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   810
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport rptHash 
      Left            =   10950
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Parts Issuance"
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
   Begin VB.PictureBox picAdd 
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
      Height          =   870
      Left            =   720
      ScaleHeight     =   870
      ScaleWidth      =   10230
      TabIndex        =   7
      Top             =   6990
      Width           =   10230
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
         Height          =   825
         Left            =   9390
         MouseIcon       =   "ATM_Summary.frx":DB14
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":DC66
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command6 
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
         Height          =   825
         Left            =   8670
         MouseIcon       =   "ATM_Summary.frx":DFCC
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":E11E
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   7950
         MouseIcon       =   "ATM_Summary.frx":E46E
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":E5C0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   7230
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "ATM_Summary.frx":E926
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":EA78
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Unpost this Transaction"
         Top             =   0
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4950
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6510
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "ATM_Summary.frx":EDBD
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":EF0F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Create Disk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   5580
         MouseIcon       =   "ATM_Summary.frx":F234
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":F386
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.PictureBox picSaves 
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
      Height          =   885
      Left            =   7470
      ScaleHeight     =   885
      ScaleWidth      =   3390
      TabIndex        =   12
      Top             =   7020
      Width           =   3390
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
         Height          =   825
         Left            =   2640
         MouseIcon       =   "ATM_Summary.frx":F6D6
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":F828
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   735
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
         Height          =   825
         Left            =   1920
         MouseIcon       =   "ATM_Summary.frx":FB66
         MousePointer    =   99  'Custom
         Picture         =   "ATM_Summary.frx":FCB8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmHRMS_ATM_Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GETHASHVALUE(acctno As String, net As Double) As Double
    On Error Resume Next
    Dim xAccountNo                                                    As String
    xAccountNo = Repleys(acctno)

    Dim firstvalue
    Dim secondvalue
    Dim thirdvalue
    Dim firstvalue_sal
    Dim secondvalue_sal
    Dim thirdvalue_sal

    Dim X                                                             As Integer
    Dim count                                                         As Integer
    count = 0
    Dim matt(15)                                                      As Integer

    For X = 1 To Len(xAccountNo)
        If IsNumeric(Mid(xAccountNo, X, 1)) Then
            count = count + 1
            matt(count) = Mid(xAccountNo, X, 1)
        End If
    Next

    firstvalue = CInt(CStr(matt(5)) & CStr(matt(6)))
    secondvalue = CInt(CStr(matt(7)) & CStr(matt(8)))
    thirdvalue = CInt(CStr(matt(9)) & CStr(matt(10)))

    firstvalue_sal = NumericVal(firstvalue) * net
    secondvalue_sal = NumericVal(secondvalue) * net
    thirdvalue_sal = NumericVal(thirdvalue) * net

    GETHASHVALUE = Round((firstvalue_sal + secondvalue_sal + thirdvalue_sal), 2)

End Function

Function GET_EMP_NAME(EMPNO As String) As String
    GET_EMP_NAME = ""
    Dim rsTemp                                                        As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME AS EMPNAME FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_EMP_NAME = Null2String(rsTemp!EMPNAME)
    Else
        GET_EMP_NAME = ""
    End If
    Set rsTemp = Nothing
End Function

Function GET_EMP_ACCTNO(EMPNO As String) As String
    GET_EMP_ACCTNO = ""
    Dim rsTemp                                                        As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT ACCOUNTNO FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_EMP_ACCTNO = Null2String(rsTemp!ACCOUNTNO)
    End If
    Set rsTemp = Nothing
End Function

Function GET_NUMBER(ACCOUNTNO As String) As Double
    Dim X                                                             As Integer
    Dim count                                                         As Integer
    count = 0
    Dim matt(15)                                                      As String
    Dim AMOUNTSTRING                                                  As String
    AMOUNTSTRING = ""
    For X = 1 To Len(ACCOUNTNO)
        If IsNumeric(Mid(ACCOUNTNO, X, 1)) Then
            count = count + 1
            matt(count) = Mid(ACCOUNTNO, X, 1)
            AMOUNTSTRING = AMOUNTSTRING & matt(count)
        End If
    Next
    GET_NUMBER = Round(NumericVal(AMOUNTSTRING), 2)
End Function

Function GET_STRING(ACCOUNTNO As String) As String
    Dim X                                                             As Integer
    Dim AMOUNTSTRING                                                  As String
    AMOUNTSTRING = ""
    For X = 1 To Len(ACCOUNTNO)
        If IsNumeric(Mid(ACCOUNTNO, X, 1)) Then
            AMOUNTSTRING = AMOUNTSTRING & Mid(ACCOUNTNO, X, 1)
        End If
    Next
    GET_STRING = CStr(AMOUNTSTRING)
End Function

Sub InitGrid()
    With Grid1
        .Enabled = True
        .Cols = 7
        .Cell(0, 0).Text = "L/N"
        .Cell(0, 1).Text = "OPTION"
        .Cell(0, 2).Text = "EMPLOYEE NO"
        .Cell(0, 3).Text = "EMPLOYEE NAME"
        .Cell(0, 4).Text = "ACCOUNT NO"
        .Cell(0, 5).Text = "NET AMOUNT"
        .Cell(0, 6).Text = "HORIZONTAL HASH"
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = False
        .Column(6).Locked = True
        .Column(0).Width = 25
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 195
        .Column(4).Width = 120
        .Column(5).Width = 122
        .Column(6).Width = 110
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = CellLeft
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellRightGeneral
        .Column(6).Alignment = cellRightGeneral
        .Column(5).DecimalLength = 2
        .Column(5).Mask = cellValue
        .Column(6).DecimalLength = 2
        .Column(6).Mask = cellValue
        .Enabled = False
    End With
End Sub

Sub enabledisable(COND As Boolean)
    With Grid1
        .Column(4).Locked = COND
        .Column(6).Locked = COND
    End With
End Sub

Sub Fill_Header()
    Dim rsHeader                                                      As ADODB.Recordset
    Set rsHeader = New ADODB.Recordset
    Set rsHeader = gconDMIS.Execute("SELECT * FROM HRMS_HOR_HASH_HEADER")
    If Not rsHeader.EOF And Not rsHeader.BOF Then
        txtCOMPANYCODE = Null2String(rsHeader!COMPANY_CODE)
        txtPRESENTINGOFFICE = Null2String(rsHeader!PRESENTING_OFFICE)
        txtACCOUNTNO = Null2String(rsHeader!ACCOUNT_NO)
        txtCEILINGAMOUNT = Format(N2Str2Zero(rsHeader!CEILING_AMOUNT), "#,###,##0.00")
    Else
        txtCOMPANYCODE = ""
        txtPRESENTINGOFFICE = ""
        txtACCOUNTNO = ""
        txtCEILINGAMOUNT = "0.00"
    End If
End Sub

Sub Fill_Pay_Period()
    Dim CUT_OFF_PERIOD                                                As String
    Dim rsTemp                                                        As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT DISTINCT CUT_OFF FROM HRMS_PAYROLL WHERE ISNULL(CUT_OFF,'') <> '' ORDER BY CUT_OFF DESC")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            If Null2String(rsTemp!CUT_OFF) = "1" Then
                CUT_OFF_PERIOD = "1st Cut-Off"
            ElseIf Null2String(rsTemp!CUT_OFF) = "2" Then
                CUT_OFF_PERIOD = "2nd Cut-Off"
            End If
            Combo2.AddItem CUT_OFF_PERIOD
            rsTemp.MoveNext
        Wend
    End If
    Set rsTemp = Nothing

    Combo4.AddItem MonthName(1)
    Combo4.AddItem MonthName(2)
    Combo4.AddItem MonthName(3)
    Combo4.AddItem MonthName(4)
    Combo4.AddItem MonthName(5)
    Combo4.AddItem MonthName(6)
    Combo4.AddItem MonthName(7)
    Combo4.AddItem MonthName(8)
    Combo4.AddItem MonthName(9)
    Combo4.AddItem MonthName(10)
    Combo4.AddItem MonthName(11)
    Combo4.AddItem MonthName(12)
    Combo4.AddItem "13th Month"
    
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT DISTINCT PAY_YEAR FROM HRMS_PAYROLL ORDER BY PAY_YEAR DESC")
    Combo_Loadval Combo5, rsTemp
    'Call FillcboYear(Combo5)
    fillcombo_up Combo5
    Set rsTemp = Nothing
End Sub

Function CheckIfExmployeeIsActive(XEMPNO As String) As Boolean
    Dim RSTMP                           As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT ACTIVEINACTIVE FROM HRMS_EMPINFO WHERE " & _
        " EMPNO = " & N2Str2Null(XEMPNO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!ACTIVEINACTIVE) = "A" Then
            CheckIfExmployeeIsActive = True
        Else
            CheckIfExmployeeIsActive = False
        End If
    End If
    Set RSTMP = Nothing
End Function

Sub FillGrid(VCUT_OFF As String, vMONTH As Integer, vYEAR As Integer, VOPTION As String, CREDITDATE As String, BATCHNO As String, Optional xx13thMonth As String)
    Grid1.Enabled = True
    Grid1.Rows = 1
    Dim rsTemp                                                        As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Dim NOT_ON_LIST                                                   As String
    
    If VOPTION = "PAYROLL" Then
        If xx13thMonth = "13th Month" Then
            Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL WHERE " & _
                " PAY_MONTH = 13 " & _
                " AND PAY_YEAR = '" & vYEAR & "'")
            If Not rsTemp.EOF And Not rsTemp.BOF Then
                rsTemp.MoveFirst
                While Not rsTemp.EOF
                    If CheckIfExmployeeIsActive(rsTemp!EMPNO) = True Then
                        If Null2String(GET_EMP_ACCTNO(Null2String(rsTemp!EMPNO))) <> "" And (N2Str2Zero(rsTemp!ALLOWANCE) + N2Str2Zero(rsTemp!NETPAY)) >= 0 Then
                            Grid1.AddItem "REMOVE" & Chr(9) & _
                                    Null2String(rsTemp!EMPNO) & Chr(9) & _
                                    GET_EMP_NAME(Null2String(rsTemp!EMPNO)) & Chr(9) & _
                                    GET_EMP_ACCTNO(Null2String(rsTemp!EMPNO)) & Chr(9) & _
                                    (N2Str2Zero(rsTemp!ALLOWANCE) + N2Str2Zero(rsTemp!NETPAY)) & Chr(9) & _
                                    GETHASHVALUE(GET_EMP_ACCTNO(Null2String(rsTemp!EMPNO)), N2Str2Zero(rsTemp!ALLOWANCE) + N2Str2Zero(rsTemp!NET_AMT13))
                        Else
                            NOT_ON_LIST = NOT_ON_LIST & vbCrLf & GET_EMP_NAME(Null2String(rsTemp!EMPNO))
                        End If
                    End If
                    rsTemp.MoveNext
                Wend
            End If
        Else
            Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL WHERE CUT_OFF = '" & VCUT_OFF & "' AND  PAY_MONTH = '" & vMONTH & "' AND PAY_YEAR = '" & vYEAR & "'")
            If Not rsTemp.EOF And Not rsTemp.BOF Then
                rsTemp.MoveFirst
                While Not rsTemp.EOF
                    If CheckIfExmployeeIsActive(rsTemp!EMPNO) = True Then
                        If Null2String(GET_EMP_ACCTNO(Null2String(rsTemp!EMPNO))) <> "" And (N2Str2Zero(rsTemp!ALLOWANCE) + N2Str2Zero(rsTemp!NETPAY)) >= 0 Then
                            Grid1.AddItem "REMOVE" & Chr(9) & _
                                Null2String(rsTemp!EMPNO) & Chr(9) & _
                                GET_EMP_NAME(Null2String(rsTemp!EMPNO)) & Chr(9) & _
                                GET_EMP_ACCTNO(Null2String(rsTemp!EMPNO)) & Chr(9) & _
                                (N2Str2Zero(rsTemp!ALLOWANCE) + N2Str2Zero(rsTemp!NETPAY)) & Chr(9) & _
                                GETHASHVALUE(GET_EMP_ACCTNO(Null2String(rsTemp!EMPNO)), N2Str2Zero(rsTemp!ALLOWANCE) + N2Str2Zero(rsTemp!NETPAY))
                        Else
                            NOT_ON_LIST = NOT_ON_LIST & vbCrLf & GET_EMP_NAME(Null2String(rsTemp!EMPNO))
                        End If
                    End If
                    rsTemp.MoveNext
                Wend
            End If
        End If
        
        If Len(NOT_ON_LIST) > 0 Then
            MsgInformation "EMPLOYEE(S) NOT ON LIST BECAUSE OF INVALID AND MISSING ENTRIES" & NOT_ON_LIST
        End If
        NOT_ON_LIST = ""
    ElseIf VOPTION = "HASH" Then
        'Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_HORIZONTAL_HASH WHERE " & _
        '    " BATCH_NO = '" & BATCHNO & _
        '    "' AND CREDIT_DATE = '" & CREDITDATE & "'")
        
        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_HORIZONTAL_HASH WHERE " & _
            " BATCH_NO = '" & BATCHNO & _
            "' AND MONTH(CREDIT_DATE) = " & MONTH(CREDITDATE) & _
            " AND DAY(CREDIT_DATE) = " & Day(CREDITDATE) & _
            " AND YEAR(CREDIT_DATE) = " & YEAR(CREDITDATE) & "")
        If Not rsTemp.EOF And Not rsTemp.BOF Then
            rsTemp.MoveFirst
            While Not rsTemp.EOF
                Grid1.AddItem "REMOVE" & Chr(9) & Null2String(rsTemp!EMPNO) & Chr(9) & GET_EMP_NAME(Null2String(rsTemp!EMPNO)) & Chr(9) & Null2String(rsTemp!ACCT_NO) & Chr(9) & N2Str2Zero(rsTemp!net) & Chr(9) & N2Str2Zero(rsTemp!HOR_HASH)
                rsTemp.MoveNext
            Wend
        End If
    End If
    
    Grid1.Column(3).Sort (cellAscending)
    Call ComputeTotal
End Sub

Sub ComputeTotal()
    Dim I                                                             As Integer
    I = 1

    Dim TotalNet                                                      As Double
    Dim TotalHash                                                     As Double
    Dim TotalAccnt                                                    As Double

    TotalNet = 0
    TotalHash = 0
    TotalAccnt = 0

    While I < Grid1.Rows
        TotalNet = TotalNet + N2Str2Zero(Grid1.Cell(I, 5).Text)
        TotalHash = TotalHash + N2Str2Zero(Grid1.Cell(I, 6).Text)
        TotalAccnt = TotalAccnt + GET_NUMBER(Grid1.Cell(I, 4).Text)
        I = I + 1
    Wend
    TotalNet = Round(TotalNet, 2)
    Text2.Text = Format(TotalNet, "#####################.00")
    TotalHash = Round(TotalHash, 2)
    Text3.Text = Format(TotalHash, "#####################.00")
    TotalAccnt = Round(TotalAccnt, 2)
    Text1.Text = Format(TotalAccnt, "#####################.00")
    
    'Text2.Text = ToDoubleNumber(TotalNet)
    'TotalHash = ToDoubleNumber(TotalHash)
    'Text3.Text = ToDoubleNumber(TotalHash)
    'TotalAccnt = ToDoubleNumber(TotalAccnt)
    'Text1.Text = ToDoubleNumber(TotalAccnt)
End Sub

Private Sub cmdEdit_Click()
    picSaves.Visible = True
    picAdd.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    If txtBATCHNO.Text <> "" Then
        Dim rsTemp                                                    As New ADODB.Recordset
        Set rsTemp = New ADODB.Recordset

'        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_HORIZONTAL_HASH WHERE " & _
'            " BATCH_NO = " & N2Str2IntZero(txtBATCHNO) & _
'            " AND CREDIT_DATE = " & N2Str2Null(DateValue(txtCREDITDATE)))

        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_HORIZONTAL_HASH WHERE " & _
            " BATCH_NO = " & N2Str2IntZero(txtBATCHNO) & _
            " AND MONTH(CREDIT_DATE) = " & txtCREDITDATE.MONTH & _
            " AND DAY(CREDIT_DATE) = " & txtCREDITDATE.Day & _
            " AND YEAR(CREDIT_dATE) = " & txtCREDITDATE.YEAR & "")
        If Not (rsTemp.EOF Or rsTemp.BOF) Then
            Call InitGrid
            Call FillGrid("", 0, 0, "HASH", DateValue(txtCREDITDATE.Value), txtBATCHNO.Text)
        Else
            If MsgBox("No record(s) with that credit date and " & vbCrLf & "batch number exist(s).Create?", vbInformation + vbYesNo, "Confirm") = vbYes Then
                Grid1.Rows = 1
                ComputeTotal
                Picture1.Enabled = True
            Else
                Grid1.Rows = 1
                ComputeTotal
                Picture1.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    If txtBATCHNO <> "" And Text2.Text <> "" And Text1.Text <> "" And Text3.Text <> "" Then
        rptHash.WindowTitle = "ATM Summary Report"
        rptHash.Formulas(0) = "COMPANY_CODE = '" & txtCOMPANYCODE & "'"
        rptHash.Formulas(1) = "ACCOUNT_NO = '" & txtACCOUNTNO & "'"
        rptHash.Formulas(3) = "COMPANY_CODE = '" & txtCOMPANYCODE & "'"
        rptHash.Formulas(4) = "CEILING_AMT = '" & txtCEILINGAMOUNT & "'"
        rptHash.Formulas(5) = "PAYROLL_AMT = " & Text2.Text
        rptHash.Formulas(6) = "ACCNT_TOTAL = " & Text1.Text
        rptHash.Formulas(7) = "RECORD_CNT = '" & Grid1.Rows - 1 & "'"
        rptHash.Formulas(8) = "CREDIT_DATE = '" & txtCREDITDATE & "'"
        PrintSQLReport rptHash, HRMS_REPORT_PATH & "HOR_HASH.rpt", "date({HOR_HAS.CREDIT_DATE}) = CDATE('" & txtCREDITDATE & "') AND {HOR_HAS.BATCH_NO} = " & txtBATCHNO, DMIS_REPORT_Connection, 1
    End If
End Sub

Private Sub Command2_Click()
    Command2.Visible = False
    Command3.Visible = True
    txtCOMPANYCODE.Enabled = True
    txtPRESENTINGOFFICE.Enabled = True
    txtACCOUNTNO.Enabled = True
    txtCEILINGAMOUNT.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Visible = False
    Command2.Visible = True
    txtCOMPANYCODE.Enabled = False
    txtPRESENTINGOFFICE.Enabled = False
    txtACCOUNTNO.Enabled = False
    txtCEILINGAMOUNT.Enabled = False

    Dim VtxtCOMPANYCODE                                               As String
    Dim VTXTACCOUNTNO                                                 As String
    Dim VtxtCEILINGAMOUNT                                             As String
    Dim VtxtPRESENTINGOFFICE                                          As String

    VtxtCOMPANYCODE = N2Str2Null(txtCOMPANYCODE)
    VTXTACCOUNTNO = N2Str2Null(txtACCOUNTNO)
    VtxtCEILINGAMOUNT = N2Str2Zero(NumericVal(txtCEILINGAMOUNT))
    VtxtPRESENTINGOFFICE = N2Str2Null(txtPRESENTINGOFFICE)

    gconDMIS.Execute ("UPDATE HRMS_HOR_HASH_HEADER SET" & _
                    " COMPANY_CODE = " & VtxtCOMPANYCODE & ", " & _
                    " ACCOUNT_NO = " & VTXTACCOUNTNO & ", " & _
                    " CEILING_AMOUNT = " & VtxtCEILINGAMOUNT & ", " & _
                    " PRESENTING_OFFICE = " & VtxtPRESENTINGOFFICE & "")
    Fill_Header
End Sub

Private Sub Command4_Click()
    Dim CUT_OFF_PERIOD                                                As String
    CUT_OFF_PERIOD = ""
    If Combo2.Text = "1st Cut-Off" Then
        CUT_OFF_PERIOD = "1"
    ElseIf Combo2.Text = "2nd Cut-Off" Then
        CUT_OFF_PERIOD = "2"
    End If
    If Combo2.Text <> "" And Combo4.Text <> "" And Combo5.Text <> "" Then
        Dim rsTemp                                                    As New ADODB.Recordset
        If Combo4.Text = "13th Month" Then
            Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL WHERE " & _
                " PAY_MONTH = 13 " & _
                " AND PAY_YEAR = '" & Combo5.Text & "'")
            If Not rsTemp.EOF And Not rsTemp.BOF Then
                Call InitGrid
                Call FillGrid(CUT_OFF_PERIOD, What_month(Combo4.Text), Combo5.Text, "PAYROLL", "", "", Combo4)
                Picture1.Enabled = False
            Else
                MsgBox "No 13month pay is generated on this pay period. Please select another pay period"
                Picture1.Enabled = True
                Grid1.Rows = 1
                Call ComputeTotal
                Grid1.Enabled = False
            End If
        Else
            Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL WHERE CUT_OFF = '" & CUT_OFF_PERIOD & "' AND  PAY_MONTH = '" & What_month(Combo4.Text) & "' AND PAY_YEAR = '" & Combo5.Text & "'")
            If Not rsTemp.EOF And Not rsTemp.BOF Then
                InitGrid
                Call FillGrid(CUT_OFF_PERIOD, What_month(Combo4.Text), Combo5.Text, "PAYROLL", "", "")
                Picture1.Enabled = False
            Else
                MsgBox "No payroll is generated on this pay period. Please select another pay period"
                Picture1.Enabled = True
                Grid1.Rows = 1
                Call ComputeTotal
                Grid1.Enabled = False
            End If
        End If
    Else
        MsgBox "Please complete the entries"
        Picture1.Enabled = True
        Grid1.Rows = 1
        Call ComputeTotal
        Grid1.Enabled = False
    End If
End Sub

Private Sub Command5_Click()
        
    If MsgBox("Create ATM Disk, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    Dim I                                                             As Integer
    I = 1
    Dim HEADER                                                        As String * 128
    Dim DETAIL                                                        As String * 128
    Dim TAIL                                                          As String * 128
    Dim ENDING                                                        As String * 1
    HEADER = ""
    DETAIL = ""
    TAIL = ""
    CommonDialog1.Filename = "16823"
    
    On Error GoTo ErrorHandler
    CommonDialog1.ShowSave
    Open CommonDialog1.Filename For Output As #1
    HEADER = "H" & txtCOMPANYCODE & Format(txtCREDITDATE, "mmddyy") & Format(txtBATCHNO, "00") & "1" & GET_STRING(txtACCOUNTNO) & txtPRESENTINGOFFICE & Format(GET_STRING(txtCEILINGAMOUNT), "000000000000") & Format(GET_STRING(Text2), "000000000000") & "1"
    Print #1, HEADER
    While I < Grid1.Rows
        DETAIL = "D" & txtCOMPANYCODE & Format(txtCREDITDATE, "mmddyy") & Format(txtBATCHNO, "00") & "3" & Format(GET_NUMBER(Grid1.Cell(I, 4).Text), "0000000000") & Format(GET_STRING(Format(Grid1.Cell(I, 5).Text, "#######################.00")), "000000000000") & Format(GET_STRING(Format(Grid1.Cell(I, 6).Text, "##################.00")), "000000000000")
        Print #1, DETAIL
        I = I + 1
    Wend
    TAIL = "T" & txtCOMPANYCODE & Format(txtCREDITDATE, "mmddyy") & Format(txtBATCHNO, "00") & "2" & GET_STRING(txtACCOUNTNO) & Format(Text1, "000000000000000") & Format(GET_STRING(Text2), "000000000000000") & Format(GET_STRING(Text3), "000000000000000000") & Format(I - 1, "00000")
    Print #1, TAIL
    ENDING = Chr(26)
    Print #1, ENDING;
    Close #1
    
    MsgBox "Finished Creating the file", vbInformation, "Info"
    
ErrorHandler:
    Exit Sub
    
End Sub

Private Sub Command6_Click()
    If MsgBox("Save this ATM Summary, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    gconDMIS.Execute ("DELETE FROM HRMS_HORIZONTAL_HASH WHERE CREDIT_DATE = '" & txtCREDITDATE & "' AND BATCH_NO = '" & txtBATCHNO & "'")
    Dim I                                                             As Integer
    I = 1
    While I < Grid1.Rows
        gconDMIS.Execute "INSERT INTO HRMS_HORIZONTAL_HASH (EMPNO, ACCT_NO, NET, HOR_HASH, BATCH_NO, CREDIT_DATE)" & _
                       " VALUES (" & N2Str2Null(Grid1.Cell(I, 2).Text) & "," & N2Str2Null(Grid1.Cell(I, 4).Text) & "," & N2Str2Null(Grid1.Cell(I, 5).Text) & "," & N2Str2Null(Grid1.Cell(I, 6).Text) & "," & N2Str2Null(txtBATCHNO) & "," & N2Str2Null(txtCREDITDATE) & ")"
        I = I + 1
    Wend
    MsgBox "Finish Saving The changes", vbInformation, "Info"
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    DrawXPCtl Me
    
    'Call FillcboYear(Combo5)
    
    fillcombo_up Combo5
    Fill_Header
    Fill_Pay_Period
    txtCREDITDATE.Value = Now
    Picture6.Enabled = False
    Picture1.Enabled = False
End Sub

Private Sub Grid1_DblClick()
    If Grid1.ActiveCell.Col = 1 Then
        Grid1.RemoveItem (Grid1.ActiveCell.Row)
        Grid1.Refresh
        ComputeTotal
    End If
End Sub

Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    If Col = 4 Or Col = 5 Then
        Grid1.Cell(Row, 6).Text = GETHASHVALUE(Grid1.Cell(Row, 4).Text, N2Str2Zero(Grid1.Cell(Row, 5).Text))
        ComputeTotal
    End If
End Sub


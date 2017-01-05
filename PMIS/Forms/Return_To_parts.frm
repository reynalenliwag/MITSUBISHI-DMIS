VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Return_To_parts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Parts to Parts Deptpartment"
   ClientHeight    =   8220
   ClientLeft      =   240
   ClientTop       =   1410
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Return_To_parts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   10965
   Begin VB.PictureBox PIC_SEARCH 
      Appearance      =   0  'Flat
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
      Height          =   7875
      Left            =   0
      ScaleHeight     =   7845
      ScaleWidth      =   2085
      TabIndex        =   10
      Top             =   0
      Width           =   2115
      Begin XtremeReportControl.ReportControl rcFind 
         Height          =   6255
         Left            =   60
         TabIndex        =   65
         Top             =   900
         Width           =   1965
         _Version        =   655364
         _ExtentX        =   3466
         _ExtentY        =   11033
         _StockProps     =   64
         BorderStyle     =   1
         EditOnClick     =   0   'False
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   1440
         Top             =   7530
      End
      Begin VB.TextBox txtSearch 
         Height          =   345
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   2025
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption5 
         Height          =   315
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   7545
         _Version        =   655364
         _ExtentX        =   13309
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Search by RO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label LblVerify 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NOT YET VERIFY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   390
         TabIndex        =   56
         Top             =   7350
         Width           =   1275
      End
   End
   Begin VB.PictureBox PIC_MAIN 
      Appearance      =   0  'Flat
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
      Height          =   8145
      Left            =   2100
      ScaleHeight     =   8115
      ScaleWidth      =   8835
      TabIndex        =   18
      Top             =   0
      Width           =   8865
      Begin VB.PictureBox picRetrn 
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         Height          =   825
         Left            =   150
         ScaleHeight     =   825
         ScaleWidth      =   1785
         TabIndex        =   66
         Top             =   7230
         Width           =   1785
         Begin VB.CommandButton cmdverify 
            Caption         =   "&Receive"
            Height          =   795
            Left            =   30
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "Return_To_parts.frx":1082
            MousePointer    =   99  'Custom
            Picture         =   "Return_To_parts.frx":11D4
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Post this Transaction"
            Top             =   0
            Width           =   825
         End
         Begin VB.CommandButton cmdCancelCO 
            Caption         =   "Cancel "
            Height          =   795
            Left            =   840
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "Return_To_parts.frx":2256
            MousePointer    =   99  'Custom
            Picture         =   "Return_To_parts.frx":23A8
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Cancel this Transaction"
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.PictureBox frame 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   60
         ScaleHeight     =   2025
         ScaleWidth      =   8715
         TabIndex        =   60
         Top             =   5070
         Width           =   8745
         Begin XtremeReportControl.ReportControl RcReq_parts 
            Height          =   1665
            Left            =   30
            TabIndex        =   64
            Top             =   330
            Width           =   8655
            _Version        =   655364
            _ExtentX        =   15266
            _ExtentY        =   2937
            _StockProps     =   64
            BorderStyle     =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Double click to remove item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   1
            Left            =   6330
            TabIndex        =   62
            Top             =   30
            Width           =   2310
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption6 
            Height          =   315
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   8685
            _Version        =   655364
            _ExtentX        =   15319
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Parts to be Return"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
      Begin VB.PictureBox Picture1 
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
         Height          =   1335
         Left            =   1830
         ScaleHeight     =   1335
         ScaleWidth      =   6915
         TabIndex        =   30
         Top             =   7230
         Width           =   6915
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
            Height          =   855
            Left            =   1440
            ScaleHeight     =   855
            ScaleWidth      =   5565
            TabIndex        =   31
            Top             =   0
            Width           =   5565
            Begin VB.CommandButton cmdExit 
               Caption         =   "E&xit"
               Height          =   795
               Left            =   4650
               MouseIcon       =   "Return_To_parts.frx":26E2
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":2834
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Exit Window"
               Top             =   30
               Width           =   735
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "&Print"
               Enabled         =   0   'False
               Height          =   795
               Left            =   3930
               MouseIcon       =   "Return_To_parts.frx":2B9A
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":2CEC
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Print this Record"
               Top             =   30
               Width           =   735
            End
            Begin VB.CommandButton cmdUnpost 
               Caption         =   "Unpost"
               Enabled         =   0   'False
               Height          =   795
               Left            =   3210
               MaskColor       =   &H0000FFFF&
               MouseIcon       =   "Return_To_parts.frx":3052
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":31A4
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Post this Transaction"
               Top             =   30
               Width           =   735
            End
            Begin VB.CommandButton cmdPost 
               Caption         =   "Post"
               Enabled         =   0   'False
               Height          =   795
               Left            =   2460
               MaskColor       =   &H0000FFFF&
               MouseIcon       =   "Return_To_parts.frx":34C9
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":361B
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Post this Transaction"
               Top             =   30
               Width           =   765
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               Enabled         =   0   'False
               Height          =   795
               Left            =   1740
               MouseIcon       =   "Return_To_parts.frx":3940
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":3A92
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Delete Selected Record"
               Top             =   30
               Width           =   735
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "&Edit"
               Height          =   795
               Left            =   1020
               MouseIcon       =   "Return_To_parts.frx":3DBD
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":3F0F
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   "Edit Selected Record"
               Top             =   30
               Width           =   735
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   795
               Left            =   270
               MouseIcon       =   "Return_To_parts.frx":426B
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":43BD
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "Add Record"
               Top             =   30
               Width           =   765
            End
         End
         Begin Crystal.CrystalReport rptReturn 
            Left            =   1080
            Top             =   420
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
         Begin VB.PictureBox picsave 
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
            Height          =   825
            Left            =   5370
            ScaleHeight     =   825
            ScaleWidth      =   1545
            TabIndex        =   38
            Top             =   30
            Width           =   1545
            Begin VB.CommandButton cmdTranCancel 
               Caption         =   "&Cancel"
               Height          =   795
               Left            =   750
               MouseIcon       =   "Return_To_parts.frx":46D0
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":4822
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Cancel Entry"
               Top             =   0
               Width           =   735
            End
            Begin VB.CommandButton cmdTranSave 
               Caption         =   "&Save"
               Height          =   795
               Left            =   30
               MouseIcon       =   "Return_To_parts.frx":4B60
               MousePointer    =   99  'Custom
               Picture         =   "Return_To_parts.frx":4CB2
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Save Entry"
               Top             =   0
               Width           =   735
            End
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   2355
         Left            =   60
         ScaleHeight     =   2325
         ScaleWidth      =   8715
         TabIndex        =   19
         Top             =   30
         Width           =   8745
         Begin VB.TextBox txtRtnDate 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5610
            TabIndex        =   24
            Top             =   390
            Width           =   2985
         End
         Begin VB.TextBox txt_req_by 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            TabIndex        =   23
            Top             =   750
            Width           =   2865
         End
         Begin VB.TextBox txtremarks 
            Enabled         =   0   'False
            Height          =   855
            Left            =   90
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1380
            Width           =   8535
         End
         Begin VB.TextBox txtverify 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5610
            TabIndex        =   20
            Top             =   780
            Width           =   2985
         End
         Begin VB.TextBox txtRep_or 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            TabIndex        =   21
            Top             =   360
            Width           =   1935
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
            Height          =   315
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   8685
            _Version        =   655364
            _ExtentX        =   15319
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Select Parts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Return Date :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4350
            TabIndex        =   29
            Top             =   450
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Repair Order :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   28
            Top             =   420
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Request By :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   210
            TabIndex        =   27
            Top             =   810
            Width           =   1020
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Remarks:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   435
            TabIndex        =   26
            Top             =   1140
            Width           =   795
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Verfied By :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4485
            TabIndex        =   25
            Top             =   840
            Width           =   945
         End
      End
      Begin VB.PictureBox Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2685
         Left            =   60
         ScaleHeight     =   2655
         ScaleWidth      =   8715
         TabIndex        =   57
         Top             =   2370
         Width           =   8745
         Begin XtremeReportControl.ReportControl RcParts 
            Height          =   2295
            Left            =   30
            TabIndex        =   63
            Top             =   330
            Width           =   8655
            _Version        =   655364
            _ExtentX        =   15266
            _ExtentY        =   4048
            _StockProps     =   64
            BorderStyle     =   1
            AllowColumnRemove=   0   'False
            AllowColumnResize=   0   'False
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   315
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   8685
            _Version        =   655364
            _ExtentX        =   15319
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Select Parts"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
      Begin VB.Label LABID_HD 
         Caption         =   "0"
         Height          =   255
         Left            =   150
         TabIndex        =   43
         Top             =   7290
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.PictureBox pic_Select 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   1800
      ScaleHeight     =   2955
      ScaleWidth      =   6165
      TabIndex        =   11
      Top             =   2528
      Visible         =   0   'False
      Width           =   6195
      Begin VB.CommandButton cmdCancelSelect 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   5070
         TabIndex        =   14
         Top             =   2490
         Width           =   1035
      End
      Begin MSComctlLib.ListView lvwSelect 
         Height          =   2055
         Left            =   60
         TabIndex        =   13
         Top             =   390
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "itemno"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Partnumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PartDesc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Issued Qty"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Select"
         Height          =   405
         Left            =   4050
         TabIndex        =   41
         Top             =   2490
         Width           =   1035
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption7 
         Height          =   345
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   6585
         _Version        =   655364
         _ExtentX        =   11615
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "PART NUMBER TO BE RETURN IN PARTS DEPARTMENT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox pic_Returnpart 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   3510
      ScaleHeight     =   2955
      ScaleWidth      =   6165
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   6195
      Begin VB.TextBox txttype 
         Alignment       =   2  'Center
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1380
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   5250
         MousePointer    =   99  'Custom
         Picture         =   "Return_To_parts.frx":5002
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Edit Selected Record"
         Top             =   2070
         Width           =   645
      End
      Begin VB.TextBox txtTran_issued 
         Alignment       =   2  'Center
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtTran_part_Desc 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtTran_return 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   15
         Text            =   "0"
         Top             =   2220
         Width           =   855
      End
      Begin VB.CommandButton cmdview 
         Caption         =   "View"
         Height          =   315
         Left            =   4860
         TabIndex        =   2
         Top             =   510
         Width           =   1035
      End
      Begin VB.CommandButton cmdOk_det 
         Caption         =   "&Ok"
         Height          =   795
         Left            =   4620
         MousePointer    =   99  'Custom
         Picture         =   "Return_To_parts.frx":6084
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Edit Selected Record"
         Top             =   2070
         Width           =   645
      End
      Begin VB.ComboBox cbo_Tran_Partnumber 
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
         Height          =   345
         Left            =   2160
         TabIndex        =   42
         Top             =   510
         Width           =   2715
      End
      Begin VB.CommandButton cmdDelete_det 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   795
         Left            =   3960
         MousePointer    =   99  'Custom
         Picture         =   "Return_To_parts.frx":7106
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Edit Selected Record"
         Top             =   2070
         Width           =   675
      End
      Begin VB.Label lbltrantype 
         Height          =   375
         Left            =   3120
         TabIndex        =   70
         Top             =   2100
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LABID_DET 
         Caption         =   "0"
         Height          =   435
         Left            =   4740
         TabIndex        =   53
         Top             =   1500
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblID 
         Caption         =   "0"
         Height          =   375
         Left            =   3450
         TabIndex        =   52
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1110
         TabIndex        =   51
         Top             =   1470
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PARTNUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   915
         TabIndex        =   3
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "QTY TO BE RETURN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   540
         TabIndex        =   7
         Top             =   2310
         Width           =   1590
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ISSUED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   870
         TabIndex        =   6
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PARTS DESRIPTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   495
         TabIndex        =   5
         Top             =   1050
         Width           =   1605
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
         Height          =   345
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6585
         _Version        =   655364
         _ExtentX        =   11615
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "PART NUMBER TO BE RETURN IN PARTS DEPARTMENT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.Frame frameSupp 
      BackColor       =   &H80000013&
      Caption         =   "Select Supplier"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4995
      Left            =   3570
      TabIndex        =   71
      Top             =   1680
      Width           =   6015
      Begin VB.CommandButton cmdCancelSupp 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4770
         TabIndex        =   72
         Top             =   4440
         Width           =   1125
      End
      Begin VB.TextBox txtSupCode 
         Enabled         =   0   'False
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
         Left            =   1260
         TabIndex        =   79
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox txtSuppNAME 
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
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   750
         Width           =   4605
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Caption         =   "Find Supplier"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   120
         TabIndex        =   74
         Top             =   1230
         Width           =   5805
         Begin VB.TextBox txtFindSupp 
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
            Left            =   1140
            TabIndex        =   75
            Top             =   300
            Width           =   4515
         End
         Begin MSComctlLib.ListView lvwSupp 
            Height          =   2295
            Left            =   60
            TabIndex        =   76
            Top             =   720
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   4048
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Supplier Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Supplier Name"
               Object.Width           =   7937
            EndProperty
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Find Supplier:"
            Height          =   285
            Left            =   90
            TabIndex        =   77
            Top             =   330
            Width           =   1155
         End
      End
      Begin VB.CommandButton cmdSaveSupp 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3660
         TabIndex        =   73
         Top             =   4440
         Width           =   1125
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code:"
         Height          =   315
         Left            =   150
         TabIndex        =   81
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name:"
         Height          =   315
         Left            =   150
         TabIndex        =   80
         Top             =   780
         Width           =   1125
      End
   End
   Begin VB.PictureBox pic_Post 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   3180
      ScaleHeight     =   1785
      ScaleWidth      =   3405
      TabIndex        =   44
      Top             =   3113
      Visible         =   0   'False
      Width           =   3435
      Begin VB.CommandButton cmd_POST_Cancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   2520
         MouseIcon       =   "Return_To_parts.frx":8188
         MousePointer    =   99  'Custom
         Picture         =   "Return_To_parts.frx":920A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Edit Selected Record"
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox txtVerified_BY 
         Enabled         =   0   'False
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   450
         Width           =   3195
      End
      Begin VB.CommandButton cmd_POST_OK 
         Caption         =   "&Ok"
         Height          =   795
         Left            =   1740
         MouseIcon       =   "Return_To_parts.frx":A28C
         MousePointer    =   99  'Custom
         Picture         =   "Return_To_parts.frx":A3DE
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Edit Selected Record"
         Top             =   900
         Width           =   795
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   6585
         _Version        =   655364
         _ExtentX        =   11615
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Verify By"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   330
      Left            =   3840
      TabIndex        =   69
      Top             =   3690
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   582
      Picture         =   "Return_To_parts.frx":B460
      ForeColor       =   0
      Appearance      =   2
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "Return_To_parts.frx":B47C
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F1 - Select Supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   0
      TabIndex        =   82
      Top             =   7860
      Width           =   2115
   End
End
Attribute VB_Name = "Return_To_parts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADDOREDIT_HD                    As String
Dim ADDOREDIT_DET                   As String
Dim RS_RETURN                       As ADODB.Recordset
Dim ISCLICK                         As Boolean
Dim LOAD_DATA                       As Boolean
Dim LOCALACESS                      As String
Dim OnUpdate                        As Boolean
Dim What_Func                       As Boolean
Dim issued                          As Boolean
Dim str_catcher                     As String
Dim Valid                           As Boolean
Dim unionquery                      As String

Private Sub cmdAdd_Click()
        If Function_Access(LOGID, "Acess_Add", LOCALACESS) = False Then Exit Sub
        
        ADDOREDIT_HD = "ADD"
        ADDOREDIT_DET = "ADD"
        txtRep_or.Enabled = True
        txtverify.Text = ""
        picsave.ZOrder 0
        picsave.Visible = True
        picAdd.Visible = False
        Call cleartxt
        On Error Resume Next
        txtRep_or.SetFocus
        RcParts.Records.DeleteAll
        RcParts.Populate
        RcReq_parts.Records.DeleteAll
        RcReq_parts.Populate
        rcFind.Enabled = False
        LABID_HD = GetLastData() + 1
        LblVerify.Caption = "NOT YET VERIFY"
        Picture2.Enabled = True
        txtRtnDate = Date
        picRetrn.Visible = False
        PIC_SEARCH.Enabled = False
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "PARTS RETURN TRANSACTION") = False Then Exit Sub
        Dim SQLMSG As String
        Dim lng As Integer
        On Error GoTo ErrorCode
        
        SQLMSG = "System Error Due to Proccess on - @09099386087." & vbCrLf
        SQLMSG = SQLMSG & "This will RollBack Transaction!" & vbCrLf
        SQLMSG = SQLMSG & "Please conctact Netspeed Software Helpdesk immediately!" & vbCrLf
        SQLMSG = SQLMSG & "Thank you!"
       
'updated By: IEBV_062720110222pm
'description:
        If (gconDMIS.Execute("SELECT count(*) FROM csms_repor where invoice is not null and rep_or = '" & txtRep_or.Text & "'").Fields(0).Value) > 0 Then
            MsgBox "Repair order alredy invoice." & vbCrLf & "System will not allow Cancellation.", vbInformation + vbOKOnly
            Exit Sub
        End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------
        If MsgBox("Do you want to Cancel the RR Entry of this Transaction? ", vbQuestion + vbYesNo) = vbYes Then
            gconDMIS.BeginTrans
            
             lng = 6
             If lng > 0 Then
                prgExcelGen.Max = lng
                prgExcelGen.Value = 0
            End If
            prgExcelGen.ZOrder 0
            prgExcelGen.Visible = True
            
            gconDMIS.Execute "UPDATE CSMS_RETURN_HD SET STATUS = 'C',veri_by = NULL WHERE REP_OR = '" & txtRep_or & "' and ID = '" & LABID_HD & "'"
            'gconDMIS.Execute "update csms_return_det set QTY_APP = 1 where ID_HD ='" & LABID_HD & "' "
            
            If ImportDetails(txtRep_or, "P", "2") = False Then
                SQLMSG = Replace(SQLMSG, "@09099386087", "Importing Process")
                MsgBox SQLMSG, vbCritical, "Processing Error"
                Screen.MousePointer = 0
                gconDMIS.RollbackTrans
                prgExcelGen.Visible = False
                'cmdverify.Enabled = True
                cmdCancelCO.Enabled = True
                Exit Sub
            Else
                prgExcelGen.Text = "Cancelling Transaction ..."
                prgExcelGen.Value = prgExcelGen.Value + 1
            End If
            
            If ImportDetails(txtRep_or, "M", "3") = False Then
                SQLMSG = Replace(SQLMSG, "@09099386087", "Importing Process")
                MsgBox SQLMSG, vbCritical, "Processing Error"
                Screen.MousePointer = 0
                gconDMIS.RollbackTrans
                prgExcelGen.Visible = False
                'cmdverify.Enabled = True
                cmdCancelCO.Enabled = True
                Exit Sub
            Else
                prgExcelGen.Text = "Cancelling Transaction ..."
                prgExcelGen.Value = prgExcelGen.Value + 1
            End If
            
            If ImportDetails(txtRep_or, "A", "4") = False Then
                SQLMSG = Replace(SQLMSG, "@09099386087", "Importing Process")
                MsgBox SQLMSG, vbCritical, "Processing Error"
                Screen.MousePointer = 0
                gconDMIS.RollbackTrans
                prgExcelGen.Visible = False
                'cmdverify.Enabled = True
                cmdCancelCO.Enabled = True
                Exit Sub
            Else
                prgExcelGen.Text = "Cancelling Transaction ..."
                prgExcelGen.Value = prgExcelGen.Value + 1
            End If
'-------------------------------------------------------------------------------------------------
'updated by:    IEBV 02192011_0350pm
'description:   to avoid double recieving if canceled

            If CancelRecieving(txtRep_or.Text) = False Then
                If issued = True Then
                    SQLMSG = "Cannot cancel recieving. " & vbCrLf
                    SQLMSG = SQLMSG & "Parts already issued. "
                Else
                    SQLMSG = Replace(SQLMSG, "@09099386087", "Importing Process")
                End If
                
                MsgBox SQLMSG, vbCritical, "Processing Error"
                Screen.MousePointer = 0
                gconDMIS.RollbackTrans
                prgExcelGen.Visible = False
                'cmdverify.Enabled = True
                cmdCancelCO.Enabled = True
                issued = False
                Exit Sub
            Else
                prgExcelGen.Value = prgExcelGen.Value + 1
                prgExcelGen.Text = "Canceling Receiving Entry ... "
            End If
'-------------------------------------------------------------------------------------------------
            If UPDATE_COLUMN_ONHAND_PARTSMASTERFILE = False Then
                SQLMSG = Replace(SQLMSG, "@09099386087", "UPDATING ONHAND IN MASTERFILE")
                MsgBox SQLMSG, vbCritical, "Processing Error"
                Screen.MousePointer = 0
                gconDMIS.RollbackTrans
                prgExcelGen.Visible = False
                'cmdverify.Enabled = True
                cmdCancelCO.Enabled = True
                Exit Sub
            Else
                prgExcelGen.Value = prgExcelGen.Value + 1
                prgExcelGen.Text = "Canceling Receiving Entry ... "
            End If
        
            If UPDATE_COLUMN_RECEIPTS_PARTSMASTERFILE = False Then
                SQLMSG = Replace(SQLMSG, "@09099386087", "UPDATING TOTAL RECEIPTS IN MASTERFILE")
                MsgBox SQLMSG, vbCritical, "Processing Error"
                Screen.MousePointer = 0
                gconDMIS.RollbackTrans
                prgExcelGen.Visible = False
                'cmdverify.Enabled = True
                cmdCancelCO.Enabled = True
                Exit Sub
            Else
                prgExcelGen.Value = prgExcelGen.Value + 1
                'prgExcelGen.Text = "Creating Receiving Entry ... "
            End If
'           Call DELETE_PARTS_IN_CSMS_RO_DET(txtRep_or.Text)
'           Call GET_THE_PREV_RO_DETAILS(txtRep_or.Text)
'           Call UPDATE_CSMS_RO_DET_LINE_SEQUENTIALLY(txtRep_or.Text)
            rcFind.Records.DeleteAll
            rcFind.Populate
            
            cmdverify.Enabled = False
            cmdCancelCO.Enabled = False
            prgExcelGen.Visible = False
            
            Call rsRefresh
            RS_RETURN.Find "ID = " & LABID_HD
            Call StoreMemVars
            Call ShowTranNo
            Call NEW_LogAudit("C", "PARTS RETURN TRANSACTION", "", LABID_HD, "", "CODE: " & LABID_HD, "", "")
            gconDMIS.CommitTrans
        Else
            Exit Sub
        End If
      
        Exit Sub
ErrorCode:
        MsgBox err.Description
        Exit Sub
End Sub

Private Sub cmdCancelSelect_Click()
        pic_Returnpart.ZOrder 0
        pic_Returnpart.Visible = True
        pic_Select.Visible = False
End Sub

Sub cleartxt()
        txt_req_by = ""
        txtRemarks = ""
        txtRep_or = ""
        txtVerified_BY = ""
       '; rcParts.ListItems.Clear
End Sub

Sub rsRefresh()
    Set RS_RETURN = New ADODB.Recordset
    Call RS_RETURN.Open("select * from csms_return_hd  order by id desc", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub
Sub StoreMemVars()
    If Not (RS_RETURN.EOF Or RS_RETURN.BOF) Then
            LABID_HD = RS_RETURN!ID
            txtRep_or = Null2String(RS_RETURN!REP_OR)
            txtRtnDate = Null2String(RS_RETURN!DATE_REQ)
            txt_req_by = Null2String(RS_RETURN!REQ_BY)
            txtRemarks = Null2String(RS_RETURN!REMARKS)
            'txtRep_or = Null2String(RS_RETURN!Status)
            txtverify.Text = Null2String(RS_RETURN!VERI_BY)
            
            OnUpdate = True
            If Null2String(RS_RETURN!Status) = "P" Then
                cmdUnPost.Enabled = True
                cmdPost.Enabled = False
                'cmdAdd.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdPrint.Enabled = True
                picRetrn.Visible = True
                txtverify.Enabled = True
                cmdCancelCO.Enabled = True
                'Picture2.Enabled = False
            ElseIf Null2String(RS_RETURN!Status) = "N" Then
                cmdUnPost.Enabled = False
                cmdPost.Enabled = True
                'cmdAdd.Enabled = True
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
                cmdPrint.Enabled = False
                picRetrn.Visible = False
                txtverify.Enabled = False
                cmdCancelCO.Enabled = False
            ElseIf Null2String(RS_RETURN!Status) = "C" Then
                cmdUnPost.Enabled = False
                cmdPost.Enabled = True
                'cmdAdd.Enabled = False
                cmdEdit.Enabled = True
                cmdDelete.Enabled = False
                cmdPrint.Enabled = False
                picRetrn.Visible = False
                txtverify.Enabled = False
                cmdCancelCO.Enabled = False
            Else
                cmdUnPost.Enabled = False
                cmdPost.Enabled = True
                'cmdAdd.Enabled = True
                cmdEdit.Enabled = True
                cmdDelete.Enabled = True
                cmdPrint.Enabled = False
                picRetrn.Visible = False
                txtverify.Enabled = False
                cmdCancelCO.Enabled = False
            End If
            
            If txtverify = "" Then
                If Null2String(RS_RETURN!Status) = "C" Then
                    LblVerify.Caption = "CANCELLED"
                    cmdverify.Visible = True
                    cmdCancelCO.Visible = True
                Else
                    LblVerify.Caption = "NOT YET VERIFY"
                    cmdverify.Visible = True
                    cmdCancelCO.Visible = True
                End If
            Else
                If Null2String(RS_RETURN!Status) = "P" Then
                    LblVerify.Caption = "VERIFIED"
                    cmdUnPost.Enabled = False
                ElseIf Null2String(RS_RETURN!Status) = "C" Then
                    LblVerify.Caption = "CANCELLED"
                    cmdUnPost.Enabled = True
                End If
            End If
            
            If txtverify = "" Then
                cmdverify.Enabled = True
                txtverify.Enabled = True
                cmdCancelCO.Enabled = False
            Else
                cmdverify.Enabled = False
                txtverify.Enabled = False
                cmdCancelCO.Enabled = True
            End If
            
            Call show_allparts(txtRep_or)
            Call Show_req_parts(txtRep_or)
    Else
            
            ShowNoRecord
            cmdAdd.Value = True
            
    End If

End Sub

Function CancelRecieving(ro As String) As Boolean
On Error GoTo ErrorCode

Dim SQLTXT                  As String
Dim rsTMP                   As New ADODB.Recordset
Dim itemno                  As Integer
Dim rssopot                 As ADODB.Recordset
Dim rsboto                  As ADODB.Recordset
Dim rscheckissue            As ADODB.Recordset
Dim asofonhand              As Integer

SQLTXT = "SELECT * FrOM CSMS_RETURN_DET A INNER JOIN PMIS_ALLDAYTRAN B" & vbCrLf
SQLTXT = SQLTXT & "ON A.ITEMID = B.ID WHERE A.REP_OR = '" & ro & "' and ID_HD = '" & LABID_HD & "'" & vbCrLf

Set rsTMP = gconDMIS.Execute(SQLTXT)

If Not (rsTMP.BOF And rsTMP.BOF) Then
    rsTMP.MoveFirst
    Do While Not rsTMP.EOF
        Dim rsivan As ADODB.Recordset
        Set rsivan = New ADODB.Recordset
        Set rsivan = gconDMIS.Execute("Select 'history' as wen,* from  PMIS_rec_hist where type = '" & rsTMP!Type & "' and  RIV_TRANNO = '" & rsTMP!TRANNO & "' and CLASSCODE = 'RRV' and STATUS = 'P'" & _
                                      " UNION ALL " & _
                                      " Select 'now' as wen,* from  PMIS_rr_HD where type = '" & rsTMP!Type & "' and  RIV_TRANNO = '" & rsTMP!TRANNO & "' and CLASSCODE = 'RRV' and STATUS = 'P'")
            If Not (rsivan.EOF And rsivan.BOF) Then
                    Dim Xrr_HD               As String
                    Dim Xrr_Detail           As String
                    If (rsivan!wen) = "history" Then
                        Xrr_HD = "PMIS_REC_HIST"
                        Xrr_Detail = "PMIS_DAYTRAN"
                    Else
                        Xrr_HD = "PMIS_RR_HD"
                        Xrr_Detail = "PMIS_TDAYTRAN"
                    End If
                    
                    Set rssopot = New ADODB.Recordset
                    Set rssopot = gconDMIS.Execute("Select * from " & Xrr_HD & " where rrno = '" & rsivan!RRNO & "' and [type] = '" & rsivan!Type & "' ")
                    If Not (rssopot.EOF And rssopot.BOF) Then
                        Set rsboto = New ADODB.Recordset
                        Set rsboto = gconDMIS.Execute("Select * from " & Xrr_Detail & " where tranno = '" & rssopot!RRNO & "' and [type] = '" & rssopot!Type & "' and trantype = 'RR' order by itemno asc")
                        If Not (rsboto.EOF And rsboto.BOF) Then
                            rsboto.MoveFirst
                            Do While Not rsboto.EOF
                                Set rscheckissue = New ADODB.Recordset
                                Set rscheckissue = gconDMIS.Execute("select * from pmis_alldaytran where [type] = '" & rsboto!Type & "' and status in ('P','B') and (trantype in('DR','RIV','CHG','CSH') or (trantype = 'ADJ' and IN_OUT = 'O')) and ID > '" & rsboto!ID & "' and trandate >= '" & rssopot!RRDATE & "' and stock_ord = '" & (rsboto!STOCK_ORD) & "'")
                                If Not (rscheckissue.EOF And rscheckissue.BOF) Then
                                    issued = True
                                    CancelRecieving = False
                                    Exit Function
                                End If
                                rsboto.MoveNext
                            Loop
                        End If
                    End If
                    
                    Dim rsRR_det As ADODB.Recordset
                    Set rsRR_det = New ADODB.Recordset
                    Dim sqlcommands As String
                    
                    sqlcommands = "Update " & Xrr_Detail & " set STATUS = 'C', "
                    sqlcommands = sqlcommands + "LASTUPDATE = '" & LOGDATE & "', "
                    sqlcommands = sqlcommands + "USERCODE = " & N2Str2Null(LOGCODE) & " "
                    sqlcommands = sqlcommands + " where [TYPE] = '" & rsivan!Type & "' AND TRANTYPE = 'RR' and TRANNO = '" & Null2String(rsivan!RRNO) & "' "
                    gconDMIS.Execute (sqlcommands)
                    
                    sqlcommands = "Update " & Xrr_HD & " set Status = 'C', "
                    sqlcommands = sqlcommands + "LASTUPDATE = '" & LOGDATE & "', "
                    sqlcommands = sqlcommands + "USERCODE = " & N2Str2Null(LOGCODE) & " "
                    sqlcommands = sqlcommands + " where [TYPE] = '" & rsTMP!Type & "' and  RIV_TRANNO = '" & rsTMP!TRANNO & "' and CLASSCODE = 'RRV' and STATUS = 'P' "
                    gconDMIS.Execute (sqlcommands)
                    
                    Dim rsclcttrec                      As New ADODB.Recordset
                    Set rsclcttrec = gconDMIS.Execute("Select stock_ord , type,tranqty from " & Xrr_Detail & " where [TYPE] = '" & rsivan!Type & "' and trantype = 'RR' and tranno = '" & Null2String(rsivan!RRNO) & "' order by itemno asc")
                    If Not (rsclcttrec.EOF And rsclcttrec.BOF) Then
                        rsclcttrec.MoveFirst
                        Do While Not rsclcttrec.EOF
                            gconDMIS.Execute ("Update PMIS_STOCKMAS set trecqty = isnull(trecqty,0) - " & N2Str2Zero(rsclcttrec!TRANQTY) & " where [TYPE] = '" & Null2String(rsclcttrec!Type) & "' and stockno = '" & Null2String(rsclcttrec!STOCK_ORD) & "'")
                            rsclcttrec.MoveNext
                        Loop
                    End If
                
            End If
        rsTMP.MoveNext
    Loop
End If
    CancelRecieving = True
    Set rsTMP = Nothing
    Set rsclcttrec = Nothing
    Exit Function
ErrorCode:
    CancelRecieving = False
End Function
Sub Show_req_parts(ro As String)
       Dim SQLTXT                  As String
        Dim rsTMP                   As New ADODB.Recordset
        Dim REC                     As XtremeReportControl.ReportRecord
        Dim itemno                  As Integer
        Dim TOTALQTY                As Integer
        Dim XXX                     As String
        
        itemno = 1
        
        SQLTXT = "SELECT * FROM CSMS_RETURN_DET A INNER JOIN PMIS_ALLDAYTRAN B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.ITEMID = B.ID WHERE A.REP_OR = '" & ro & "' and ID_HD = '" & LABID_HD & "' and b.TRANTYPE='RIV'" & vbCrLf
        
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        
        RcReq_parts.Records.DeleteAll
        
        If Not (rsTMP.BOF And rsTMP.BOF) Then
            Do While Not rsTMP.EOF
            Set REC = RcReq_parts.Records.Add
            REC.AddItem Trim(itemno)
            REC.AddItem Null2String(rsTMP!STOCKNO)
            REC.AddItem Null2String(rsTMP!STOCK_TYPE)
            REC.AddItem N2Str2Zero(rsTMP!TRANQTY)
            REC.AddItem N2Str2Zero(rsTMP!QTY_REQ)
            REC.AddItem Null2String(rsTMP!TRANNO)
            REC.AddItem Null2Date(rsTMP!trandate)
            REC.AddItem Null2String(rsTMP!ID_RDET)
            
            itemno = itemno + 1
            rsTMP.MoveNext
            Loop
        End If
        RcReq_parts.Populate
        Set rsTMP = Nothing
End Sub

Sub configure_reportcontrol()
    With RcParts
        .Columns.DeleteAll
        .Columns.Add 0, "PARTNUMBER", 7, True: .Columns(0).Alignment = xtpAlignmentIconLeft: .Columns(0).BestFit: .Columns(0).Resizable = False: .Columns(0).AllowRemove = False
        .Columns.Add 1, "PART DESCRIPTION", 10, True: .Columns(1).Alignment = xtpAlignmentIconLeft: .Columns(1).BestFit: .Columns(1).Resizable = False: .Columns(1).AllowRemove = False
        .Columns.Add 2, "TYPE", 3, True: .Columns(2).Alignment = xtpAlignmentCenter: .Columns(2).BestFit: .Columns(2).Resizable = False: .Columns(2).AllowRemove = False
        .Columns.Add 3, "ISS QTY", 3, True: .Columns(3).Alignment = xtpAlignmentCenter: .Columns(3).BestFit: .Columns(3).Resizable = False: .Columns(3).AllowRemove = False
        .Columns.Add 4, "RET QTY", 4, True: .Columns(4).Alignment = xtpAlignmentCenter: .Columns(4).BestFit: .Columns(4).Resizable = False: .Columns(4).AllowRemove = False
        .Columns.Add 5, "PIS#", 5, True: .Columns(5).Alignment = xtpAlignmentCenter: .Columns(5).BestFit: .Columns(5).Resizable = False: .Columns(5).AllowRemove = False
        .Columns.Add 6, "TRANDATE", 5, True: .Columns(6).Alignment = xtpAlignmentCenter: .Columns(6).BestFit: .Columns(6).Resizable = False: .Columns(6).AllowRemove = False
        .Columns.Add 7, "ID", 0, True: .Columns(7).Alignment = xtpAlignmentCenter: .Columns(7).BestFit: .Columns(7).Resizable = False: .Columns(7).AllowRemove = False
        .Columns.Add 8, "ID_DET", 0, True: .Columns(8).Alignment = xtpAlignmentCenter: .Columns(8).BestFit: .Columns(8).Resizable = False: .Columns(8).AllowRemove = False

        
        .PaintManager.HorizontalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbBSDot
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomDrawItem
        .PaintManager.ColumnStyle = XPCOLOR_3DFACE 'XPCOLOR_3DFACE
        .PaintManager.CaptionFont.Bold = False
        .PaintManager.TextFont.Bold = False
    End With
    
    With RcReq_parts
        .Columns.DeleteAll
        .Columns.Add 0, "ITEMNO", 7, True: .Columns(0).Alignment = xtpAlignmentIconRight: .Columns(0).BestFit: .Columns(0).Resizable = False: .Columns(0).AllowRemove = False
        .Columns.Add 1, "PART NUMBER", 10, True: .Columns(1).Alignment = xtpAlignmentIconLeft: .Columns(1).BestFit: .Columns(1).Resizable = False: .Columns(1).AllowRemove = False
        .Columns.Add 2, "TYPE", 3, True: .Columns(2).Alignment = xtpAlignmentCenter: .Columns(2).BestFit: .Columns(2).Resizable = False: .Columns(2).AllowRemove = False
        .Columns.Add 3, "ISS QTY", 3, True: .Columns(3).Alignment = xtpAlignmentCenter: .Columns(3).BestFit: .Columns(3).Resizable = False: .Columns(3).AllowRemove = False
        .Columns.Add 4, "RET QTY", 4, True: .Columns(4).Alignment = xtpAlignmentCenter: .Columns(4).BestFit: .Columns(4).Resizable = False: .Columns(4).AllowRemove = False
        .Columns.Add 5, "PIS#", 5, True: .Columns(5).Alignment = xtpAlignmentCenter: .Columns(5).BestFit: .Columns(5).Resizable = False: .Columns(5).AllowRemove = False
        .Columns.Add 6, "TRANDATE", 5, True: .Columns(6).Alignment = xtpAlignmentCenter: .Columns(6).BestFit: .Columns(6).Resizable = False: .Columns(6).AllowRemove = False
        .Columns.Add 7, "ID_DET", 0, True: .Columns(7).Alignment = xtpAlignmentCenter: .Columns(7).BestFit: .Columns(7).Resizable = False: .Columns(7).AllowRemove = False

        
        .PaintManager.HorizontalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonGraphical
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomDrawItem
        .PaintManager.ColumnStyle = XPCOLOR_3DFACE
        .PaintManager.CaptionFont.Bold = False
        .PaintManager.TextFont.Bold = False
    End With
    
    With rcFind
        .Columns.DeleteAll
        .Columns.Add 0, "REPAIR ORDER", 10, True: .Columns(0).Alignment = xtpAlignmentCenter: .Columns(0).BestFit: .Columns(0).Resizable = False: .Columns(0).AllowRemove = False
        .Columns.Add 1, "ID", 0, True: .Columns(1).Alignment = xtpAlignmentCenter: .Columns(1).BestFit: .Columns(1).Resizable = False: .Columns(1).AllowRemove = False
    
        .PaintManager.HorizontalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSolid    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonGraphical
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomDrawItem
        .PaintManager.ColumnStyle = XPCOLOR_3DFACE
        .PaintManager.CaptionFont.Bold = False
        .PaintManager.TextFont.Bold = False
    End With

End Sub
Sub show_allparts(RONO As String)
        Dim SQLTXT                  As String
        Dim rsTMP                   As ADODB.Recordset
        Dim rsreturn                As ADODB.Recordset
        Dim REC                     As XtremeReportControl.ReportRecord
        Dim RSTEMP                  As ADODB.Recordset
        Dim UNION_QUERY             As String
        Dim i                       As Integer
        
        
               UNION_QUERY = ("SELECT T.TRANNO,T.TRANDATE,T.RONO,T.STOCK_ORD,T.STOCKDESC,T.[TYPE],T.TRANQTY,T.ID, sum(Y.[QTY_REQ])as QTY_REQ FROM( " & _
                                     "SELECT B.TRANNO,B.TRANDATE,B.RONO,B.STOCK_ORD,A.STOCKDESC,A.[TYPE],B.TRANQTY,B.ID FROM PMIS_STOCKMAS A INNER JOIN " & _
                                     "(SELECT A.TRANNO,B.TRANDATE,B.RONO,A.STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,A.ID  " & _
                                     "FROM PMIS_DAYTRAN A INNER JOIN PMIS_ORD_HIST B ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] " & _
                                     "AND A.TRANTYPE = B.TRANTYPE WHERE  B.STATUS IN ('P','B') AND B.TRANTYPE IN('RIV','ADB')" & _
                                     "Union All " & _
                                     "SELECT A.TRANNO,B.TRANDATE,B.RONO,A.STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,A.ID  " & _
                                     "FROM PMIS_TDAYTRAN A INNER JOIN PMIS_ORD_HD B ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] " & _
                                     "AND A.TRANTYPE = B.TRANTYPE WHERE B.STATUS IN ('P','B') AND B.TRANTYPE IN('RIV','ADB')) B " & _
                                     "ON LTRIM(RTRIM(A.STOCKNO)) = LTRIM(RTRIM(B.STOCK_ORD)))T LEFT OUTER JOIN CSMS_RETURN_DET Y " & _
                                     "ON LTRIM(RTRIM(T.STOCK_ORD)) = LTRIM(RTRIM(Y.STOCKNO)) AND T.RONO = Y.REP_OR and T.ID = Y.ITEMID " & _
                                     "WHERE T.RONO = '" & RONO & "' group BY T.TRANNO,T.TRANDATE,T.RONO,T.STOCK_ORD,T.STOCKDESC,T.[TYPE],T.TRANQTY,T.ID order by T.TRANNO")
                                     'where B.ID NOT IN(SELECT ITEMID FROM CSMS_RETURN_DET WHERE REP_OR = '" & RONO & "')
    
    Set rsTMP = gconDMIS.Execute(UNION_QUERY)
        RcParts.Records.DeleteAll
        
        Valid = False
        If Not (rsTMP.BOF And rsTMP.EOF) Then
        Valid = True
           Do While Not rsTMP.EOF
             Set REC = RcParts.Records.Add
                    
                    REC.AddItem Null2String(rsTMP!STOCK_ORD)
                    REC.AddItem Null2String(rsTMP!STOCKDESC)
                    REC.AddItem Null2String(rsTMP![Type])
                    REC.AddItem Null2String(rsTMP!TRANQTY)
                    REC.AddItem N2Str2IntZero(rsTMP!QTY_REQ)
                    REC.AddItem Null2String(rsTMP!TRANNO)
                    REC.AddItem Null2String(rsTMP!trandate)
                    REC.AddItem Null2String(rsTMP!ID)
                    'REC.AddItem Null2String(rsTMP!ID_RDET)
          rsTMP.MoveNext
          Loop
        End If
        RcParts.Populate
    Set rsTMP = Nothing
End Sub
Function checkIfReturnParts(XXX As String)
    Dim Rsrono As ADODB.Recordset
    Dim xRsrono As ADODB.Recordset
    Dim XPART, xID As String
        Set Rsrono = gconDMIS.Execute("SELECT * FROM CSMS_RETURN_DET A inner join CSMS_RETURN_HD B on A.rep_or = B.rep_or " & _
                                      "where A.id_hd = B.id and  B.status is not null and B.rep_or =" & N2Str2Null(XXX) & "")
        If Not Rsrono.EOF And Not Rsrono.BOF Then
              If Null2String(Rsrono!Status) = "" Then
                     checkIfReturnParts = 0
               Else
                    checkIfReturnParts = 1
                End If
        End If
End Function

Private Sub cmdCancelSupp_Click()
    txtFindSupp.Text = ""
    frameSupp.ZOrder 1
    frameSupp.Visible = False
    What_Func = False
    PIC_MAIN.ZOrder 0
    PIC_MAIN.Enabled = True
End Sub

Private Sub cmdDelete_Click()
        Dim SQLTXT                  As String
        
        If Function_Access(LOGID, "Acess_Delete", LOCALACESS) = False Then Exit Sub

        If MsgBox("Do you want to Delete this Transaction", vbQuestion + vbYesNo) = vbYes Then
            SQLTXT = "DELETE FROM CSMS_RETURN_HD WHERE ID = '" & LABID_HD & "'"
            Call gconDMIS.Execute(SQLTXT)
            rsRefresh
            StoreMemVars
            Call Command2_Click
            ShowDeletedMsg
            Call NEW_LogAudit("X", "PARTS RETURN TRANSACTION", "DELETE FROM CSMS_RETURN_HD WHERE ID = '" & LABID_HD & "'", LABID_HD, "", "CODE: " & LABID_HD, "", "")

        Else
            Exit Sub
        End If
        picRetrn.Visible = False
        
        Call ShowTranNo
End Sub

Private Sub cmdDelete_det_Click()
        Dim SQLTXT As String
        On Error GoTo ErrorCode:
        
        SQLTXT = "DELETE FROM CSMS_RETURN_DET WHERE ID_rdet = '" & lblid & "' AND ID_HD = '" & LABID_HD & "'"
        Call gconDMIS.Execute(SQLTXT)
        
        rsRefresh
        RS_RETURN.Find "id = " & LABID_HD
        StoreMemVars
        Call Command2_Click
        
        Call NEW_LogAudit("XX", "PARTS RETURN TRANSACTION", SQLTXT, LABID_HD, "", "CODE: " & LABID_HD, "", "")

ErrorCode:
        Exit Sub
End Sub

Private Sub cmdEdit_Click()
        If Function_Access(LOGID, "Acess_Edit", LOCALACESS) = False Then Exit Sub

        ADDOREDIT_HD = "EDIT"
        rcFind.Enabled = False
        picsave.ZOrder 0
        picsave.Visible = True
        picAdd.ZOrder 1
        picAdd.Visible = False
        On Error Resume Next
        txt_req_by.SetFocus
        Picture2.Enabled = True
        txtRemarks.Enabled = True
        txt_req_by.Enabled = True
        txtverify.Enabled = False
End Sub

Private Sub cmdExit_Click()
        Unload Me
End Sub

 Function checkttlreturn(XXX As String)
    Dim rscheckonhand                                   As ADODB.Recordset
    Dim sqlcommand                                      As String

    Set rscheckonhand = New ADODB.Recordset
    sqlcommand = "SELECT A.REP_OR,A.STATUS,B.STOCKNO,B.ITEMID,B.STOCK_TYPE,sum(B.QTY_REQ)as QTY_REQ " & _
                 "FROM CSMS_RETURN_HD A INNER JOIN CSMS_RETURN_DET B  ON A.ID = B.ID_HD WHERE A.STATUS = 'P'" & _
                 "AND A.VERI_BY IS NOT NULL AND B.ITEMID='" & XXX & "'" & _
                 "group by A.REP_OR,A.STATUS,B.STOCKNO,B.ITEMID,B.STOCK_TYPE order by  QTY_REQ"

    Set rscheckonhand = gconDMIS.Execute(sqlcommand)
    If Not (rscheckonhand.EOF And rscheckonhand.BOF) Then
        checkttlreturn = N2Str2Zero(rscheckonhand!QTY_REQ)
    Else
        checkttlreturn = "0"
    End If
    Set rscheckonhand = Nothing
End Function
 
Private Sub cmdOk_det_Click()
        Dim SQLTXT As String
        On Error GoTo ErrorCode:
        
        
        If N2Str2IntZero(txtTran_return) = 0 Or N2Str2IntZero(txtTran_return) > N2Str2IntZero(txtTran_issued) Then
            MsgBox "Return Qty Is Invalid Entry!", vbInformation
            Exit Sub
        End If

         If (checkttlreturn(lblid) + N2Str2IntZero(txtTran_return.Text)) > N2Str2IntZero(txtTran_issued.Text) Then
            MsgBox "Cannot Receive more than Issued Qty!", vbInformation
            Exit Sub
        End If
        

        
        PIC_MAIN.Enabled = True
        pic_Returnpart.ZOrder 0
        PIC_SEARCH.Enabled = True
        pic_Returnpart.Visible = True
        
        
        If ADDOREDIT_DET = "ADD" Then
        
            SQLTXT = "INSERT INTO CSMS_RETURN_DET" & vbCrLf
            SQLTXT = SQLTXT & "(REP_OR,STOCKNO,STOCK_TYPE,QTY_ISS,QTY_REQ,ITEMID,ID_HD,TRANTYPE)" & vbCrLf
            SQLTXT = SQLTXT & "VALUES('" & txtRep_or & "','" & cbo_Tran_Partnumber & "','" & txttype & "'," & vbCrLf
            SQLTXT = SQLTXT & "'" & txtTran_issued & "','" & txtTran_return & "','" & lblid & "','" & LABID_HD & "','" & lbltrantype & "')"
            
        Else
        
            SQLTXT = "UPDATE CSMS_RETURN_DET SET " & vbCrLf
            SQLTXT = SQLTXT & "REP_OR = '" & txtRep_or & "'," & vbCrLf
            SQLTXT = SQLTXT & "STOCKNO = '" & cbo_Tran_Partnumber & "'," & vbCrLf
            SQLTXT = SQLTXT & "STOCK_TYPE = '" & txttype & "'," & vbCrLf
            SQLTXT = SQLTXT & "QTY_ISS = '" & txtTran_issued & "'," & vbCrLf
            SQLTXT = SQLTXT & "QTY_REQ = '" & txtTran_return & "'," & vbCrLf
            SQLTXT = SQLTXT & "ITEMID = '" & lblid & "'," & vbCrLf
            SQLTXT = SQLTXT & "ID_HD = '" & LABID_HD & "'," & vbCrLf
            SQLTXT = SQLTXT & "TRANTYPE = '" & lbltrantype & "'"
            SQLTXT = SQLTXT & "WHERE ID_RDET = '" & LABID_DET & "'"
        End If
        
        Call gconDMIS.Execute(SQLTXT)
        LOAD_DATA = False
        SQLTXT = ""
        rsRefresh
        RS_RETURN.Find "ID = " & LABID_HD
        StoreMemVars
        RcReq_parts.Populate
        rcFind.Enabled = True
        txtsearch.Enabled = True
        'Show_req_parts (txtRep_or)
        Command2_Click
        If ADDOREDIT_DET = "ADD" Then
             Call NEW_LogAudit("AA", "PARTS RETURN TRANSACTION", SQLTXT, lblid.Caption, "", "", "RRV", "")
        Else
             Call NEW_LogAudit("EE", "PARTS RETURN TRANSACTION", SQLTXT, lblid.Caption, "", "", "RRV", "")
        End If
        
ErrorCode:
    'MsgBox Err.Number
        Select Case err.Number
           Case -2147217873
                MsgBox "Please Save first the repair order info! ", vbInformation
        
            Case Else
                'DO THING
        End Select
        
        
        Exit Sub
End Sub
Function If_HD_Have_DET(ro As String, ID As String) As Boolean
        Dim SQLTXT As String
        Dim rsTMP As New ADODB.Recordset
        
        SQLTXT = "SELECT * FROM CSMS_RETURN_DET WHERE REP_OR = '" & ro & "' AND ID_HD = '" & ID & "'"
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        
        If Not (rsTMP.BOF And rsTMP.EOF) Then
            If_HD_Have_DET = True
        Else
            If_HD_Have_DET = False
        End If
        Set rsTMP = Nothing
End Function
Function If_Item_Exits_inCsmS_det(ID As String, itemno As String) As Boolean
        Dim SQLTXT As String
        Dim rsTMP As New ADODB.Recordset
        
        SQLTXT = "SELECT * FROM CSMS_RETURN_det WHERE  ID_HD = '" & ID & "' AND  ITEMID = '" & itemno & "'"
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        
        If Not (rsTMP.BOF And rsTMP.EOF) Then
            If_Item_Exits_inCsmS_det = True
        Else
           If_Item_Exits_inCsmS_det = False
        End If
        Set rsTMP = Nothing
End Function

Private Sub cmdPost_Click()
      Dim s As String
          
      If Function_Access(LOGID, "Acess_Post", LOCALACESS) = False Then Exit Sub
    
      s = "You cannot Post transaction!" & vbCrLf
      s = s & "Select Parts to be Return!"
      
      If If_HD_Have_DET(txtRep_or, LABID_HD) = False Then
          MsgBox s, vbInformation
          Exit Sub
      End If
      
      'Call Display_Repor(txtRep_or)
      If MsgBox("Do you want to Post this Transaction", vbQuestion + vbYesNo) = vbYes Then
            rcFind.Enabled = True
            Call gconDMIS.Execute("UPDATE CSMS_RETURN_HD SET STATUS = 'P' WHERE ID = '" & LABID_HD & "'")
            rsRefresh
            RS_RETURN.Find "ID = " & LABID_HD
            StoreMemVars
            txtRep_or.Enabled = False
            
            Call NEW_LogAudit("P", "PARTS RETURN TRANSACTION", "UPDATE CSMS_RETURN_HD SET STATUS = 'P' WHERE ID = '" & LABID_HD & "'", lblid.Caption, "", "", "RRV", "")

       Else
            Exit Sub
       End If
    
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LOCALACESS) = False Then Exit Sub

    On Error GoTo ErrorCode:

    Screen.MousePointer = 11
    rptReturn.WindowTitle = "RETURN PARTS"
    rptReturn.Formulas(0) = "Company_Name = '" & COMPANY_NAME & "'"
    rptReturn.Formulas(1) = "Company_Address = '" & COMPANY_ADDRESS & "'"
    
    PrintSQLReport rptReturn, CSMS_REPORT_PATH & "return_parts.rpt", "{CSMS_RETURN_HD.ID} =  " & N2Str2IntZero(LABID_HD), DMIS_REPORT_Connection, 1

    Screen.MousePointer = 0
    
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSaveSupp_Click()
     Call SaveSetting("DMIS 2.0", "RETURN FROM SERVICE", "SUPP_CODE", txtSupCode)
    
    MessagePop InfoFriend, "Selected Supplier", "Selected Supplier has been Saved."
    frameSupp.ZOrder 1
    frameSupp.Visible = False
    cmdCancelSupp.Value = True
    Exit Sub
End Sub

Private Sub cmdTranCancel_Click()
        picAdd.ZOrder 0
        picAdd.Visible = True
        txtRep_or.Enabled = False
        rcFind.Enabled = True
        'rsRefresh
        StoreMemVars
        txtRemarks.Enabled = False
        txt_req_by.Enabled = False
        picRetrn.Visible = False
        PIC_SEARCH.Enabled = True
End Sub
Private Sub cmdTranSave_Click()
        Dim SQLTXT As String
        
        On Error GoTo ErrorCode:
       
        If txt_req_by = "" Then
            MsgBox "Request By must not be Blank!", vbInformation
            txt_req_by.SetFocus
            Exit Sub
        End If
               
        If txtRep_or = "" Then
            MsgBox "Repair order cannot be blank!", vbInformation
            txtRep_or.SetFocus
            Exit Sub
        End If

        If ADDOREDIT_HD = "ADD" Then
            SQLTXT = "INSERT INTO CSMS_RETURN_HD (REP_OR,DATE_REQ,REQ_BY,REMARKS)" & vbCrLf
            SQLTXT = SQLTXT & "VALUES('" & txtRep_or & "','" & txtRtnDate & "'," & vbCrLf
            SQLTXT = SQLTXT & "'" & txt_req_by & "','" & txtRemarks & "')"
        Else
            SQLTXT = "UPDATE CSMS_RETURN_HD SET "
            SQLTXT = SQLTXT & "REP_OR ='" & txtRep_or & "'," & vbCrLf
            SQLTXT = SQLTXT & "DATE_REQ = '" & txtRtnDate & "'," & vbCrLf
            SQLTXT = SQLTXT & "REQ_BY = '" & txt_req_by & "'," & vbCrLf
            SQLTXT = SQLTXT & "REMARKS ='" & txtRemarks & "' " & vbCrLf
            SQLTXT = SQLTXT & " WHERE ID = '" & LABID_HD & "'" & vbCrLf
        End If
        
        LOAD_DATA = False
        Call gconDMIS.Execute(SQLTXT)
        Call rsRefresh
        RS_RETURN.Find "id = " & LABID_HD
        Call StoreMemVars
        cmdTranCancel.Value = True
        cmdDelete.Enabled = True
        Call ShowTranNo
        txt_req_by.Enabled = False
        txtRemarks.Enabled = False
        'rcParts.Enabled = True
        PIC_SEARCH.Enabled = True
        If ADDOREDIT_HD = "ADD" Then
            Call NEW_LogAudit("A", "PARTS RETURN TRANSACTION", SQLTXT, lblid.Caption, "", "", "RRV", "")
        Else
            Call NEW_LogAudit("E", "PARTS RETURN TRANSACTION", SQLTXT, lblid.Caption, "", "", "RRV", "")
        End If
ErrorCode:
        Exit Sub
End Sub
Sub ShowTranNo()
        Dim SQLTXT                  As String
        Dim rsTMP                   As New ADODB.Recordset
        Dim REC                     As XtremeReportControl.ReportRecord
        
        SQLTXT = "SELECT * FROM CSMS_RETURN_HD"
        Set rsTMP = gconDMIS.Execute(SQLTXT)
         
        rcFind.Records.DeleteAll
         
        If Not (rsTMP.BOF And rsTMP.EOF) Then
            Do While Not rsTMP.EOF
                Set REC = rcFind.Records.Add
                REC.AddItem Null2String(rsTMP!REP_OR)
                REC.AddItem N2Str2IntZero(rsTMP!ID)
            
            rsTMP.MoveNext
            Loop
        rcFind.Populate
        End If
            
        Set rsTMP = Nothing
End Sub

'        Dim CMD As New ADODB.Command
'        With CMD
'            .ActiveConnection = gconDMIS
'            .CommandType = adCmdStoredProc
'            '.Prepared = True
'            .CommandText = "SP_CSMS_RETURN_HD"
'
'        .Parameters.Append .CreateParameter("@REP_OR", adVarChar, adParamInput, 50, txtRep_or)
'        .Parameters.Append .CreateParameter("@DATE_REQ", adDate, adParamInput, 50, txtRtnDate)
'        .Parameters.Append .CreateParameter("@REQ_BY", adVarChar, adParamInput, 50, txt_req_by)
'        .Parameters.Append .CreateParameter("@REMARKS", adVarChar, adParamInput, 150, txtremarks)
'        .Parameters.Append .CreateParameter("@IDX", adInteger, adParamInputOutput, , 0)
'
'        .Execute
'
'        End With
'End Sub

'Private Sub cmdVerify_Click()
       ' picVerify.Visible = True
        'txtverify.SetFocus
'End Sub

Private Sub cmdUnPost_Click()
        Dim Msg As String
        
    If Function_Access(LOGID, "Acess_UNPost", LOCALACESS) = False Then Exit Sub
            
        
        rsRefresh
        RS_RETURN.Find "ID = " & LABID_HD
        StoreMemVars
        
        If txtverify <> "" Then
            
            Msg = "You cannot Unpost the Transaction" & vbCrLf
            Msg = Msg & "Its already Verified by Parts Department!"
            
            MsgBox Msg, vbInformation
        Else
            If MsgBox("Do you want to Unpost this Transaction", vbQuestion + vbYesNo) = vbYes Then
                Call gconDMIS.Execute("UPDATE CSMS_RETURN_HD SET STATUS = 'N' WHERE ID = '" & LABID_HD & "'")
                rsRefresh
                RS_RETURN.Find "ID = " & LABID_HD
                StoreMemVars
                Call NEW_LogAudit("U", "PARTS RETURN TRANSACTION", "UPDATE CSMS_RETURN_HD SET STATUS = 'N' WHERE ID = '" & LABID_HD & "'", lblid.Caption, "", "", "RRV", "")
            Else
                Exit Sub
            End If
        End If
   
        
End Sub

Private Sub cmdverify_Click()
       If Function_Access(LOGID, "Acess_Post", "PARTS RETURN TRANSACTION") = False Then Exit Sub
        
        Dim lng As Long
        Dim Msg As String
        Dim SQLMSG As String
        Dim CTR As String
        
        On Error GoTo errocode
        
        
        If CheckIfRoIsAlreadyInvoice(txtRep_or.Text) = True Then
            MsgBox "Repair Order is Already Billed. Transaction will not be imported!", vbInformation, txtRep_or.Text & " Already Billed"
            Exit Sub
        End If
        SQLMSG = "System Error Due to Proccess on - @09099386087." & vbCrLf
        SQLMSG = SQLMSG & "This will RollBack Transaction!" & vbCrLf
        SQLMSG = SQLMSG & "Please conctact Netspeed Software Helpdesk immediately!" & vbCrLf
        SQLMSG = SQLMSG & "Thank you!"
         
        If COMPANY_CODE = "HPT" Then
            str_catcher = "H00062"
        ElseIf COMPANY_CODE = "HPC" Then
            str_catcher = "H00008"
        ElseIf COMPANY_CODE = "HMH" Then
            str_catcher = "H00019"
        Else
            str_catcher = GetSetting("DMIS 2.0", "RETURN FROM SERVICE", "SUPP_CODE", "")
        End If
        
        'update by: IEBV
        'description:
        If RTrim(LTrim(str_catcher)) = "" Then MessagePop InfoFriend, "Action Void", "No Supplier Selected": Call Label15_Click: Exit Sub
        '----------------------------------------------------------------------------------------------------------------------------------
        
        If txtverify.Text = "" Then
            MsgBox "Please Put who verify this Transaction! ", vbInformation
            txtverify.Enabled = True
            txtverify.SetFocus
            prgExcelGen.Visible = False
            cmdverify.Enabled = True
            Exit Sub
        Else
           If MsgBox("Do you want to received this Partnumber(s)", vbQuestion + vbYesNo) = vbYes Then
                
                lng = 8
                If lng > 0 Then
                    prgExcelGen.Max = lng
                    prgExcelGen.Value = 0
                End If
                
                gconDMIS.BeginTrans
                cmdverify.Enabled = False
                prgExcelGen.ZOrder 0
                prgExcelGen.Visible = True
                prgExcelGen.Value = prgExcelGen.Value + 1
                prgExcelGen.Text = "Creating Receiving Entry ... "
                Call gconDMIS.Execute("Update CSMS_RETURN_HD SET VERI_BY = '" & txtverify & "' WHERE REP_OR = '" & txtRep_or & "' AND STATUS = 'P'")
            Else
                Exit Sub
            End If
        End If
        
        If ImportDetails(txtRep_or, "P", "2") = False Then
            SQLMSG = Replace(SQLMSG, "@09099386087", "Importing Process")
            MsgBox SQLMSG, vbCritical, "Processing Error"
            Screen.MousePointer = 0
            gconDMIS.RollbackTrans
            prgExcelGen.Visible = False
            cmdverify.Enabled = True
            Exit Sub
        Else
            prgExcelGen.Value = prgExcelGen.Value + 1
            prgExcelGen.Text = "Creating Receiving Entry ... "
        End If
            
        If ImportDetails(txtRep_or, "M", "3") = False Then
            SQLMSG = Replace(SQLMSG, "@09099386087", "Importing Process")
            MsgBox SQLMSG, vbCritical, "Processing Error"
            Screen.MousePointer = 0
            gconDMIS.RollbackTrans
            prgExcelGen.Visible = False
            cmdverify.Enabled = True
            Exit Sub
        Else
            prgExcelGen.Value = prgExcelGen.Value + 1
            prgExcelGen.Text = "Creating Receiving Entry ... "
        End If
            
        If ImportDetails(txtRep_or, "A", "4") = False Then
            SQLMSG = Replace(SQLMSG, "@09099386087", "Importing Process")
            MsgBox SQLMSG, vbCritical, "Processing Error"
            Screen.MousePointer = 0
            gconDMIS.RollbackTrans
            prgExcelGen.Visible = False
            cmdverify.Enabled = True
            Exit Sub
        Else
            prgExcelGen.Value = prgExcelGen.Value + 1
            prgExcelGen.Text = "Creating Receiving Entry ... "
        End If
'        Call ImportDetails(txtRep_or, "P", "2"): prgExcelGen.Value = prgExcelGen.Value + 1
'        Call ImportDetails(txtRep_or, "M", "3"): prgExcelGen.Value = prgExcelGen.Value + 1
'        Call ImportDetails(txtRep_or, "A", "4"): prgExcelGen.Value = prgExcelGen.Value + 1
        'Call DELETE_PARTS_IN_CSMS_RO_DET(txtRep_or):  prgExcelGen.Value = prgExcelGen.Value + 1
        'Call INSERT_PARTS_IN_CSMS_RO_DET(txtRep_or): prgExcelGen.Value = prgExcelGen.Value + 1
        'Call UPDATE_CSMS_RO_DET_LINE_SEQUENTIALLY(txtRep_or):prgExcelGen.Value = prgExcelGen.Value + 1
        
'       If GETSUP_CODE(COMPANY_CODE & "001") = False Then
'            If INSERT_SUPPLIER = False Then
'                SQLMSG = Replace(SQLMSG, "@09099386087", "INSERTING SUPPLIER")
'                MsgBox SQLMSG, vbCritical, "Processing Error"
'                Screen.MousePointer = 0
'                gconDMIS.RollbackTrans
'                prgExcelGen.Visible = False
'                cmdverify.Enabled = True
'                Exit Sub
'            Else
'                prgExcelGen.Value = prgExcelGen.Value + 1
'                prgExcelGen.Text = "Creating Receiving Entry ... "
'
'            End If
'        End If

        'Call INSERT_RECEIVING_ENTRY_IN_PARTS(txtRep_or) : prgExcelGen.Value = prgExcelGen.Value + 1
'         ctr = (gconDMIS.Execute("select count(*) from csms_return_det where ID_HD ='" & LABID_HD & "' and QTY_APP = 1 ").Fields(0).Value)
'updated by: IEBV   02172011_0650pm
'description:   to avoid double recieving
'         If ctr > 0 Then
         ' do nothing
'         Else
            If CREATE_RR_HEADER(txtRep_or, str_catcher) = False Then
                SQLMSG = Replace(SQLMSG, "@09099386087", "CREATING RR HEADER")
                MsgBox SQLMSG, vbCritical, "Processing Error"
                Screen.MousePointer = 0
                gconDMIS.RollbackTrans
                prgExcelGen.Visible = False
                cmdverify.Enabled = True
                Exit Sub
            Else
                 prgExcelGen.Value = prgExcelGen.Value + 1
                 prgExcelGen.Text = "Creating Receiving Entry ... "
            End If
            
            If CREATE_RR_DETAILS(txtRep_or) = False Then
                SQLMSG = Replace(SQLMSG, "@09099386087", "CREATING RR DETAILS")
                MsgBox SQLMSG, vbCritical, "Processing Error"
                Screen.MousePointer = 0
                gconDMIS.RollbackTrans
                prgExcelGen.Visible = False
                cmdverify.Enabled = True
                Exit Sub
            Else
                prgExcelGen.Value = prgExcelGen.Value + 1
                prgExcelGen.Text = "Creating Receiving Entry ... "
            End If
'        End If
'--------------------------------------------------------------------------------------------
'        If UPDATE_COLUMN_ONHAND_RECEIPTS_TRECQTY = False Then
'            SQLMSG = Replace(SQLMSG, "@09099386087", "UPDATING MASTERFILE . . . ")
'            MsgBox SQLMSG, vbCritical, "Processing Error"
'            Screen.MousePointer = 0
'            gconDMIS.RollbackTrans
'            prgExcelGen.Visible = False
'            cmdverify.Enabled = True
'            Exit Sub
'        Else
'            prgExcelGen.Value = prgExcelGen.Value + 1
'            prgExcelGen.Text = "Updating Qty Onhand ... "
'        End If
        If UPDATE_COLUMN_ONHAND_PARTSMASTERFILE = False Then
            SQLMSG = Replace(SQLMSG, "@09099386087", "UPDATING ONHAND IN MASTERFILE")
            MsgBox SQLMSG, vbCritical, "Processing Error"
            Screen.MousePointer = 0
            gconDMIS.RollbackTrans
            prgExcelGen.Visible = False
            cmdverify.Enabled = True
            Exit Sub
        Else
            prgExcelGen.Value = prgExcelGen.Value + 1
            prgExcelGen.Text = "Updating Qty Onhand ... "
        End If
        
        If UPDATE_COLUMN_RECEIPTS_PARTSMASTERFILE = False Then
            SQLMSG = Replace(SQLMSG, "@09099386087", "UPDATING TOTAL RECEIPTS IN MASTERFILE")
            MsgBox SQLMSG, vbCritical, "Processing Error"
            Screen.MousePointer = 0
            gconDMIS.RollbackTrans
            prgExcelGen.Visible = False
            cmdverify.Enabled = True
            Exit Sub
        Else
            prgExcelGen.Value = prgExcelGen.Value + 1
            prgExcelGen.Text = "Updating Receipts Entry ... "
        End If
        
        prgExcelGen.Value = prgExcelGen.Value + 1
        prgExcelGen.Text = " (100% Completed)"
        prgExcelGen.Visible = False
        
        Call rsRefresh
        RS_RETURN.Find "ID = " & LABID_HD
        Call StoreMemVars
        Call ShowTranNo
       
        Call NEW_LogAudit("R", "PARTS RETURN TRANSACTION", "", lblid.Caption, "", "", "RRV", "")

        gconDMIS.CommitTrans
        MsgBox "Already Received!", vbInformation
       
       
      Exit Sub
      
errocode:
    MsgBox err.Description
End Sub


Private Sub cmdview_Click()
        pic_Select.ZOrder 0
        pic_Select.Visible = True
        'Call DisplayItems
End Sub

Private Sub Command2_Click()
        pic_Returnpart.ZOrder 0
        pic_Returnpart.Visible = False
        picAdd.Enabled = True
End Sub

Private Sub Command3_Click()
    PIC_MAIN.ZOrder 0
    PIC_MAIN.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
          Case vbKeyF12
'                frameSupp.ZOrder 0
'                frameSupp.Visible = True
'                Call DisplaySupp(What_Func)
          Case vbKeyF1
                Call Label15_Click
                Call DisplaySupp(What_Func)
    End Select
End Sub

Private Sub Form_Load()
        CenterMe frmMain, Me, 1
        txtRtnDate.Text = Date
        picsave.Visible = False: picAdd.Visible = True
        LABID_HD = GetLastData()
        rsRefresh
        RS_RETURN.Find "ID = " & LABID_HD
        StoreMemVars
        ShowTranNo
        configure_reportcontrol
        'Picture2.Enabled = False
        ADDOREDIT_HD = "ADD"
        ADDOREDIT_DET = "ADD"
        LOCALACESS = "PARTS RETURN TRANSACTION"
        prgExcelGen.Visible = False
        prgExcelGen.ZOrder 1
       
End Sub
Function GetLastData() As String
        Dim SQLTXT As String
        Dim rsTMP As New ADODB.Recordset
        
        SQLTXT = "SELECT ISNULL(MAX(ID),0) AS ID FROM CSMS_RETURN_HD"
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        If Not (rsTMP.BOF And rsTMP.EOF) Then
            GetLastData = Null2String(rsTMP!ID)
        End If
        
        Set rsTMP = Nothing
End Function



Private Sub Label15_Click()

    frameSupp.Visible = True
    frameSupp.ZOrder 0
    frameSupp.Enabled = True
    PIC_MAIN.ZOrder 1
    PIC_MAIN.Enabled = False
    str_catcher = GetSetting("DMIS 2.0", "RETURN FROM SERVICE", "SUPP_CODE", "")
    Dim RS  As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute("Select NAMEOFVENDOR from ALL_VENDOR where Code = '" & str_catcher & "'")
    If Not (RS.EOF And RS.BOF) Then
        txtFindSupp.Text = Null2String(RS!nameofvendor)
    Else
        txtFindSupp.Text = ""
    End If
End Sub

Private Sub lvwSupp_DblClick()
    cmdSaveSupp.Value = True
End Sub

Private Sub lvwSupp_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Dim ITEM As ListItem
    Dim SQLTXT As String
    Dim rsTMP As New ADODB.Recordset
    Dim Xcode As String
    
    Xcode = lvwSupp.SelectedItem.Text
    SQLTXT = "SELECT CODE,NAMEOFVENDOR FROM ALL_VENDOR WHERE CODE= '" & Null2String(Xcode) & "'"
    Set rsTMP = gconDMIS.Execute(SQLTXT)
    
    If Not (rsTMP.EOF And rsTMP.BOF) Then
        txtSupCode = Null2String(rsTMP!Code)
        txtSuppNAME = Null2String(rsTMP!nameofvendor)
    End If
    
    Set rsTMP = Nothing
End Sub


Private Sub pic_Returnpart_KeyUp(KeyCode As Integer, Shift As Integer)
         Select Case KeyCode
            Case vbKeyEscape
                pic_Returnpart.ZOrder 1
                pic_Returnpart.Visible = False
        End Select
End Sub

Function CheckIfthisMonthIssue(RONO As String, ID As Long) As Boolean
        Dim SQLTXT As String
        Dim rsTMP As New ADODB.Recordset
        
        SQLTXT = "SELECT * FROM (SELECT ID FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT B.TRANDATE,B.RONO,B.STOCK_ORD,A.STOCKDESC,A.[TYPE],B.TRANQTY,B.ID FROM PMIS_STOCKMAS A INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT B.TRANDATE,B.RONO,A.STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,A.ID FROM PMIS_DAYTRAN A INNER JOIN PMIS_ORD_HIST B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & "WHERE  B.STATUS IN ('P','B') AND B.TRANTYPE IN('RIV','ADB')" & vbCrLf
        SQLTXT = SQLTXT & "Union" & vbCrLf
        SQLTXT = SQLTXT & "SELECT B.TRANDATE,B.RONO,A.STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,A.ID FROM PMIS_TDAYTRAN A INNER JOIN PMIS_ORD_HD B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & "WHERE B.STATUS IN ('P','B') AND B.TRANTYPE IN('RIV','ADB')" & vbCrLf
        SQLTXT = SQLTXT & ") B" & vbCrLf
        SQLTXT = SQLTXT & "ON LTRIM(RTRIM(A.STOCKNO)) = LTRIM(RTRIM(B.STOCK_ORD))" & vbCrLf
        SQLTXT = SQLTXT & ")" & vbCrLf
        SQLTXT = SQLTXT & "T  WHERE T.ID = '" & ID & "'" & vbCrLf
        SQLTXT = SQLTXT & ") Y  WHERE Y.ID IN (SELECT ID FROM PMIS_TDAYTRAN)" & vbCrLf
        
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        
        If Not (rsTMP.EOF And rsTMP.BOF) Then
            'If (Null2String(Month((RSTMP!TRANDATE))) <> Null2String(Month((Now)))) Then
                CheckIfthisMonthIssue = True
            'Else
        Else
                CheckIfthisMonthIssue = False
            'End If
        End If
        
    
        Set rsTMP = Nothing
End Function
Private Sub ReturnedParts(xID As String, ro As String)
        Dim SQLTXT                  As String
        Dim rsTMP                   As New ADODB.Recordset
        
'        SQLTXT = "SELECT T.TRANNO,T.TRANDATE,T.RONO,T.STOCK_ORD,T.STOCKDESC,T.[TYPE], (T.TRANQTY - Y.QTY_REQ) AS TRANQTY , T.ID FROM" & vbCrLf
'        SQLTXT = SQLTXT & "(" & vbCrLf
'        SQLTXT = SQLTXT & "SELECT B.TRANNO,B.TRANDATE,B.RONO,B.STOCK_ORD,A.STOCKDESC,A.[TYPE],B.TRANQTY,B.ID FROM PMIS_STOCKMAS A INNER JOIN" & vbCrLf
'        SQLTXT = SQLTXT & "(SELECT A.TRANNO,B.TRANDATE,B.RONO,A.STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,A.ID FROM PMIS_DAYTRAN A INNER JOIN PMIS_ORD_HIST B" & vbCrLf
'        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE  B.STATUS IN ('P','B') AND B.TRANTYPE IN('RIV','ADB')" & vbCrLf
'        SQLTXT = SQLTXT & "Union ALL" & vbCrLf
'        SQLTXT = SQLTXT & "SELECT A.TRANNO,B.TRANDATE,B.RONO,A.STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,A.ID FROM PMIS_TDAYTRAN A INNER JOIN PMIS_ORD_HD B" & vbCrLf
'        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE B.STATUS IN ('P','B') AND B.TRANTYPE IN('RIV','ADB')" & vbCrLf
'        SQLTXT = SQLTXT & ") B" & vbCrLf
'        SQLTXT = SQLTXT & "ON LTRIM(RTRIM(A.STOCKNO)) = LTRIM(RTRIM(B.STOCK_ORD))" & vbCrLf
'        SQLTXT = SQLTXT & ")T" & vbCrLf
'        SQLTXT = SQLTXT & "LEFT OUTER JOIN CSMS_RETURN_DET Y" & vbCrLf
'        SQLTXT = SQLTXT & "ON LTRIM(RTRIM(T.STOCK_ORD)) = LTRIM(RTRIM(Y.STOCKNO)) AND T.RONO = Y.REP_OR and T.ID = Y.ITEMID" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE  T.RONO = '" & ro & "' and T.ID  = '" & xID & "'  ORDER BY T.TRANNO" & vbCrLf

        SQLTXT = "SELECT T.TRANDATE,T.RONO,T.STOCK_ORD,T.STOCKDESC,T.[TYPE], T.TRANQTY, T.ID FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT B.TRANDATE,B.RONO,B.STOCK_ORD,A.STOCKDESC,A.[TYPE],B.TRANQTY,B.ID FROM PMIS_STOCKMAS A INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT B.TRANDATE,B.RONO,A.STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,A.ID FROM PMIS_DAYTRAN A INNER JOIN PMIS_ORD_HIST B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & "WHERE  B.STATUS IN ('P','B') AND B.TRANTYPE IN('RIV','ADB')" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT B.TRANDATE,B.RONO,A.STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,A.ID FROM PMIS_TDAYTRAN A INNER JOIN PMIS_ORD_HD B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & "WHERE B.STATUS IN ('P','B') AND B.TRANTYPE IN('RIV','ADB')" & vbCrLf
        SQLTXT = SQLTXT & ") B" & vbCrLf
        SQLTXT = SQLTXT & "ON LTRIM(RTRIM(A.STOCKNO)) = LTRIM(RTRIM(B.STOCK_ORD))  " & vbCrLf
        SQLTXT = SQLTXT & ")T WHERE T.ID  = '" & xID & "' AND T.RONO = '" & ro & "'" & vbCrLf

        'WHERE B.ID NOT IN(SELECT ITEMID FROM CSMS_RETURN_DET WHERE REP_OR = '" & ro & "')
    '    SQLTXT = "SELECT * FROM" & vbCrLf
    '    SQLTXT = SQLTXT & "(" & vbCrLf
    '    SQLTXT = SQLTXT & "SELECT B.RONO,B.STOCK_ORD,A.STOCKDESC,A.[TYPE],B.TRANQTY,B.ID FROM PMIS_STOCKMAS A INNER JOIN" & vbCrLf
    '    SQLTXT = SQLTXT & "(SELECT B.RONO,A.STOCK_ORD,A.TRANQTY,B.STATUS,A.ID FROM PMIS_DAYTRAN A INNER JOIN PMIS_ORD_HIST B" & vbCrLf
    '    SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
    '    SQLTXT = SQLTXT & " WHERE B.RONO IS NOT NULL AND B.STATUS = 'P') B" & vbCrLf
    '    SQLTXT = SQLTXT & "ON A.STOCKNO = B.STOCK_ORD" & vbCrLf
    '    SQLTXT = SQLTXT & ")" & vbCrLf
    '    SQLTXT = SQLTXT & "T  LEFT OUTER JOIN CSMS_RETURN_DET Y" & vbCrLf
    '    SQLTXT = SQLTXT & " ON T.STOCK_ORD = Y.STOCKNO AND T.RONO = Y.REP_OR" & vbCrLf
    '    SQLTXT = SQLTXT & "WHERE T.ID= '" & xID & "' ORDER BY T.[TYPE]" & vbCrLf
        
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        If Not (rsTMP.BOF And rsTMP.EOF) Then
            cbo_Tran_Partnumber.Text = Null2String(rsTMP!STOCK_ORD)
            txtTran_part_Desc.Text = Null2String(rsTMP!STOCKDESC)
            txtTran_issued.Text = Null2String(rsTMP!TRANQTY)
            txttype = Null2String(rsTMP![Type])
            txtTran_return.Text = Null2String(rsTMP!TRANQTY)
            'If N2Str2Zero(RSTMP!QTY_REQ) = 0 Then
                'txtTran_return.Text = 0
            'Else
                'txtTran_return = N2Str2Zero(RSTMP!QTY_REQ)
                
            'End If
           
        End If
        Set rsTMP = Nothing
    
End Sub
Function UPDATE_COLUMN_ONHAND_RECEIPTS_TRECQTY() As Boolean
        Dim SQLTXT As String
        
        On Error GoTo ErrorCode

        SQLTXT = "UPDATE PMIS_STOCKMAS SET RECEIPTS = B.T_RECEIPTS,TRECQTY = B.T_TRECQTY,ONHAND = B.TOTAL_ONHAND FROM" & vbCrLf
        SQLTXT = SQLTXT & "PMIS_STOCKMAS A INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCKNO,STOCKDESC,[TYPE],T_RECEIPTS,T_TRECQTY,TOTAL_ONHAND FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCKNO,STOCKDESC,[TYPE]," & vbCrLf
        SQLTXT = SQLTXT & "ISNULL((" & vbCrLf
        SQLTXT = SQLTXT & "SELECT SUM(ISNULL(TRANQTY,0))  FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],ISNULL(TRANQTY,0) AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE TRANTYPE = 'RR' AND STATUS IN ('P','B')" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],ISNULL(TRANQTY,0) AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE TRANTYPE = 'ADJ' AND STATUS IN ('P','B') AND IN_OUT = 'I'" & vbCrLf
        SQLTXT = SQLTXT & ") T  WHERE T.STOCK_ORD = LTRIM(RTRIM(A.STOCKNO)) AND T.[TYPE] = A.[TYPE] GROUP BY STOCK_ORD ,[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "),0) AS T_RECEIPTS," & vbCrLf
        SQLTXT = SQLTXT & "ISNULL((" & vbCrLf
        SQLTXT = SQLTXT & "SELECT SUM(ISNULL(TRANQTY,0))  FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],ISNULL(TRANQTY,0) AS TRANQTY FROM PMIS_TDAYTRAN WHERE TRANTYPE = 'RR' AND STATUS IN ('P','B')" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],ISNULL(TRANQTY,0) AS TRANQTY FROM PMIS_TDAYTRAN WHERE TRANTYPE = 'ADJ' AND STATUS IN ('P','B') AND IN_OUT = 'I'" & vbCrLf
        SQLTXT = SQLTXT & ") T  WHERE T.STOCK_ORD = LTRIM(RTRIM(A.STOCKNO)) AND T.[TYPE] = A.[TYPE] GROUP BY STOCK_ORD ,[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "),0) AS T_TRECQTY," & vbCrLf
        SQLTXT = SQLTXT & "ISNULL((SELECT SUM(TRANQTY) AS TOTAL_ONHAND FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'RR' AS TRANTYPE,1 * SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE TRANTYPE = 'RR' AND STATUS IN ('P','B')" & vbCrLf
        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD ,[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'BEG' AS TRANTYPE,1 * SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE TRANTYPE = 'BEG' AND STATUS IN ('P','B') AND IN_OUT = 'I'" & vbCrLf
        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD ,[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'ADJ-IN' AS TRANTYPE,1 * SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE TRANTYPE = 'ADJ' AND STATUS IN ('P','B') AND IN_OUT = 'I'" & vbCrLf
        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD,[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'ADJ-OUT' AS TRANTYPE,-1 * SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE TRANTYPE = 'ADJ' AND STATUS IN ('P','B') AND IN_OUT = 'O'" & vbCrLf
        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD,[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'ISS' AS TRANTYPE,-1 * SUM(ISNULL(TRANQTY,0))AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN ('RIV','DR','CSH','CHG') AND STATUS IN ('P','B') AND IN_OUT = 'O'" & vbCrLf
        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD,[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & ")T WHERE T.STOCK_ORD = A.STOCKNO AND T.[TYPE] = A.[TYPE] GROUP BY STOCK_ORD,[TYPE]),0) AS TOTAL_ONHAND" & vbCrLf
        SQLTXT = SQLTXT & "FROM PMIS_STOCKMAS A" & vbCrLf
        SQLTXT = SQLTXT & ") B ) B ON A.STOCKNO = B.STOCKNO AND A.[TYPE] = B.[TYPE] WHERE ISNULL(A.ONHAND,0) <> B.TOTAL_ONHAND OR ISNULL(A.RECEIPTS,0) <> B.T_RECEIPTS OR ISNULL(A.TRECQTY,0) <> B.T_TRECQTY" & vbCrLf
      
        Call gconDMIS.Execute(SQLTXT)

        UPDATE_COLUMN_ONHAND_RECEIPTS_TRECQTY = True
        Exit Function
        
ErrorCode:
        UPDATE_COLUMN_ONHAND_RECEIPTS_TRECQTY = False
End Function
Function UPDATE_COLUMN_ONHAND_PARTSMASTERFILE() As Boolean
        Dim SQLTXT As String
        Dim cmd                                            As ADODB.Command
        Dim ictx(2)                                        As String
        Dim i                                              As Integer
        
        On Error GoTo ErrorCode
        
'        SQLTXT = "UPDATE PMIS_STOCKMAS SET ONHAND = B.TOTAL_ONHAND FROM PMIS_STOCKMAS A INNER JOIN (" & vbCrLf
'        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],SUM(TRANQTY) AS TOTAL_ONHAND FROM" & vbCrLf
'        SQLTXT = SQLTXT & "(" & vbCrLf
'        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'RR' AS TRANTYPE,1 * SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM PMIS_ALLDAYTRAN" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE TRANTYPE = 'RR' AND STATUS IN ('P','B')" & vbCrLf
'        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD ,[TYPE]" & vbCrLf
'        SQLTXT = SQLTXT & "Union All" & vbCrLf
'        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'BEG' AS TRANTYPE,1 * SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM PMIS_ALLDAYTRAN" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE TRANTYPE = 'BEG' AND STATUS IN ('P','B') AND IN_OUT = 'I'" & vbCrLf
'        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD ,[TYPE]" & vbCrLf
'        SQLTXT = SQLTXT & "Union All" & vbCrLf
'        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'ADJ-IN' AS TRANTYPE,1 * SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM PMIS_ALLDAYTRAN" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE TRANTYPE = 'ADJ' AND STATUS IN ('P','B') AND IN_OUT = 'I'" & vbCrLf
'        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD,[TYPE]" & vbCrLf
'        SQLTXT = SQLTXT & "Union All" & vbCrLf
'        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'ADJ-OUT' AS TRANTYPE,-1 * SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM PMIS_ALLDAYTRAN" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE TRANTYPE = 'ADJ' AND STATUS IN ('P','B') AND IN_OUT = 'O'" & vbCrLf
'        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD,[TYPE]" & vbCrLf
'        SQLTXT = SQLTXT & "Union All" & vbCrLf
'        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],'ISS' AS TRANTYPE,-1 * SUM(ISNULL(TRANQTY,0))AS TRANQTY FROM PMIS_ALLDAYTRAN" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE TRANTYPE IN ('RIV','DR','CSH','CHG') AND STATUS IN ('P','B') AND IN_OUT = 'O'" & vbCrLf
'        SQLTXT = SQLTXT & "GROUP BY STOCK_ORD,[TYPE]" & vbCrLf
'        SQLTXT = SQLTXT & ")T GROUP BY STOCK_ORD,[TYPE]" & vbCrLf
'        SQLTXT = SQLTXT & ") B ON LTRIM(RTRIM(A.STOCKNO)) = LTRIM(RTRIM(B.STOCK_ORD)) AND A.[TYPE] = B.[TYPE]" & vbCrLf
'        SQLTXT = SQLTXT & "Where IsNull(ONHAND, 0) <> IsNull(TOTAL_ONHAND, 0)" & vbCrLf
'
'        Call gconDMIS.Execute(SQLTXT)
        
        ictx(0) = "P": ictx(1) = "M": ictx(2) = "A"
        For i = 0 To 2
            Set cmd = New ADODB.Command
            With cmd
                .NamedParameters = True
                .CommandType = adCmdStoredProc
                .CommandText = "SP_PMIS_BALANCED_STOCK"
                .ActiveConnection = gconDMIS
                .CommandTimeout = 1000
                .Parameters.Append .CreateParameter("@TYPE", adVarChar, adParamInput, 50, ictx(i))
            End With
            cmd.Execute
        Next i
        UPDATE_COLUMN_ONHAND_PARTSMASTERFILE = True
        Exit Function
        
ErrorCode:
        UPDATE_COLUMN_ONHAND_PARTSMASTERFILE = False
End Function
Function GETSUP_CODE(Code As String) As Boolean
        Dim SQLTXT As String
        Dim rsTMP As New ADODB.Recordset
    
        SQLTXT = "SELECT * FROM ALL_VENDOR_TABLE WHERE CODE = '" & Code & "'"
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        If Not (rsTMP.BOF And rsTMP.EOF) Then
            GETSUP_CODE = True
        Else
            GETSUP_CODE = False
        End If
End Function
Function CREATE_RR_HEADER(ro As String, Supp_Code As String) As Boolean
            Dim SQLTXT As String
           
        On Error GoTo ErrorCode
        
        SQLTXT = "DECLARE @RRNO_LENGTH AS INT" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @SUPPCODE    AS NVARCHAR(10)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @RRNO_P      AS NVARCHAR(10)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @RRNO_M      AS NVARCHAR(10)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @RRNO_A      AS NVARCHAR(10)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @SUPPNAME    AS NVARCHAR(50)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @USERCODE    AS NVARCHAR(5)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @INVOICE AS NVARCHAR(20)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @HEAD_ID AS INT" & vbCrLf
        SQLTXT = SQLTXT & "SET @USERCODE = " & N2Str2Null(LOGCODE) & "" & vbCrLf
        SQLTXT = SQLTXT & "SET @HEAD_ID = '" & LABID_HD & "'" & vbCrLf
        SQLTXT = SQLTXT & "SET @RRNO_P = (SELECT MAX(ISNULL(RRNO,0)) + 1 AS RRNO FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT MAX(CASE" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) < 6 THEN REPLICATE('0',1) + ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) = 6 THEN ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "END) AS RRNO FROM PMIS_RR_HD WHERE [TYPE] = 'P' and isnumeric(rrno) = 1 " & vbCrLf
        SQLTXT = SQLTXT & "Union" & vbCrLf
        SQLTXT = SQLTXT & "SELECT MAX(CASE" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) < 6 THEN REPLICATE('0',1) + ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) = 6 THEN ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "END) AS RRNO  FROM PMIS_REC_HIST WHERE [TYPE] = 'P' and isnumeric(rrno) = 1" & vbCrLf
        SQLTXT = SQLTXT & ")T)" & vbCrLf
        SQLTXT = SQLTXT & "SET @RRNO_M = (SELECT MAX(ISNULL(RRNO,0)) + 1 AS RRNO FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT MAX(CASE" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) < 6 THEN REPLICATE('0',1) + ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) = 6 THEN ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "END) AS RRNO FROM PMIS_RR_HD WHERE [TYPE] = 'M' and isnumeric(rrno) = 1" & vbCrLf
        SQLTXT = SQLTXT & "Union" & vbCrLf
        SQLTXT = SQLTXT & "SELECT MAX(CASE" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) < 6 THEN REPLICATE('0',1) + ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) = 6 THEN ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "END) AS RRNO  FROM PMIS_REC_HIST WHERE [TYPE] = 'M' and isnumeric(rrno) = 1" & vbCrLf
        SQLTXT = SQLTXT & ")T)" & vbCrLf
        SQLTXT = SQLTXT & "SET @RRNO_A = (SELECT MAX(ISNULL(RRNO,0)) + 1 AS RRNO FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT MAX(CASE" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) < 6 THEN REPLICATE('0',1) + ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) = 6 THEN ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "END) AS RRNO FROM PMIS_RR_HD WHERE [TYPE] = 'A' and isnumeric(rrno) = 1" & vbCrLf
        SQLTXT = SQLTXT & "Union" & vbCrLf
        SQLTXT = SQLTXT & "SELECT MAX(CASE" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) < 6 THEN REPLICATE('0',1) + ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(ISNULL(RRNO,0)) = 6 THEN ISNULL(RRNO,0)" & vbCrLf
        SQLTXT = SQLTXT & "END) AS RRNO  FROM PMIS_REC_HIST WHERE [TYPE] = 'A' and isnumeric(rrno) = 1" & vbCrLf
        SQLTXT = SQLTXT & ")T)" & vbCrLf
        SQLTXT = SQLTXT & "SET @RRNO_LENGTH = (SELECT MAX(RRNO) FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT MAX(LEN(ISNULL(RRNO,0))) AS RRNO FROM PMIS_RR_HD A WHERE isnumeric(rrno) = 1" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT MAX(LEN(ISNULL(RRNO,0))) AS RRNO FROM PMIS_REC_HIST A where isnumeric(rrno) = 1" & vbCrLf
        SQLTXT = SQLTXT & ") T )" & vbCrLf
        SQLTXT = SQLTXT & "SET @SUPPCODE = '" & Supp_Code & "'" & vbCrLf
        SQLTXT = SQLTXT & "SET @SUPPNAME = (SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = @SUPPCODE)" & vbCrLf
        SQLTXT = SQLTXT & "SET @INVOICE = (SELECT RIGHT(REP_OR,6)  FROM CSMS_REPOR WHERE REP_OR = '" & ro & "' )" & vbCrLf
        SQLTXT = SQLTXT & "INSERT INTO PMIS_RR_HD ([TYPE],RRNO,RRDATE,RIV_TRANNO,RECVD_CODE,RECVD_FROM,CLASSCODE,TERMS," & vbCrLf
        SQLTXT = SQLTXT & "TTLRRAMT,  NETRRAMT,STATUS,LASTUPDATE,REMARKS,INVNO,USERCODE)" & vbCrLf
        SQLTXT = SQLTXT & "SELECT [TYPE],ISNULL(RRNO,0) AS RRNO,RRDATE,CAST(LEFT(DBO.CONCAT(ISNULL(RRNO,0),RONO,@HEAD_ID),6) AS NVARCHAR(100)) AS RIV_TRANNO,RECVD_CODE,RECVD_FROM,CLASSCODE,TERMS," & vbCrLf
        'SQLTXT = SQLTXT & "CAST(SUM(TTLRRAMT) AS DECIMAL(18,2)) AS TTLRRAMT, CAST(SUM(NETRRAMT) AS DECIMAL(18,2)) AS NETRRAMT," & vbCrLf
        SQLTXT = SQLTXT & "0 AS TTLRRAMT,0 AS NETRRAMT, " & vbCrLf
        SQLTXT = SQLTXT & "STATUS,LASTUPDATE,CAST('REFERENCE PIS #' + ' ' + DBO.CONCAT(RRNO,RONO,@HEAD_ID) AS NVARCHAR(100)) AS REMARKS,INVNO,USERCODE" & vbCrLf
        SQLTXT = SQLTXT & "FROM (" & vbCrLf
        SQLTXT = SQLTXT & "SELECT [TYPE] ,CASE [TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'P' THEN (SELECT  RIGHT(REPLICATE('0',@RRNO_LENGTH) + CAST(@RRNO_P AS varchar(10)),@RRNO_LENGTH))" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'M' THEN (SELECT  RIGHT(REPLICATE('0',@RRNO_LENGTH) + CAST(@RRNO_M AS varchar(10)),@RRNO_LENGTH))" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'A' THEN (SELECT  RIGHT(REPLICATE('0',@RRNO_LENGTH) + CAST(@RRNO_A AS varchar(10)),@RRNO_LENGTH))" & vbCrLf
        SQLTXT = SQLTXT & "END AS RRNO," & vbCrLf
        SQLTXT = SQLTXT & "CONVERT(VARCHAR(10), GETDATE(),110) AS RRDATE,Case [Type]" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'P' THEN ITEMID" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'M' THEN ITEMID" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'A' THEN ITEMID" & vbCrLf
        SQLTXT = SQLTXT & "END AS RIV_TRANNO," & vbCrLf
        SQLTXT = SQLTXT & "@SUPPCODE AS RECVD_CODE,@SUPPNAME AS RECVD_FROM,'RRV' AS CLASSCODE, 0 AS TERMS," & vbCrLf
        SQLTXT = SQLTXT & "CAST(QTY_REQ * S_MAC AS DECIMAL(18,2)) AS TTLRRAMT," & vbCrLf
        SQLTXT = SQLTXT & "CAST(QTY_REQ * S_MAC AS DECIMAL(18,2)) AS NETRRAMT," & vbCrLf
        SQLTXT = SQLTXT & "'P' AS STATUS,CONVERT(VARCHAR(10), GETDATE(),110) AS LASTUPDATE, REMARKS AS REMARKS,RONO,@INVOICE AS INVNO,@USERCODE AS USERCODE" & vbCrLf
        SQLTXT = SQLTXT & "FROM (" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,SI_TYPE,TRANUPRICE,S_MAC,RONO,STOCKDESC,[TYPE],TRANQTY,MAC,TRANTYPE FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT SI_TYPE,TRANUPRICE,S_MAC,RONO,STOCK_ORD,STOCKDESC,[TYPE],SUM(TRANQTY) AS TRANQTY," & vbCrLf
        SQLTXT = SQLTXT & "CAST((SUM(MAC)/(COUNT(STOCK_ORD))) AS DECIMAL(18,2)) AS MAC,TRANTYPE FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT B.SI_TYPE,B.TRANUPRICE,B.RONO,ISNULL(A.MAC,0) AS S_MAC,B.STOCK_ORD,A.STOCKDESC,A.[TYPE],B.TRANQTY,B.MAC,B.TRANTYPE FROM PMIS_STOCKMAS A INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT B.SI_TYPE,ISNULL(A.TRANUPRICE,0) AS TRANUPRICE,B.RONO,LTRIM(RTRIM(A.STOCK_ORD)) AS STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,ISNULL(A.MAC,0) AS MAC,A.TRANTYPE FROM PMIS_DAYTRAN A INNER JOIN PMIS_ORD_HIST B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & "WHERE B.RONO IS NOT NULL AND B.STATUS in('P','B') AND A.TRANTYPE IN ('RIV','ADB')) B" & vbCrLf
        SQLTXT = SQLTXT & "ON LTRIM(RTRIM(A.STOCKNO)) = LTRIM(RTRIM(B.STOCK_ORD))" & vbCrLf
        SQLTXT = SQLTXT & ")" & vbCrLf
        SQLTXT = SQLTXT & "G GROUP BY STOCK_ORD,RONO,STOCKDESC,[TYPE], SI_TYPE,TRANUPRICE,S_MAC,TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & ") T WHERE  STOCK_ORD IN(SELECT LTRIM(RTRIM(STOCKNO)) FROM CSMS_RETURN_DET)" & vbCrLf
        SQLTXT = SQLTXT & ") Y INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT A.REMARKS, B.STOCK_TYPE,A.REP_OR,B.STOCKNO, ITEMID,B.QTY_REQ,B.ID_HD  FROM CSMS_RETURN_HD A INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT ID_HD,STOCK_TYPE,REP_OR,LTRIM(RTRIM(STOCKNO)) AS STOCKNO,SUM(ISNULL(QTY_REQ,0)) AS QTY_REQ," & vbCrLf
        SQLTXT = SQLTXT & "ITEMID FROM CSMS_RETURN_DET WHERE TRANTYPE = 'RIV'" & vbCrLf
        SQLTXT = SQLTXT & "GROUP BY REP_OR,STOCKNO,ID_HD,STOCK_TYPE,ITEMID )" & vbCrLf
        SQLTXT = SQLTXT & "B ON A.ID = B.ID_HD AND A.REP_OR = B.REP_OR) X" & vbCrLf
        SQLTXT = SQLTXT & "ON Y.RONO = X.REP_OR AND Y.STOCK_ORD = X.STOCKNO AND Y.[TYPE] = X.[STOCK_TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "WHERE X.REP_OR = '" & ro & "' AND TRANTYPE = 'RIV' AND ID_HD = '" & LABID_HD & "' " & vbCrLf
        SQLTXT = SQLTXT & ") T" & vbCrLf
        SQLTXT = SQLTXT & "GROUP BY [TYPE],RRNO,RRDATE,RECVD_CODE,RECVD_FROM,CLASSCODE,TERMS, STATUS,LASTUPDATE,REMARKS,RONO,INVNO,USERCODE" & vbCrLf
        

        Call gconDMIS.Execute(SQLTXT)
        
        CREATE_RR_HEADER = True
        Exit Function
        
ErrorCode:
        CREATE_RR_HEADER = False

End Function
Function CREATE_RR_DETAILS(ro As String) As Boolean
        Dim SQLTXT              As String
        Dim rsTMP               As New ADODB.Recordset
        Dim TRAN_MAC            As Double
        Dim M_MAC               As Double
        Dim ITEM_NO             As String
        Dim XSTOCK_ORD          As String
        Dim XTYPE               As String
        Dim xTranDate           As Date
        Dim xTRANTYPE           As String
        Dim XTRANNO             As String
        Dim XSTOCK_SUP          As String
        Dim XTRANQTY            As Long
        Dim Xstatus             As String
        Dim XIN_OUT             As String
        Dim XUSERCODE           As String
        Dim XTREMARKS           As String
        Dim XNON_HARI           As String
        Dim XLASTUPDATE         As Date
        Dim XTRANUCOST          As Double
        Dim P_TTLRRAMT          As Double
        Dim M_TTLRRAMT          As Double
        Dim A_TTLRRAMT          As Double
        Dim P_TRANNO            As String
        Dim M_TRANNO            As String
        Dim A_TRANNO            As String
        
        
        ITEM_NO = 0: TRAN_MAC = 0: M_MAC = 0: XSTOCK_ORD = "": XTYPE = "": xTRANTYPE = "":
        XTRANNO = "": XSTOCK_SUP = "": XTRANQTY = 0: Xstatus = "": XIN_OUT = "": XUSERCODE = "": XTREMARKS = "": XNON_HARI = "":
        XTRANUCOST = 0: P_TTLRRAMT = 0: M_TTLRRAMT = 0: A_TTLRRAMT = 0: P_TRANNO = "": M_TRANNO = "": A_TRANNO = "":
        SQLTXT = "":
        
        On Error GoTo ErrorCode
        
        SQLTXT = "DECLARE @RRNO_LENGTH AS INT" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @RRNO_P      AS NVARCHAR(10)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @RRNO_M      AS NVARCHAR(10)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @RRNO_A      AS NVARCHAR(10)" & vbCrLf
        SQLTXT = SQLTXT & "DECLARE @USERCODE     AS NVARCHAR(10)" & vbCrLf
        SQLTXT = SQLTXT & "SET @RRNO_P = (SELECT CAST(MAX(RRNO)  AS NVARCHAR) FROM PMIS_RR_HD WHERE [TYPE] = 'P' AND CLASSCODE = 'RRV')" & vbCrLf
        SQLTXT = SQLTXT & "SET @RRNO_M = (SELECT CAST(MAX(RRNO)  AS NVARCHAR) FROM PMIS_RR_HD WHERE [TYPE] = 'M' AND CLASSCODE = 'RRV')" & vbCrLf
        SQLTXT = SQLTXT & "SET @RRNO_A = (SELECT CAST(MAX(RRNO)  AS NVARCHAR) FROM PMIS_RR_HD WHERE [TYPE] = 'A' AND CLASSCODE = 'RRV')" & vbCrLf
        SQLTXT = SQLTXT & "SET @RRNO_LENGTH = (SELECT DISTINCT(LEN(RRNO)) FROM PMIS_RR_HD)" & vbCrLf
        SQLTXT = SQLTXT & "SET @USERCODE = " & N2Str2Null(LOGCODE) & "" & vbCrLf
'        SQLTXT = SQLTXT & "INSERT INTO PMIS_TDAYTRAN(STOCK_ORD,[TYPE],TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_SUP,TRANQTY,TRANUCOST,STATUS,IN_OUT,MAC," & vbCrLf
'        SQLTXT = SQLTXT & "TRANINVAMT,USERCODE,LASTUPDATE,TREMARKS,NON_HARI)" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE] ,CONVERT(VARCHAR(10), GETDATE(),110) AS TRANDATE,'RR' AS TRANTYPE," & vbCrLf
        SQLTXT = SQLTXT & "Case [Type]" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'P' THEN (SELECT  RIGHT(REPLICATE('0',@RRNO_LENGTH) + CAST(@RRNO_P AS varchar(10)),@RRNO_LENGTH))" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'M' THEN (SELECT  RIGHT(REPLICATE('0',@RRNO_LENGTH) + CAST(@RRNO_M AS varchar(10)),@RRNO_LENGTH))" & vbCrLf
        SQLTXT = SQLTXT & "WHEN 'A' THEN (SELECT  RIGHT(REPLICATE('0',@RRNO_LENGTH) + CAST(@RRNO_A AS varchar(10)),@RRNO_LENGTH))" & vbCrLf
        SQLTXT = SQLTXT & "END AS TRANNO," & vbCrLf
        SQLTXT = SQLTXT & "'0001'  AS ITEMNO,STOCK_ORD AS STOCK_SUP, QTY_REQ AS TRANQTY,TRANUCOST AS TRANUCOST," & vbCrLf
        SQLTXT = SQLTXT & "'P' AS STATUS,'I' AS IN_OUT,S_MAC AS MAC,S_MAC AS TRANINVAMT," & vbCrLf
        SQLTXT = SQLTXT & "@USERCODE AS USERCODE,CONVERT(VARCHAR(10), GETDATE(),110) AS LASTUPDATE," & vbCrLf
        SQLTXT = SQLTXT & "'Verified' as TREMARKS,NON_HARI" & vbCrLf
        SQLTXT = SQLTXT & "FROM (" & vbCrLf
        SQLTXT = SQLTXT & "SELECT NON_HARI,STOCK_ORD,SI_TYPE,S_MAC,RONO,STOCKDESC,[TYPE],TRANQTY,TRANUCOST,TRANTYPE FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT NON_HARI,SI_TYPE,S_MAC,RONO,STOCK_ORD,STOCKDESC,[TYPE],SUM(TRANQTY) AS TRANQTY," & vbCrLf
        SQLTXT = SQLTXT & "CAST((SUM(TRANUCOST)/(COUNT(STOCK_ORD))) AS DECIMAL(18,2)) AS TRANUCOST,TRANTYPE FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT NON_HARI,B.SI_TYPE,B.TRANUPRICE,B.TRANTYPE,B.RONO,ISNULL(A.MAC,0) AS S_MAC,B.STOCK_ORD,A.STOCKDESC,A.[TYPE],B.TRANQTY,B.TRANUCOST FROM PMIS_STOCKMAS A INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT B.SI_TYPE,ISNULL(A.TRANUPRICE,0) AS TRANUPRICE,B.RONO,LTRIM(RTRIM(A.STOCK_ORD)) AS STOCK_ORD,ISNULL(A.TRANQTY,0) AS TRANQTY,B.STATUS,ISNULL(A.TRANUCOST,0) AS TRANUCOST,A.TRANTYPE FROM PMIS_DAYTRAN A INNER JOIN PMIS_ORD_HIST B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & "WHERE B.RONO IS NOT NULL AND B.STATUS IN ('P','B') AND A.TRANTYPE IN ('RIV','ADB')) B" & vbCrLf
        SQLTXT = SQLTXT & "ON LTRIM(RTRIM(A.STOCKNO)) = LTRIM(RTRIM(B.STOCK_ORD))" & vbCrLf
        SQLTXT = SQLTXT & ")" & vbCrLf
        SQLTXT = SQLTXT & "G GROUP BY NON_HARI,STOCK_ORD,RONO,STOCKDESC,[TYPE], SI_TYPE,S_MAC,TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & ") T WHERE  STOCK_ORD IN(SELECT LTRIM(RTRIM(STOCKNO)) FROM CSMS_RETURN_DET)" & vbCrLf
        SQLTXT = SQLTXT & ") Y INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT A.REMARKS, B.STOCK_TYPE,A.REP_OR,B.STOCKNO,B.QTY_REQ,B.ID_HD  FROM CSMS_RETURN_HD A INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT ID_HD,STOCK_TYPE,REP_OR,LTRIM(RTRIM(STOCKNO)) AS STOCKNO,SUM(ISNULL(QTY_REQ,0)) AS QTY_REQ" & vbCrLf
        SQLTXT = SQLTXT & "From CSMS_RETURN_DET WHERE TRANTYPE = 'RIV'" & vbCrLf
        SQLTXT = SQLTXT & "GROUP BY REP_OR,STOCKNO,ID_HD,STOCK_TYPE )" & vbCrLf
        SQLTXT = SQLTXT & "B ON A.ID = B.ID_HD AND A.REP_OR = B.REP_OR) X" & vbCrLf
        SQLTXT = SQLTXT & "ON Y.RONO = X.REP_OR AND Y.STOCK_ORD = X.STOCKNO AND Y.[TYPE] = X.[STOCK_TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "WHERE X.REP_OR = '" & ro & "' AND Y.TRANTYPE = 'RIV' AND ID_HD = '" & LABID_HD & "' " & vbCrLf
'        SQLTXT = SQLTXT & "DECLARE @ID AS INT" & vbCrLf
'        SQLTXT = SQLTXT & "UPDATE PMIS_TDAYTRAN SET ITEMNO = '1' WHERE TRANNO = @RRNO_P AND [TYPE] = 'P' AND TRANTYPE = 'RR'" & vbCrLf
'        SQLTXT = SQLTXT & "SET @ID = 0" & vbCrLf
'        SQLTXT = SQLTXT & "UPDATE PMIS_TDAYTRAN SET ITEMNO =" & vbCrLf
'        SQLTXT = SQLTXT & "CASE WHEN LEN(@ID) < 2 THEN" & vbCrLf
'        SQLTXT = SQLTXT & "REPLICATE('0',3) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
'        SQLTXT = SQLTXT & "WHEN LEN(@ID) = 2 THEN" & vbCrLf
'        SQLTXT = SQLTXT & "REPLICATE('0',2) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
'        SQLTXT = SQLTXT & "END, @ID = @ID + 1" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE TRANNO = @RRNO_P  AND [TYPE] = 'P' AND ITEMNO = '1' AND TRANTYPE = 'RR'" & vbCrLf
'        SQLTXT = SQLTXT & "UPDATE PMIS_TDAYTRAN SET ITEMNO = '1' WHERE TRANNO = @RRNO_M AND [TYPE] = 'M' AND TRANTYPE = 'RR'" & vbCrLf
'        SQLTXT = SQLTXT & "SET @ID = 0" & vbCrLf
'        SQLTXT = SQLTXT & "UPDATE PMIS_TDAYTRAN SET ITEMNO =" & vbCrLf
'        SQLTXT = SQLTXT & "CASE WHEN LEN(@ID) < 2 THEN" & vbCrLf
'        SQLTXT = SQLTXT & "REPLICATE('0',3) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
'        SQLTXT = SQLTXT & "WHEN LEN(@ID) = 2 THEN" & vbCrLf
'        SQLTXT = SQLTXT & "REPLICATE('0',2) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
'        SQLTXT = SQLTXT & "END, @ID = @ID + 1" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE TRANNO = @RRNO_M  AND [TYPE] = 'M' AND ITEMNO = '1' AND TRANTYPE = 'RR'" & vbCrLf
'        SQLTXT = SQLTXT & "UPDATE PMIS_TDAYTRAN SET ITEMNO = '1' WHERE TRANNO = @RRNO_A AND [TYPE] = 'A' AND TRANTYPE = 'RR'" & vbCrLf
'        SQLTXT = SQLTXT & "SET @ID = 0" & vbCrLf
'        SQLTXT = SQLTXT & "UPDATE PMIS_TDAYTRAN SET ITEMNO =" & vbCrLf
'        SQLTXT = SQLTXT & "CASE WHEN LEN(@ID) < 2 THEN" & vbCrLf
'        SQLTXT = SQLTXT & "REPLICATE('0',3) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
'        SQLTXT = SQLTXT & "WHEN LEN(@ID) = 2 THEN" & vbCrLf
'        SQLTXT = SQLTXT & "REPLICATE('0',2) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
'        SQLTXT = SQLTXT & "END, @ID = @ID + 1" & vbCrLf
'        SQLTXT = SQLTXT & "WHERE TRANNO = @RRNO_A  AND [TYPE] = 'A' AND ITEMNO = '1' AND TRANTYPE = 'RR'" & vbCrLf


        Set rsTMP = gconDMIS.Execute(SQLTXT)
      
        If Not (rsTMP.EOF And rsTMP.BOF) Then
            Do While Not rsTMP.EOF
                
            
                TRAN_MAC = ComputeTransactionMac(rsTMP!STOCK_ORD, rsTMP!TRANQTY, rsTMP!TRANUCOST, rsTMP!trandate)
                M_MAC = ComputeMacasofDate(rsTMP!STOCK_ORD, rsTMP!trandate, rsTMP![Type])
                
                ITEM_NO = Format(ITEM_NO + 1, "0000")
                XSTOCK_ORD = rsTMP!STOCK_ORD
                XTYPE = rsTMP![Type]
                xTranDate = rsTMP!trandate
                xTRANTYPE = rsTMP!TranType
                XTRANNO = rsTMP!TRANNO
                XSTOCK_SUP = rsTMP!STOCK_SUP
                XTRANQTY = rsTMP!TRANQTY
                Xstatus = rsTMP!Status
                XIN_OUT = rsTMP!IN_OUT
                XUSERCODE = rsTMP!USERCODE
                XTREMARKS = rsTMP!TREMARKS
                XNON_HARI = rsTMP!NON_HARI
                XLASTUPDATE = rsTMP!LASTUPDATE
                XTRANUCOST = rsTMP!TRANUCOST
                
                If XTYPE = "P" Then
                     P_TTLRRAMT = P_TTLRRAMT + (rsTMP!TRANUCOST * rsTMP!TRANQTY)
                     P_TRANNO = XTRANNO
                ElseIf XTYPE = "M" Then
                     M_TTLRRAMT = M_TTLRRAMT + (rsTMP!TRANUCOST * rsTMP!TRANQTY)
                     M_TRANNO = XTRANNO
                Else
                     A_TTLRRAMT = A_TTLRRAMT + (rsTMP!TRANUCOST * rsTMP!TRANQTY)
                     A_TRANNO = XTRANNO
                End If
                
                Dim SQL As String
                
                SQL = "INSERT INTO PMIS_TDAYTRAN (STOCK_ORD,[TYPE],TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_SUP,TRANQTY,TRANUCOST,STATUS," & vbCrLf
                SQL = SQL & "IN_OUT,MAC,TRANINVAMT,USERCODE,LASTUPDATE,TREMARKS,NON_HARI)" & vbCrLf
                SQL = SQL & "VALUES('" & XSTOCK_ORD & "','" & XTYPE & "','" & xTranDate & "','" & xTRANTYPE & "','" & XTRANNO & "'," & vbCrLf
                SQL = SQL & "'" & ITEM_NO & "','" & XSTOCK_SUP & "','" & XTRANQTY & "','" & XTRANUCOST & "','" & Xstatus & "'," & vbCrLf
                SQL = SQL & "'" & XIN_OUT & "','" & TRAN_MAC & "','" & XTRANUCOST & "','" & XUSERCODE & "'," & vbCrLf
                SQL = SQL & "'" & XLASTUPDATE & "','" & XTREMARKS & "','" & XNON_HARI & "')" & vbCrLf
                
                Call gconDMIS.Execute(SQL)
                Call gconDMIS.Execute("UPDATE PMIS_STOCKMAS SET MAC = '" & TRAN_MAC & "',trecqty =(isnull(trecqty,0) + " & XTRANQTY & ")   WHERE STOCKNO = '" & XSTOCK_ORD & "' and [TYPE] = '" & XTYPE & "'")

            rsTMP.MoveNext
            Loop
            
                Call gconDMIS.Execute("UPDATE PMIS_RR_HD SET TTLRRAMT = '" & P_TTLRRAMT & "', NETRRAMT = '" & P_TTLRRAMT & "' WHERE RRNO = '" & P_TRANNO & "' AND [TYPE] = 'P'")
                Call gconDMIS.Execute("UPDATE PMIS_RR_HD SET TTLRRAMT = '" & M_TTLRRAMT & "', NETRRAMT = '" & M_TTLRRAMT & "' WHERE RRNO = '" & M_TRANNO & "' AND [TYPE] = 'M'")
                Call gconDMIS.Execute("UPDATE PMIS_RR_HD SET TTLRRAMT = '" & A_TTLRRAMT & "', NETRRAMT = '" & A_TTLRRAMT & "' WHERE RRNO = '" & A_TRANNO & "' AND [TYPE] = 'A'")

        End If
        
        
        SQL = ""
        Set rsTMP = Nothing
        CREATE_RR_DETAILS = True
        Exit Function
        
ErrorCode:
        MsgBox err.Description
        CREATE_RR_DETAILS = False
End Function
Function INSERT_SUPPLIER() As Boolean
        Dim SQLTXT As String
        
        On Error GoTo ErrorCode
        
        SQLTXT = "DECLARE @VENDORCODE AS NVARCHAR(10)"
        SQLTXT = SQLTXT & "DECLARE @VENDORNAME AS NVARCHAR(100)" & vbCrLf
        SQLTXT = SQLTXT & "SET @VENDORCODE =(SELECT DISTINCT(COMPANYCODE) + '001' FROM ALL_PROFILE) " & vbCrLf
        SQLTXT = SQLTXT & "SET @VENDORNAME = (SELECT DISTINCT(COMPANYCODE) + ' ' + 'HYUNDAI HUB SERVICE DEPARTMENT' FROM ALL_PROFILE)" & vbCrLf
        SQLTXT = SQLTXT & "INSERT INTO all_vendor_table (CODE,NAMEOFVENDOR,NONVAT)" & vbCrLf
        SQLTXT = SQLTXT & "VALUES(@VENDORCODE,@VENDORNAME,'Y')" & vbCrLf

        Call gconDMIS.Execute(SQLTXT)
        INSERT_SUPPLIER = True
        Exit Function
        
ErrorCode:
        INSERT_SUPPLIER = False
End Function
Function UPDATE_COLUMN_RECEIPTS_PARTSMASTERFILE() As Boolean
        Dim SQLTXT As String
    
        On Error GoTo ErrorCode
        
        SQLTXT = "UPDATE PMIS_STOCKMAS SET RECEIPTS = Y.ALL_RECEIPTS FROM PMIS_STOCKMAS X INNER JOIN" & vbCrLf
        SQLTXT = SQLTXT & "(SELECT STOCK_ORD, [TYPE],SUM(ISNULL(TRANQTY,0)) AS ALL_RECEIPTS FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],ISNULL(TRANQTY,0) AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE" & vbCrLf
        SQLTXT = SQLTXT & "TRANTYPE = 'RR' AND STATUS IN ('P','B')" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT STOCK_ORD,[TYPE],ISNULL(TRANQTY,0) AS TRANQTY FROM PMIS_ALLDAYTRAN WHERE" & vbCrLf
        SQLTXT = SQLTXT & "TRANTYPE = 'ADJ' AND STATUS IN ('P','B') AND IN_OUT = 'I'" & vbCrLf
        SQLTXT = SQLTXT & ") T GROUP BY STOCK_ORD ,[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & ") Y ON LTRIM(RTRIM(X.STOCKNO)) = LTRIM(RTRIM(Y.STOCK_ORD)) AND X.[TYPE] = Y.[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "Where IsNull(X.RECEIPTS, 0) <> IsNull(Y.ALL_RECEIPTS, 0)" & vbCrLf

        Call gconDMIS.Execute(SQLTXT)
        UPDATE_COLUMN_RECEIPTS_PARTSMASTERFILE = True
        Exit Function
        
ErrorCode:
        UPDATE_COLUMN_RECEIPTS_PARTSMASTERFILE = False
End Function
Function UPDATE_CSMS_RO_DET_LINE_SEQUENTIALLY(ro As String) As Boolean
        Dim SQLTXT As String

        On Error GoTo ErrorCode:
        
        
        SQLTXT = "DECLARE @ID AS INT" & vbCrLf
        SQLTXT = SQLTXT & "UPDATE CSMS_RO_DET SET LINE_NO = '1' WHERE REP_OR = '" & ro & "' AND  LIVIL ='2'" & vbCrLf
        SQLTXT = SQLTXT & "SET @ID = 0" & vbCrLf
        SQLTXT = SQLTXT & "UPDATE CSMS_RO_DET SET LINE_NO =" & vbCrLf
        SQLTXT = SQLTXT & "CASE WHEN LEN(@ID) < 2 THEN" & vbCrLf
        SQLTXT = SQLTXT & "REPLICATE('0',1) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(@ID) = 2  THEN" & vbCrLf
        SQLTXT = SQLTXT & "CAST(@ID AS NVARCHAR(20))" & vbCrLf
        SQLTXT = SQLTXT & "END ,@ID = @ID + 1" & vbCrLf
        SQLTXT = SQLTXT & "WHERE REP_OR = '" & ro & "' AND LIVIL='2' AND LINE_NO = '1'" & vbCrLf
        SQLTXT = SQLTXT & "UPDATE CSMS_RO_DET SET LINE_NO = '1' WHERE REP_OR = '" & ro & "' AND  LIVIL ='3'" & vbCrLf
        SQLTXT = SQLTXT & "SET @ID = 0" & vbCrLf
        SQLTXT = SQLTXT & "UPDATE CSMS_RO_DET SET LINE_NO =" & vbCrLf
        SQLTXT = SQLTXT & "CASE WHEN LEN(@ID) < 2 THEN" & vbCrLf
        SQLTXT = SQLTXT & "REPLICATE('0',1) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(@ID) = 2  THEN" & vbCrLf
        SQLTXT = SQLTXT & "CAST(@ID AS NVARCHAR(20))" & vbCrLf
        SQLTXT = SQLTXT & "END ,@ID = @ID + 1" & vbCrLf
        SQLTXT = SQLTXT & "WHERE REP_OR = '" & ro & "' AND LIVIL='3' AND LINE_NO = '1'" & vbCrLf
        SQLTXT = SQLTXT & "UPDATE CSMS_RO_DET SET LINE_NO = '1' WHERE REP_OR = '" & ro & "' AND  LIVIL ='4'" & vbCrLf
        SQLTXT = SQLTXT & "SET @ID = 0" & vbCrLf
        SQLTXT = SQLTXT & "UPDATE CSMS_RO_DET SET LINE_NO =" & vbCrLf
        SQLTXT = SQLTXT & "CASE WHEN LEN(@ID) < 2 THEN" & vbCrLf
        SQLTXT = SQLTXT & "REPLICATE('0',1) + CAST(@ID AS NVARCHAR(20))" & vbCrLf
        SQLTXT = SQLTXT & "WHEN LEN(@ID) = 2  THEN" & vbCrLf
        SQLTXT = SQLTXT & "CAST(@ID AS NVARCHAR(20))" & vbCrLf
        SQLTXT = SQLTXT & "END ,@ID = @ID + 1" & vbCrLf
        SQLTXT = SQLTXT & "WHERE REP_OR = '" & ro & "' AND LIVIL='4' AND LINE_NO = '1'" & vbCrLf

        Call gconDMIS.Execute(SQLTXT)
        UPDATE_CSMS_RO_DET_LINE_SEQUENTIALLY = True
        Exit Function
        
ErrorCode:
        UPDATE_CSMS_RO_DET_LINE_SEQUENTIALLY = False
End Function



Private Sub rcFind_SelectionChanged()
        On Error Resume Next
        Call rsRefresh
        RS_RETURN.Find "ID = " & rcFind.SelectedRows(0).Record(1).Value
        StoreMemVars
End Sub

Private Sub RcParts_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
     Dim ID                      As Long
         Dim Msg                     As String
         
         rsRefresh
         RS_RETURN.Find "ID = " & LABID_HD
         On Error GoTo ErrorCode:
         
         lblid = N2Str2IntZero(LTrim(RTrim(RcParts.SelectedRows(0).Record(7).Value)))
         ID = N2Str2IntZero(LTrim(RTrim(RcParts.SelectedRows(0).Record(7).Value)))
         lblid.Caption = ID
         
         lbltrantype.Caption = GETTRANTYPE(lblid)
         
         If Null2String(RS_RETURN!Status) = "P" Then
                Msg = "You cannot edit transaction its already Posted" & vbCrLf
                Msg = Msg & "Please Unpost the transaction to Edit!"
                
                MsgBox Msg, vbInformation
                Exit Sub
         End If
         
          If CheckIfthisMonthIssue(txtRep_or, lblid) = True Then
            Msg = "You cannot return this Partnumber" & vbCrLf
            Msg = Msg & "Its still can be Unpost in Parts Issuance!"
          
             MsgBox Msg, vbInformation
             Msg = ""
            Exit Sub
         End If
         
                If If_Item_Exits_inCsmS_det(LABID_HD, lblid) = True Then
                    ADDOREDIT_DET = "EDIT"
                Else
                    ADDOREDIT_DET = "ADD"
                End If
               
                cbo_Tran_Partnumber = RcParts.SelectedRows(0).Record(0).Value
                txtTran_return = RcParts.SelectedRows(0).Record(4).Value
                Call ReturnedParts(Null2String(ID), txtRep_or)
                pic_Returnpart.ZOrder 0
                pic_Returnpart.Visible = True
                pic_Select.ZOrder 1
                pic_Select.Visible = False
                picAdd.Enabled = False
                rcFind.Enabled = False
                txtsearch.Enabled = False
        
ErrorCode:
    Exit Sub
    
End Sub
Function GETTRANTYPE(XXX As Long) As String
    Dim SQLTXT As String
    Dim rsTMP As New ADODB.Recordset
    
    SQLTXT = "SELECT TRANTYPE FROM PMIS_ALLDAYTRAN WHERE ID = '" & XXX & "'"
    Set rsTMP = gconDMIS.Execute(SQLTXT)
    
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        GETTRANTYPE = Null2String(rsTMP!TranType)
    End If

    Set rsTMP = Nothing
End Function


Private Sub RcReq_parts_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
     Dim Msg As String
        rsRefresh
        RS_RETURN.Find "ID = " & LABID_HD
        
        If Null2String(RS_RETURN!Status) = "P" Then
               Msg = "You cannot delete the item" & vbCrLf
               Msg = Msg & "Transaction is already Posted!"
               MsgBox Msg, vbInformation
               Exit Sub
        End If
    
               lblid = RcReq_parts.SelectedRows(0).Record(7).Value
               cmdDelete_det.Value = True
               RcReq_parts.Populate
End Sub


Private Sub txtFindSupp_Change()
      What_Func = True
        Call DisplaySupp(What_Func, txtFindSupp)
End Sub

Private Sub DisplaySupp(X As Boolean, Optional str_data As String)
    Dim Item As ListItem
    Dim rsTMP As New ADODB.Recordset
    Dim SQLTXT As String
    
    If X = True Then
        SQLTXT = "SELECT Code,NameofVendor FROM ALL_VENDOR WHERE NameofVendor LIKE '" & Null2String(str_data) & "%' "
    Else
        SQLTXT = "SELECT Code,NameofVendor FROM ALL_VENDOR"
    End If
    Set rsTMP = gconDMIS.Execute(SQLTXT)
    
    lvwSupp.ListItems.Clear
    
    If Not (rsTMP.EOF And rsTMP.BOF) Then
        Do While Not rsTMP.EOF
           Set Item = lvwSupp.ListItems.Add(, , rsTMP!Code)
           Item.SubItems(1) = rsTMP!nameofvendor
        
        rsTMP.MoveNext
        Loop
    
    End If

    Set rsTMP = Nothing
End Sub

'Private Sub Timer1_Timer()
'    If OnUpdate = True Then
'    On Error Resume Next
'        If (RS_RETURN!Status) = "P" Then
'              LblVerify.Caption = "POSTED"
'        ElseIf (RS_RETURN!Status) = "P" And (RS_RETURN!veri_by) Is Not Null Then
'                LblVerify.Caption = "VERIFIED"
'        ElseIf (RS_RETURN!Status) = "C" Then
'                LblVerify.Caption = "CANCELLED"
'        ElseIf (RS_RETURN!Status) = "N" Then
'                 LblVerify.Caption = "NOT YET VERIFY"
'        End If
'
'    Else
'        LblVerify.Visible = True
'    End If
'
'End Sub
Private Sub txtRep_or_KeyPress(KeyAscii As Integer)
   
      If KeyAscii = 13 Then
        
            Dim RONOStr  As String
            
            RONOStr = txtRep_or.Text
            If VALID_COMPANY_CODE_FORHAI = True Then
                If UCase(Left(RONOStr, 2)) = "R-" Then
                    RONOStr = "J-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
                ElseIf UCase(Left(RONOStr, 2)) = "J-" Then
                    RONOStr = "J-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
                Else
                    RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
                    If CheckifROhasIssuedParts(RONOStr) = False Then RONOStr = "J-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
                End If
            Else
                If UCase(Left(RONOStr, 2)) = "R-" Then
                    RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
                ElseIf UCase(Left(RONOStr, 2)) = "J-" Then
                    RONOStr = "J-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
                Else
                    RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
                End If
            End If
            txtRep_or.Text = RONOStr
            
            If CheckIfRoIsAlreadyInvoice(txtRep_or.Text) = True Then
                MessagePop InfoStop, "Invalid RO number", "Repair Order is Already Billed out"
                txtRep_or.Text = ""
                Exit Sub
            End If
             
            If CheckIfROStillExist(txtRep_or.Text) = False Then
                MessagePop InfoStop, "Invalid RO number", "Repair Order Cannot Find in Database!"
                txtRep_or.Text = ""
                Exit Sub
            End If
            
            If CheckifROhasIssuedParts(txtRep_or) = False Then
                MessagePop InfoStop, "Invalid RO number", "No Parts issued in this repair order!"
                txtRep_or.Text = ""
                txtRep_or.SetFocus
                Exit Sub
            End If
        
'            If CheckIfROalreadyVerified(txtRep_or.Text) = True Then
'                MessagePop InfoStop, "Invalid RO number", "Repair Order is in Use!"
'                txtRep_or.Text = ""
'                Exit Sub
'            End If
            LOAD_DATA = True
            txtRemarks.Enabled = True
            txt_req_by.Enabled = True
            Call show_allparts(txtRep_or.Text)
                 'rcParts.Enabled = False
        End If
End Sub
Function CheckifROhasIssuedParts(ro As String) As Boolean
        Dim SQLTXT As String
        Dim rsTMP As New ADODB.Recordset
        
        SQLTXT = "SELECT * FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT A.RONO,A.STATUS,A.TRANNO,B.TRANTYPE FROM PMIS_ORD_HD A INNER JOIN PMIS_TDAYTRAN B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & "Union All" & vbCrLf
        SQLTXT = SQLTXT & "SELECT A.RONO,A.STATUS,A.TRANNO,B.TRANTYPE FROM PMIS_ORD_HIST A INNER JOIN PMIS_DAYTRAN B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.TRANNO = B.TRANNO AND A.[TYPE] = B.[TYPE] AND A.TRANTYPE = B.TRANTYPE" & vbCrLf
        SQLTXT = SQLTXT & ") T WHERE RONO ='" & ro & "' AND STATUS in('P','B')  AND TRANTYPE IN('RIV','ADB')" & vbCrLf
        
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        If Not (rsTMP.EOF And rsTMP.BOF) Then
             CheckifROhasIssuedParts = True
        Else
            CheckifROhasIssuedParts = False
        End If
        
        Set rsTMP = Nothing
End Function
Function CheckIfROalreadyVerified(ro As String) As Boolean
        Dim SQLTXT As String
        Dim rsTMP As New ADODB.Recordset
        
        SQLTXT = "SELECT * FROM CSMS_RETURN_HD WHERE REP_OR = '" & ro & "'"
        Set rsTMP = gconDMIS.Execute(SQLTXT)
        
        If Not (rsTMP.EOF And rsTMP.BOF) Then
            CheckIfROalreadyVerified = True
        Else
            CheckIfROalreadyVerified = False
        End If
        
        Set rsTMP = Nothing
End Function

Private Sub txtSEARCH_Change()

         rcFind.FilterText = txtsearch.Text
         rcFind.Populate
'        Dim SQLTXT  As String
'        Dim RSTMP   As New ADODB.Recordset
'        Dim REC     As XtremeReportControl.ReportRecord
'

'        SQLTXT = SearchRepairOrder(LTrim(RTrim(txtSearch.Text)))
'        Set RSTMP = gconDMIS.Execute(SQLTXT)
'        rcFind.ListItems.Clear
'
'        rcFind.Records.DeleteAll
'
'        If Not (RSTMP.EOF And RSTMP.BOF) Then
'            Do While Not RSTMP.EOF
'                Set REC = rcFind.Records.Add
'                REC.AddItem Null2String(RSTMP!REP_OR)
'                REC.AddItem N2Str2IntZero(RSTMP!ID)
'
'            RSTMP.MoveNext
'            Loop
'        rcFind.Populate
'        End If
        
'        Set RSTMP = Nothing
End Sub

Function SearchRepairOrder(XXX As String) As String
        Dim SQLTXT As String
        
        SQLTXT = "SELECT * FROM CSMS_RETURN_HD WHERE REP_OR LIKE '%" & XXX & "%'"
        SearchRepairOrder = SQLTXT
End Function




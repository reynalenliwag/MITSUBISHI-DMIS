VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAISApplications 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application Form"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmApplications.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   12270
   Begin XtremeSuiteControls.TabControl tbcApplication 
      Height          =   7785
      Left            =   2550
      TabIndex        =   58
      Top             =   0
      Width           =   9615
      _Version        =   655364
      _ExtentX        =   16960
      _ExtentY        =   13732
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.Layout=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      ItemCount       =   7
      Item(0).Caption =   "Personal Details"
      Item(0).Tooltip =   "Personal Details"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "picTAB_PER"
      Item(1).Caption =   "Family BackGround"
      Item(1).Tooltip =   "Family Background"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "picTAB_FAMILY"
      Item(2).Caption =   "Educational Attainment"
      Item(2).Tooltip =   "Educational Attainment"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "picTAB_EDU"
      Item(3).Caption =   "Trainings and Seminars Attended"
      Item(3).Tooltip =   "Trainings and Seminars Attended"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "picTAB_TRAIN"
      Item(4).Caption =   "Employment Record"
      Item(4).Tooltip =   "Employment Record"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "picTAB_EMP"
      Item(5).Caption =   "References"
      Item(5).Tooltip =   "References"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "PicTAB_REF"
      Item(6).Caption =   "Document Pass"
      Item(6).Tooltip =   "Document Pass"
      Item(6).ControlCount=   1
      Item(6).Control(0)=   "picTAB_DOC"
      Begin VB.PictureBox picTAB_EMP 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7395
         Left            =   -69970
         ScaleHeight     =   7395
         ScaleWidth      =   9555
         TabIndex        =   121
         Top             =   360
         Visible         =   0   'False
         Width           =   9555
         Begin VB.PictureBox picEMP 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   7065
            ScaleHeight     =   810
            ScaleWidth      =   2280
            TabIndex        =   122
            Top             =   2835
            Visible         =   0   'False
            Width           =   2280
            Begin VB.CommandButton cmdEMP_ADD 
               Caption         =   "&Add"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   1440
               Picture         =   "frmApplications.frx":058A
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Add Employment Record"
               Top             =   0
               Width           =   795
            End
            Begin VB.CommandButton Command1 
               Caption         =   "EDIT"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   945
               Left            =   210
               TabIndex        =   123
               Top             =   1110
               Visible         =   0   'False
               Width           =   945
            End
         End
         Begin MSComctlLib.ListView lsvEMP 
            Height          =   2600
            Left            =   105
            TabIndex        =   51
            Top             =   150
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   4577
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name Of Company"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Address"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Position"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "From - To"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   9
            EndProperty
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click To Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   45
            Left            =   150
            TabIndex        =   126
            Top             =   3120
            Width           =   1890
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Maximum of 5"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   124
            Top             =   2880
            Width           =   1380
         End
      End
      Begin VB.PictureBox PicTAB_REF 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7395
         Left            =   -69970
         ScaleHeight     =   7395
         ScaleWidth      =   9555
         TabIndex        =   99
         Top             =   360
         Visible         =   0   'False
         Width           =   9555
         Begin VB.PictureBox PicREF 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   7065
            ScaleHeight     =   840
            ScaleWidth      =   2280
            TabIndex        =   100
            Top             =   2835
            Visible         =   0   'False
            Width           =   2280
            Begin VB.CommandButton cmdREF_ADD 
               Caption         =   "&Add"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   1440
               Picture         =   "frmApplications.frx":0B15
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Add Reference"
               Top             =   0
               Width           =   795
            End
         End
         Begin MSComctlLib.ListView LsvREF 
            Height          =   2600
            Left            =   105
            TabIndex        =   53
            Top             =   150
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   4577
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Address"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Position"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Tel. no"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   18
            EndProperty
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click To Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   48
            Left            =   150
            TabIndex        =   129
            Top             =   3120
            Width           =   1890
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Maximum of 5"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   43
            Left            =   150
            TabIndex        =   101
            Top             =   2880
            Width           =   1380
         End
      End
      Begin VB.PictureBox picTAB_DOC 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7395
         Left            =   -69970
         ScaleHeight     =   7395
         ScaleWidth      =   9555
         TabIndex        =   94
         Top             =   360
         Visible         =   0   'False
         Width           =   9555
         Begin VB.PictureBox picDOC 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   8385
            ScaleHeight     =   825
            ScaleWidth      =   1035
            TabIndex        =   95
            Top             =   2835
            Visible         =   0   'False
            Width           =   1035
            Begin VB.CommandButton cmdDOC_ADD 
               Caption         =   "&Add"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   120
               Picture         =   "frmApplications.frx":10A0
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Add Document Pass"
               Top             =   0
               Width           =   795
            End
         End
         Begin MSComctlLib.ListView lsvDOC 
            Height          =   2600
            Left            =   100
            TabIndex        =   55
            Top             =   150
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   4577
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Type of Documents"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   18
            EndProperty
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click To Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   49
            Left            =   150
            TabIndex        =   130
            Top             =   2880
            Width           =   1890
         End
      End
      Begin VB.PictureBox picTAB_TRAIN 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7395
         Left            =   -69970
         ScaleHeight     =   7395
         ScaleWidth      =   9555
         TabIndex        =   83
         Top             =   360
         Visible         =   0   'False
         Width           =   9555
         Begin VB.PictureBox picTRAIN 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   7185
            ScaleHeight     =   990
            ScaleWidth      =   2235
            TabIndex        =   93
            Top             =   2835
            Visible         =   0   'False
            Width           =   2235
            Begin VB.CommandButton cmdTRAIN_ADD 
               Caption         =   "&Add"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   1410
               Picture         =   "frmApplications.frx":162B
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Add Trainings and Seminars"
               Top             =   0
               Width           =   795
            End
         End
         Begin MSComctlLib.ListView lsvTRAIN 
            Height          =   2600
            Left            =   105
            TabIndex        =   49
            Top             =   150
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   4577
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Training"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "month - Year"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Place"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Sponsor"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   18
            EndProperty
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click To Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   47
            Left            =   150
            TabIndex        =   128
            Top             =   3120
            Width           =   1890
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Maximum of 5"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   21
            Left            =   150
            TabIndex        =   96
            Top             =   2880
            Width           =   1380
         End
      End
      Begin VB.PictureBox picTAB_FAMILY 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7395
         Left            =   -69970
         ScaleHeight     =   7395
         ScaleWidth      =   9555
         TabIndex        =   82
         Top             =   360
         Visible         =   0   'False
         Width           =   9555
         Begin VB.TextBox txtFAMILY_SAGE 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   3
            TabIndex        =   25
            ToolTipText     =   "Spouse Age"
            Top             =   690
            Width           =   795
         End
         Begin VB.TextBox txtFAMILY_FAGE 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   3
            TabIndex        =   28
            ToolTipText     =   "Father's Age"
            Top             =   2190
            Width           =   825
         End
         Begin VB.TextBox txtFAMILY_MAGE 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   3
            TabIndex        =   31
            ToolTipText     =   "Mother's Age"
            Top             =   3600
            Width           =   825
         End
         Begin VB.TextBox txtFAMILY_SNAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   30
            TabIndex        =   24
            ToolTipText     =   "Spouse Full Name"
            Top             =   240
            Width           =   4065
         End
         Begin VB.TextBox txtFAMILY_SOCCU 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   30
            TabIndex        =   26
            ToolTipText     =   "Spouse Occupation"
            Top             =   1170
            Width           =   6525
         End
         Begin VB.TextBox txtFAMILY_FNAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   30
            TabIndex        =   27
            ToolTipText     =   "Father's Full Name"
            Top             =   1710
            Width           =   4005
         End
         Begin VB.TextBox txtFAMILY_FOCCU 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   30
            TabIndex        =   29
            ToolTipText     =   "Father's Occupation"
            Top             =   2640
            Width           =   6525
         End
         Begin VB.TextBox txtFAMILY_MNAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   30
            TabIndex        =   30
            ToolTipText     =   "Mother's Full Name"
            Top             =   3120
            Width           =   4005
         End
         Begin VB.TextBox txtFAMILY_MOCCU 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            MaxLength       =   30
            TabIndex        =   32
            ToolTipText     =   "Mother's Occupation"
            Top             =   4080
            Width           =   6525
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name Of Mother"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   15
            Left            =   270
            TabIndex        =   92
            Top             =   3210
            Width           =   1605
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   16
            Left            =   1410
            TabIndex        =   91
            Top             =   3690
            Width           =   375
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   17
            Left            =   750
            TabIndex        =   90
            Top             =   4110
            Width           =   1125
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name Of Father"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   18
            Left            =   300
            TabIndex        =   89
            Top             =   1830
            Width           =   1560
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   19
            Left            =   1500
            TabIndex        =   88
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   34
            Left            =   720
            TabIndex        =   87
            Top             =   2730
            Width           =   1125
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name of Spouse"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   35
            Left            =   270
            TabIndex        =   86
            Top             =   330
            Width           =   1605
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   36
            Left            =   1470
            TabIndex        =   85
            Top             =   780
            Width           =   375
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   37
            Left            =   720
            TabIndex        =   84
            Top             =   1260
            Width           =   1125
         End
      End
      Begin VB.PictureBox picTAB_PER 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7395
         Left            =   30
         ScaleHeight     =   7395
         ScaleWidth      =   9555
         TabIndex        =   60
         Top             =   360
         Width           =   9555
         Begin VB.ComboBox cboPER_CITY 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1680
            TabIndex        =   5
            Text            =   "Combo1"
            Top             =   3060
            Width           =   5085
         End
         Begin VB.TextBox txtPOSITION 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   0
            ToolTipText     =   "Position Desired"
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox txtPER_EMAIL 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   7
            ToolTipText     =   "E-Mail Address"
            Top             =   3930
            Width           =   4605
         End
         Begin VB.TextBox txtPER_CNO 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   6
            ToolTipText     =   "Contact Number"
            Top             =   3480
            Width           =   3045
         End
         Begin VB.PictureBox picAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1725
            Left            =   7110
            ScaleHeight     =   1695
            ScaleWidth      =   2085
            TabIndex        =   61
            Top             =   150
            Width           =   2115
            Begin VB.Image imgAPP 
               Height          =   1575
               Left            =   60
               Stretch         =   -1  'True
               Top             =   60
               Width           =   1965
            End
         End
         Begin VB.TextBox txtPER_LNAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   1
            ToolTipText     =   "Last Name"
            Top             =   930
            Width           =   4935
         End
         Begin VB.TextBox txtPER_FNAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   2
            ToolTipText     =   "First Name"
            Top             =   1350
            Width           =   4935
         End
         Begin VB.TextBox txtPER_MNAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   3
            ToolTipText     =   "Middle Initial"
            Top             =   1800
            Width           =   4935
         End
         Begin VB.TextBox txtPER_ADD 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            ToolTipText     =   "Applicant Address"
            Top             =   2220
            Width           =   5025
         End
         Begin VB.TextBox txtPER_BPlace 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   9
            ToolTipText     =   "BirthPlace"
            Top             =   4380
            Width           =   4485
         End
         Begin VB.TextBox txtPER_Height 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   12
            ToolTipText     =   "Height"
            Top             =   5280
            Width           =   2445
         End
         Begin VB.TextBox txtPER_Weight 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5370
            MaxLength       =   10
            TabIndex        =   13
            ToolTipText     =   "Weight"
            Top             =   5280
            Width           =   2445
         End
         Begin VB.TextBox txtPER_Religion 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   14
            ToolTipText     =   "Religion"
            Top             =   5730
            Width           =   6165
         End
         Begin VB.TextBox txtPER_Citizenship 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   15
            ToolTipText     =   "Citizenship"
            Top             =   6180
            Width           =   4065
         End
         Begin VB.ComboBox cboPER_GENDER 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmApplications.frx":1BB6
            Left            =   1680
            List            =   "frmApplications.frx":1BB8
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Gender"
            Top             =   4830
            Width           =   2205
         End
         Begin VB.ComboBox cboPER_CSTATUS 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5370
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Civil Status"
            Top             =   4830
            Width           =   2685
         End
         Begin VB.CommandButton cmdPIC_CHANGE 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7140
            TabIndex        =   16
            Top             =   2040
            Width           =   945
         End
         Begin VB.CommandButton cmdPIC_DELETE 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8250
            TabIndex        =   17
            Top             =   2040
            Width           =   945
         End
         Begin MSComCtl2.DTPicker dtpPER_BirthDate 
            Height          =   345
            Left            =   1680
            TabIndex        =   8
            ToolTipText     =   "Birthdate"
            Top             =   4410
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            _Version        =   393216
            Format          =   54657025
            CurrentDate     =   39125
         End
         Begin MSComDlg.CommonDialog CDPIC 
            Left            =   8520
            Top             =   2910
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "City/ Municipality"
            Height          =   480
            Index           =   50
            Left            =   0
            TabIndex        =   132
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label lblREQ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   4
            Left            =   6750
            TabIndex        =   131
            Top             =   570
            Width           =   135
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position Desired"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   42
            Left            =   60
            TabIndex        =   98
            Top             =   600
            Width           =   1560
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail Add."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   390
            TabIndex        =   97
            Top             =   3990
            Width           =   1125
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact no."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   360
            TabIndex        =   81
            Top             =   3540
            Width           =   1185
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   540
            TabIndex        =   80
            Top             =   4470
            Width           =   990
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   780
            TabIndex        =   79
            Top             =   4890
            Width           =   690
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   480
            TabIndex        =   78
            Top             =   2220
            Width           =   1065
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   420
            TabIndex        =   77
            Top             =   1020
            Width           =   1020
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   390
            TabIndex        =   76
            Top             =   1380
            Width           =   1050
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   300
            TabIndex        =   75
            Top             =   1890
            Width           =   1230
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   900
            TabIndex        =   74
            Top             =   5370
            Width           =   630
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   10
            Left            =   4470
            TabIndex        =   73
            Top             =   5370
            Width           =   690
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   11
            Left            =   750
            TabIndex        =   72
            Top             =   5820
            Width           =   735
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Civil Status"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   12
            Left            =   4110
            TabIndex        =   71
            Top             =   4920
            Width           =   1125
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Place"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   13
            Left            =   3600
            TabIndex        =   70
            Top             =   4440
            Width           =   1050
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Citizenship"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   14
            Left            =   450
            TabIndex        =   69
            Top             =   6240
            Width           =   1050
         End
         Begin VB.Label lblREQ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   0
            Left            =   6750
            TabIndex        =   68
            Top             =   990
            Width           =   135
         End
         Begin VB.Label lblREQ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   1
            Left            =   6750
            TabIndex        =   67
            Top             =   1410
            Width           =   135
         End
         Begin VB.Label lblREQ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   2
            Left            =   4980
            TabIndex        =   66
            Top             =   3540
            Width           =   135
         End
         Begin VB.Label lblREQ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* - Required Fields"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   3
            Left            =   1890
            TabIndex        =   65
            Top             =   6690
            Width           =   1605
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Applicant no."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   20
            Left            =   150
            TabIndex        =   64
            Top             =   150
            Width           =   1305
         End
         Begin VB.Label lblAPPNO 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1680
            TabIndex        =   63
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lblPIC_LOC 
            BackColor       =   &H000000FF&
            Height          =   345
            Left            =   7140
            TabIndex        =   62
            Top             =   2430
            Visible         =   0   'False
            Width           =   2115
         End
      End
      Begin VB.Timer tmr_APP 
         Interval        =   500
         Left            =   10260
         Top             =   3540
      End
      Begin VB.PictureBox picTAB_EDU 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7395
         Left            =   -69970
         ScaleHeight     =   7395
         ScaleWidth      =   9555
         TabIndex        =   102
         Top             =   360
         Visible         =   0   'False
         Width           =   9555
         Begin VB.Frame fme1st 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Highest Academic Qualification"
            Height          =   3525
            Left            =   60
            TabIndex        =   112
            Top             =   60
            Width           =   9465
            Begin VB.ComboBox cbo1st_Level 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   33
               ToolTipText     =   "Highest Level Educational Attainment"
               Top             =   300
               Width           =   4875
            End
            Begin VB.ComboBox cbo1st_Field 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   34
               ToolTipText     =   "Field Of Study"
               Top             =   750
               Width           =   4875
            End
            Begin VB.ComboBox cbo1st_Month 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   39
               ToolTipText     =   "Graduation Month"
               Top             =   3030
               Width           =   1965
            End
            Begin VB.TextBox txt1st_Major 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2100
               TabIndex        =   35
               ToolTipText     =   "Major "
               Top             =   1170
               Width           =   5925
            End
            Begin VB.TextBox txt1st_Grade 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2100
               MaxLength       =   3
               TabIndex        =   36
               ToolTipText     =   "Grade"
               Top             =   1620
               Width           =   1995
            End
            Begin VB.TextBox txt1st_Ins 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2100
               TabIndex        =   37
               ToolTipText     =   "School Name"
               Top             =   2100
               Width           =   5895
            End
            Begin VB.TextBox txt1st_Add 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2100
               TabIndex        =   38
               ToolTipText     =   "School  Address"
               Top             =   2550
               Width           =   5895
            End
            Begin VB.TextBox txt1st_Year 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   4260
               MaxLength       =   4
               TabIndex        =   40
               ToolTipText     =   "Graduation Year"
               Top             =   3000
               Width           =   1695
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grade"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   22
               Left            =   1290
               TabIndex        =   120
               Top             =   1680
               Width           =   570
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Highest Level"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   23
               Left            =   570
               TabIndex        =   119
               Top             =   390
               Width           =   1320
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Field Of Study"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   24
               Left            =   480
               TabIndex        =   118
               Top             =   840
               Width           =   1410
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Major"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   25
               Left            =   1350
               TabIndex        =   117
               Top             =   1260
               Width           =   540
            End
            Begin VB.Label lblCAP 
               BackStyle       =   0  'Transparent
               Caption         =   "Institute /  University"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Index           =   26
               Left            =   870
               TabIndex        =   116
               Top             =   2040
               Width           =   1065
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   27
               Left            =   1110
               TabIndex        =   115
               Top             =   2670
               Width           =   780
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Graduation Date"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   28
               Left            =   330
               TabIndex        =   114
               Top             =   3090
               Width           =   1605
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "(yyyy)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   29
               Left            =   6120
               TabIndex        =   113
               Top             =   3090
               Width           =   660
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2nd Highest Academic Qualification"
            Height          =   3675
            Left            =   60
            TabIndex        =   103
            Top             =   3660
            Width           =   9465
            Begin VB.ComboBox cbo2nd_Level 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2070
               Style           =   2  'Dropdown List
               TabIndex        =   41
               ToolTipText     =   "2nd Highest Level Educational Attainment"
               Top             =   390
               Width           =   4875
            End
            Begin VB.ComboBox cbo2nd_Field 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2070
               Style           =   2  'Dropdown List
               TabIndex        =   42
               ToolTipText     =   "Field Of Study"
               Top             =   840
               Width           =   4875
            End
            Begin VB.ComboBox cbo2nd_Month 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2070
               Style           =   2  'Dropdown List
               TabIndex        =   47
               ToolTipText     =   "Graduation Month"
               Top             =   3120
               Width           =   1965
            End
            Begin VB.TextBox txt2nd_Major 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2070
               TabIndex        =   43
               ToolTipText     =   "Major "
               Top             =   1260
               Width           =   5925
            End
            Begin VB.TextBox txt2nd_grade 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2070
               MaxLength       =   3
               TabIndex        =   44
               ToolTipText     =   "Grade"
               Top             =   1710
               Width           =   1995
            End
            Begin VB.TextBox txt2nd_Ins 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2070
               TabIndex        =   45
               ToolTipText     =   "School Name"
               Top             =   2190
               Width           =   5895
            End
            Begin VB.TextBox txt2nd_Add 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2070
               TabIndex        =   46
               ToolTipText     =   "School  Address"
               Top             =   2640
               Width           =   5895
            End
            Begin VB.TextBox txt2nd_Year 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   4230
               MaxLength       =   4
               TabIndex        =   48
               ToolTipText     =   "Graduation Year"
               Top             =   3090
               Width           =   1695
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grade"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   30
               Left            =   1260
               TabIndex        =   111
               Top             =   1770
               Width           =   570
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2nd Highest Level"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   31
               Left            =   150
               TabIndex        =   110
               Top             =   480
               Width           =   1755
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Field Of Study"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   32
               Left            =   450
               TabIndex        =   109
               Top             =   930
               Width           =   1410
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Major"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   33
               Left            =   1320
               TabIndex        =   108
               Top             =   1350
               Width           =   540
            End
            Begin VB.Label lblCAP 
               BackStyle       =   0  'Transparent
               Caption         =   "Institute /  University"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Index           =   38
               Left            =   840
               TabIndex        =   107
               Top             =   2130
               Width           =   1065
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   39
               Left            =   1080
               TabIndex        =   106
               Top             =   2760
               Width           =   780
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Graduation Date"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   40
               Left            =   300
               TabIndex        =   105
               Top             =   3180
               Width           =   1605
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "(yyyy)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   41
               Left            =   6090
               TabIndex        =   104
               Top             =   3180
               Width           =   660
            End
         End
      End
   End
   Begin Crystal.CrystalReport rptAPP 
      Left            =   3690
      Top             =   7980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picSEARCH 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8700
      Left            =   0
      ScaleHeight     =   8700
      ScaleWidth      =   2475
      TabIndex        =   57
      Top             =   0
      Width           =   2475
      Begin VB.CommandButton cmdUPLOAD 
         Caption         =   "Upload Applicant"
         Height          =   330
         Left            =   90
         TabIndex        =   23
         ToolTipText     =   "Upload Applicant"
         Top             =   8250
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdEXAM 
         Caption         =   "View Exam Taken"
         Height          =   330
         Left            =   90
         TabIndex        =   22
         ToolTipText     =   "View Exam Taken"
         Top             =   7860
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picACTIVE 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   30
         ScaleHeight     =   855
         ScaleWidth      =   2685
         TabIndex        =   59
         Top             =   90
         Width           =   2685
         Begin VB.OptionButton optACTIVE_EMP 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hired Applicant"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   90
            TabIndex        =   18
            Top             =   90
            Width           =   2415
         End
         Begin VB.OptionButton optINACTIVE_EMP 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Applicant"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   90
            TabIndex        =   19
            Top             =   420
            Width           =   2415
         End
      End
      Begin VB.TextBox txtSEARCH 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   30
         TabIndex        =   20
         Top             =   1380
         Width           =   2445
      End
      Begin MSComctlLib.ListView lsvSEARCH 
         Height          =   3255
         Left            =   30
         TabIndex        =   21
         Top             =   1830
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Full Name"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - SEARCH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   134
         Top             =   5220
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F2 - SAVE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   133
         Top             =   5550
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin VB.PictureBox picAdds 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2700
      ScaleHeight     =   960
      ScaleWidth      =   9495
      TabIndex        =   138
      Top             =   7860
      Width           =   9495
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
         Left            =   8730
         MouseIcon       =   "frmApplications.frx":1BBA
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":1D0C
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Exit Window"
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
         Height          =   795
         Left            =   8010
         MouseIcon       =   "frmApplications.frx":2072
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":21C4
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
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
         Left            =   7290
         MouseIcon       =   "frmApplications.frx":252A
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":267C
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
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
         Left            =   6570
         MouseIcon       =   "frmApplications.frx":29A7
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":2AF9
         Style           =   1  'Graphical
         TabIndex        =   144
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
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
         Left            =   5850
         MouseIcon       =   "frmApplications.frx":2E55
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":2FA7
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
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
         Left            =   5130
         MouseIcon       =   "frmApplications.frx":32BA
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":340C
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Move to Last Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
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
         Left            =   4410
         MouseIcon       =   "frmApplications.frx":375C
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":38AE
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   735
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
         Left            =   3690
         MouseIcon       =   "frmApplications.frx":3C0C
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":3D5E
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
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
         Left            =   2970
         MouseIcon       =   "frmApplications.frx":4058
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":41AA
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
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
         Left            =   2250
         MouseIcon       =   "frmApplications.frx":4502
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":4654
         Style           =   1  'Graphical
         TabIndex        =   148
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picSaves 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   10620
      ScaleHeight     =   885
      ScaleWidth      =   1590
      TabIndex        =   135
      Top             =   7860
      Width           =   1590
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
         Left            =   810
         MouseIcon       =   "frmApplications.frx":49B3
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":4B05
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Cancel"
         Top             =   30
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
         Height          =   795
         Left            =   90
         MouseIcon       =   "frmApplications.frx":4E43
         MousePointer    =   99  'Custom
         Picture         =   "frmApplications.frx":4F95
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Save Applicant Information"
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Double Click To Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   46
      Left            =   0
      TabIndex        =   127
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Maximum of 5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   44
      Left            =   0
      TabIndex        =   125
      Top             =   0
      Width           =   1380
   End
End
Attribute VB_Name = "frmAISApplications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FUNCTION /FEATURE:Add new Control Button For Prev/Next/First/Last
'DATE STARTED:06/08/2007
'LAST UPDATE:
'DATABASE UPDATE:
'WHO UPDATE:HardNard
'UPDATING CODE:BTT - 06/07/2007
'**********************************************************************************
Option Explicit
Dim RS                                                                As New ADODB.Recordset    'BTT - 07032007

Function GenerateApplicationNumber()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select Applicant_ID From HRMS_Applicant_Personal Order By Applicant_ID ASC")

    APPLICANT_ID = 0
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            APPLICANT_ID = RSTMP!APPLICANT_ID
            RSTMP.MoveNext
        Loop
    End If

    APPLICANT_ID = APPLICANT_ID + 1
    lblAPPNO.Caption = APPLICANT_ID

    Set RSTMP = Nothing
End Function

Function CheckIfSearchExist(TABLE As String, ID As Integer, EXIST As Boolean)
    Dim SQL                                                           As String
    Dim RSTMP                                                         As ADODB.Recordset

    SQL = "Select * From " & TABLE & " Where Entry_ID = " & ID & " And  Applicant_ID = " & APPLICANT_ID & ""
    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute(SQL)

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        EXIST = True
    Else
        EXIST = False
    End If

    Set RSTMP = Nothing
End Function

Function DisplayTraintobeEDITED()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_APPLICANT_Train Where Applicant_ID = " & _
                                 APPLICANT_ID & " And Entry_ID = " & TRAINING_ENTRY_ID & "")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        frmAISADD_TRAIN.txtTRAIN_TRAIN.Text = Null2String(RSTMP!Training)
        frmAISADD_TRAIN.txtTRAIN_MonthYear.Text = Null2String(RSTMP!monthYear)
        frmAISADD_TRAIN.txtTRAIN_PLACE.Text = Null2String(RSTMP!Place)
        frmAISADD_TRAIN.txtTRAIN_SPONSOR.Text = Null2String(RSTMP!Sponsor)
    End If
End Function

Function DisplayRequiredFields(COND As Boolean)
    Dim X                                                             As Integer

    For X = 0 To 4
        lblREQ(X).Visible = COND
    Next
End Function

Function CleanAllUnsaveRecords()
    gconDMIS.Execute ("Delete From HRMS_APPLICANT_TRAIN Where Applicant_ID = " & APPLICANT_ID & "")
    gconDMIS.Execute ("Delete From HRMS_APPLICANT_PAPER Where Applicant_ID = " & APPLICANT_ID & "")
    gconDMIS.Execute ("Delete From HRMS_APPLICANT_REFERENCE Where Applicant_ID = " & APPLICANT_ID & "")
    gconDMIS.Execute ("Delete From HRMS_APPLICANT_EMPLOYMENT_RECORD Where Applicant_ID = " & APPLICANT_ID & "")

    EMP_ENTRY_ID = 0
    PAPERS_ENTRY_ID = 0
    TRAINING_ENTRY_ID = 0
    PAPERS_ENTRY_ID = 0
End Function

Function DisplayEmploymentInListView()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim Item                                                          As ListItem

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_APPLICANT_EMPLOYMENT_RECORD Where Applicant_ID = " & _
                                 APPLICANT_ID & " Order By Entry_ID ASC")
    lsvEMP.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = lsvEMP.ListItems.Add(, , Null2String(RSTMP!NameOfCompany))
            Item.SubItems(1) = Null2String(RSTMP!ADDRESS)
            Item.SubItems(2) = Null2String(RSTMP!Posisyon)
            Item.SubItems(3) = Null2String(RSTMP!From_to)
            Item.SubItems(4) = Null2String(RSTMP!Entry_ID)

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Function

Function DisplayPapersInListView()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim Item                                                          As ListItem

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_APPLICANT_Paper Where Applicant_ID = " & _
                                 APPLICANT_ID & " Order By PaperID ASC")
    lsvDOC.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = lsvDOC.ListItems.Add(, , Null2String(RSTMP!PaperPass))
            Item.SubItems(1) = Null2String(RSTMP!PaperID)

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Function

Function DisplayReferenceInListView()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim Item                                                          As ListItem

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_APPLICANT_REFERENCE Where Applicant_ID = " & _
                                 APPLICANT_ID & " Order By Entry_ID ASC")
    LsvREF.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = LsvREF.ListItems.Add(, , Null2String(RSTMP!Name))
            Item.SubItems(1) = Null2String(RSTMP!ADDRESS)
            Item.SubItems(2) = Null2String(RSTMP!Posisyon)
            Item.SubItems(3) = Null2String(RSTMP!Telno)
            Item.SubItems(4) = Null2String(RSTMP!Entry_ID)

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Function

Function DisplayTrainInListView()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim Item                                                          As ListItem

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_APPLICANT_Train Where Applicant_ID = " & _
                                 APPLICANT_ID & " Order By Entry_ID ASC")
    lsvTRAIN.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = lsvTRAIN.ListItems.Add(, , Null2String(RSTMP!Training))
            Item.SubItems(1) = Null2String(RSTMP!monthYear)
            Item.SubItems(2) = Null2String(RSTMP!Place)
            Item.SubItems(3) = Null2String(RSTMP!Sponsor)
            Item.SubItems(4) = Null2String(RSTMP!Entry_ID)

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Function

Function CleanFormApplication()
    'PERSONAL INFORMATION
    txtPOSITION.Text = ""
    txtPER_ADD.Text = ""
    txtPER_EMAIL.Text = ""
    txtPER_LNAME.Text = ""
    txtPER_FNAME.Text = ""
    txtPER_MNAME.Text = ""
    txtPER_ADD.Text = ""
    cboPER_CITY.Text = ""
    txtPER_CNO.Text = ""
    txtPER_BPlace.Text = ""
    txtPER_Height.Text = ""
    txtPER_Weight.Text = ""
    txtPER_Religion.Text = ""
    txtPER_Citizenship.Text = ""

    'FAMILY INFORMATION
    txtFAMILY_SNAME.Text = ""
    txtFAMILY_SAGE.Text = ""
    txtFAMILY_FAGE.Text = ""
    txtFAMILY_MAGE.Text = ""
    txtFAMILY_SOCCU.Text = ""
    txtFAMILY_FNAME.Text = ""
    txtFAMILY_FOCCU.Text = ""
    txtFAMILY_MNAME.Text = ""
    txtFAMILY_MOCCU.Text = ""

    'EUDCATIONAL
    txt1st_Major.Text = ""
    txt1st_Grade.Text = ""
    txt1st_Ins.Text = ""
    txt1st_Add.Text = ""
    txt1st_Year.Text = ""

    txt2nd_Major.Text = ""
    txt2nd_grade.Text = ""
    txt2nd_Ins.Text = ""
    txt2nd_Add.Text = ""
    txt2nd_Year.Text = ""

    lsvEMP.ListItems.Clear
    lsvDOC.ListItems.Clear
    lsvTRAIN.ListItems.Clear
    LsvREF.ListItems.Clear

    cbo1st_Field.ListIndex = 0
    cbo2nd_Field.ListIndex = 0
    cbo1st_Level.ListIndex = 0
    cbo2nd_Level.ListIndex = 0
    cbo1st_Month.ListIndex = 0
    cbo2nd_Month.ListIndex = 0
    cbo1st_Field.ListIndex = 0
    cbo2nd_Field.ListIndex = 0
    cboPER_CSTATUS.ListIndex = 0
    cboPER_GENDER.ListIndex = 0
End Function

Function DisplayAllApplicantInListView(COND As String)
    Dim SQL                                                           As String
    Dim RSTMP                                                         As ADODB.Recordset
    Dim Item                                                          As ListItem

    lsvSEARCH.Enabled = False

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select Hired,Applicant_ID,FirstName,LastName From HRMS_APPLICANT_PERSONAL Where Hired = '" & COND & _
                                 "' Order By LastName ASC")

    lsvSEARCH.ListItems.Clear

    If Not RSTMP.EOF And Not RSTMP.BOF Then
        lsvSEARCH.Enabled = True
    End If

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = lsvSEARCH.ListItems.Add(, , RSTMP!lastname & "," & RSTMP!FIRSTNAME)
            Item.SubItems(1) = RSTMP!APPLICANT_ID

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Function

Function FillGenderandMaritalStatusPosition()
    cboPER_GENDER.AddItem "Male"
    cboPER_GENDER.ItemData(cboPER_GENDER.NewIndex) = 0
    cboPER_GENDER.AddItem "Female"
    cboPER_GENDER.ItemData(cboPER_GENDER.NewIndex) = 1
    cboPER_GENDER.ListIndex = 0

    cboPER_CSTATUS.AddItem "Single"
    cboPER_CSTATUS.ItemData(cboPER_CSTATUS.NewIndex) = 0
    cboPER_CSTATUS.AddItem "Married"
    cboPER_CSTATUS.ItemData(cboPER_CSTATUS.NewIndex) = 1
    cboPER_CSTATUS.AddItem "Separated"
    cboPER_CSTATUS.ItemData(cboPER_CSTATUS.NewIndex) = 2
    cboPER_CSTATUS.AddItem "Divorced"
    cboPER_CSTATUS.ItemData(cboPER_CSTATUS.NewIndex) = 3
    cboPER_CSTATUS.AddItem "Widowed"
    cboPER_CSTATUS.ItemData(cboPER_CSTATUS.NewIndex) = 4
    cboPER_CSTATUS.ListIndex = 0
End Function

Function ClickSearch()
    Call lsvSEARCH_Click
End Function

Function DisplayAllInformation()
    Dim SQL                                                           As String
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_APPLICANT_PERSONAL Where Applicant_ID = " & APPLICANT_ID & "")

    If Not (RSTMP.EOF And RSTMP.BOF) Then
        txtPER_LNAME.Text = Null2String(RSTMP!lastname)
        txtPER_FNAME.Text = Null2String(RSTMP!FIRSTNAME)
        txtPER_MNAME.Text = Null2String(RSTMP!MIDDLENAME)
        txtPER_ADD.Text = Null2String(RSTMP!ADDRESS)
        'txtPER_CITY.Text = Null2String(rsTMP!City)
        cboPER_CITY.Text = Null2String(RSTMP!City)
        txtPER_CNO.Text = Null2String(RSTMP!ContactNo)
        txtPER_EMAIL.Text = Null2String(RSTMP!EmailAdd)
        dtpPER_BirthDate.Day = 1
        dtpPER_BirthDate.Year = Year(Null2String(RSTMP!BIRTHDATE))
        dtpPER_BirthDate.Month = Month(Null2String(RSTMP!BIRTHDATE))
        dtpPER_BirthDate.Day = Day(Null2String(RSTMP!BIRTHDATE))
        cboPER_CSTATUS.Text = RSTMP!CIVILSTATUS
        cboPER_GENDER.Text = RSTMP!GENDER
        txtPER_BPlace.Text = Null2String(RSTMP!BIRTHPLACE)
        txtPER_Weight.Text = Null2String(RSTMP!WEIGHT)
        txtPER_Height.Text = Null2String(RSTMP!HEIGHT)
        txtPER_Religion.Text = Null2String(RSTMP!RELIGION)
        txtPER_Citizenship.Text = Null2String(RSTMP!Citizenship)
        APPLICANT_TYPE = Null2String(RSTMP!Type)
        txtPOSITION.Text = Null2String(RSTMP!PositionDesired)

        txtFAMILY_SNAME.Text = Null2String(RSTMP!SpouseName)
        txtFAMILY_SOCCU.Text = Null2String(RSTMP!SpouseOccupation)
        txtFAMILY_SAGE.Text = Null2String(RSTMP!SPOUSEAGE)
        txtFAMILY_FNAME.Text = Null2String(RSTMP!FatherName)
        txtFAMILY_FOCCU.Text = Null2String(RSTMP!FatherOccupation)
        txtFAMILY_FAGE.Text = Null2String(RSTMP!FATHERAGE)
        txtFAMILY_MNAME.Text = Null2String(RSTMP!MotherName)
        txtFAMILY_MOCCU.Text = Null2String(RSTMP!MotherOccupation)
        txtFAMILY_MAGE.Text = Null2String(RSTMP!MOTHERAGE)

        cbo1st_Level.Text = Null2String(RSTMP!HighestLevel1)
        cbo1st_Field.Text = Null2String(RSTMP!StudyFields1)
        txt1st_Major.Text = Null2String(RSTMP!Major1)
        txt1st_Grade.Text = Null2String(RSTMP!Grade1)
        txt1st_Ins.Text = Null2String(RSTMP!SchoolName1)
        txt1st_Add.Text = Null2String(RSTMP!SchoolAdd1)
        cbo1st_Month.Text = Null2String(RSTMP!GradMonth1)
        txt1st_Year.Text = Null2String(RSTMP!GradYear1)

        cbo2nd_Level.Text = Null2String(RSTMP!HighestLevel2)
        cbo2nd_Field.Text = Null2String(RSTMP!StudyFields2)
        txt2nd_Major.Text = Null2String(RSTMP!Major2)
        txt2nd_grade.Text = Null2String(RSTMP!Grade2)
        txt2nd_Ins.Text = Null2String(RSTMP!SchoolName2)
        txt2nd_Add.Text = Null2String(RSTMP!SchoolAdd2)
        cbo2nd_Month.Text = Null2String(RSTMP!GradMonth2)
        txt2nd_Year.Text = Null2String(RSTMP!GradYear2)
    End If

    SQL = "Select * From HRMS_APPLICANT_IMAGE_LOCATION Where Applicant_ID = " & APPLICANT_ID & ""
    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute(SQL)

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!ImageLocation) <> "" Then
            On Error Resume Next
            LoadPic imgAPP, Null2String(RSTMP!ImageLocation)
            lblPIC_LOC.Caption = Null2String(RSTMP!ImageLocation)
        Else
            LoadPic imgAPP, ""
        End If
    Else
        LoadPic imgAPP, ""
    End If

    Call DisplayEmploymentInListView
    Call DisplayTrainInListView
    Call DisplayPapersInListView
    Call DisplayReferenceInListView
End Function

Function CheckErrorEntry() As Boolean
    If txtPER_LNAME.Text = "" Then
        Call DisplayAMessage(INC_MSG, INC_TITLE)
        tbcApplication.SelectedItem = 0
        On Error Resume Next
        txtPER_LNAME.SetFocus
        CheckErrorEntry = True
        Exit Function
    End If
    If txtPER_FNAME.Text = "" Then
        Call DisplayAMessage(INC_MSG, INC_TITLE)
        tbcApplication.SelectedItem = 0
        CheckErrorEntry = True
        On Error Resume Next
        txtPER_FNAME.SetFocus
        Exit Function
    End If
    If txtPER_CNO.Text = "" Then
        Call DisplayAMessage(INC_MSG, INC_TITLE)
        tbcApplication.SelectedItem = 0
        CheckErrorEntry = True
        On Error Resume Next
        txtPER_CNO.SetFocus
        Exit Function
    End If
    If txtPOSITION.Text = "" Then
        Call DisplayAMessage(INC_MSG, INC_TITLE)
        tbcApplication.SelectedItem = 0
        On Error Resume Next
        txtPOSITION.SetFocus
        CheckErrorEntry = True
        Exit Function
    End If
End Function

Sub FillCityMunicipality()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select DISTINCT City From HRMS_APPLICANT_PERSONAL Where City IS NOT NULL Order By City ASC")
    cboPER_CITY.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboPER_CITY.AddItem RSTMP!City
            RSTMP.MoveNext
        Loop
        cboPER_CITY.ListIndex = 0
    End If
End Sub

Sub rsrefresh()
    'BTT -07032007
    Set RS = New ADODB.Recordset
    Call RS.Open("SELECT * FROM HRMS_APPLICANT_PERSONAL", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Sub StoreMemVars()
    'BTT - 07032007
    If Not RS.EOF And Not RS.BOF Then
        lblAPPNO.Caption = Null2String(RS!APPLICANT_ID)
        txtPER_LNAME.Text = Null2String(RS!lastname)
        txtPER_FNAME.Text = Null2String(RS!FIRSTNAME)
        txtPER_MNAME.Text = Null2String(RS!MIDDLENAME)
        txtPER_ADD.Text = Null2String(RS!ADDRESS)
        'txtPER_CITY.Text = Null2String(rsTMP!City)
        cboPER_CITY.Text = Null2String(RS!City)
        txtPER_CNO.Text = Null2String(RS!ContactNo)
        txtPER_EMAIL.Text = Null2String(RS!EmailAdd)
        dtpPER_BirthDate.Day = 1
        dtpPER_BirthDate.Year = Year(Null2String(RS!BIRTHDATE))
        dtpPER_BirthDate.Month = Month(Null2String(RS!BIRTHDATE))
        dtpPER_BirthDate.Day = Day(Null2String(RS!BIRTHDATE))
        cboPER_CSTATUS.Text = RS!CIVILSTATUS
        cboPER_GENDER.Text = RS!GENDER
        txtPER_BPlace.Text = Null2String(RS!BIRTHPLACE)
        txtPER_Weight.Text = Null2String(RS!WEIGHT)
        txtPER_Height.Text = Null2String(RS!HEIGHT)
        txtPER_Religion.Text = Null2String(RS!RELIGION)
        txtPER_Citizenship.Text = Null2String(RS!Citizenship)
        APPLICANT_TYPE = Null2String(RS!Type)
        txtPOSITION.Text = Null2String(RS!PositionDesired)

        txtFAMILY_SNAME.Text = Null2String(RS!SpouseName)
        txtFAMILY_SOCCU.Text = Null2String(RS!SpouseOccupation)
        txtFAMILY_SAGE.Text = Null2String(RS!SPOUSEAGE)
        txtFAMILY_FNAME.Text = Null2String(RS!FatherName)
        txtFAMILY_FOCCU.Text = Null2String(RS!FatherOccupation)
        txtFAMILY_FAGE.Text = Null2String(RS!FATHERAGE)
        txtFAMILY_MNAME.Text = Null2String(RS!MotherName)
        txtFAMILY_MOCCU.Text = Null2String(RS!MotherOccupation)
        txtFAMILY_MAGE.Text = Null2String(RS!MOTHERAGE)

        cbo1st_Level.Text = Null2String(RS!HighestLevel1)
        cbo1st_Field.Text = Null2String(RS!StudyFields1)
        txt1st_Major.Text = Null2String(RS!Major1)
        txt1st_Grade.Text = Null2String(RS!Grade1)
        txt1st_Ins.Text = Null2String(RS!SchoolName1)
        txt1st_Add.Text = Null2String(RS!SchoolAdd1)
        cbo1st_Month.Text = Null2String(RS!GradMonth1)
        txt1st_Year.Text = Null2String(RS!GradYear1)

        cbo2nd_Level.Text = Null2String(RS!HighestLevel2)
        cbo2nd_Field.Text = Null2String(RS!StudyFields2)
        txt2nd_Major.Text = Null2String(RS!Major2)
        txt2nd_grade.Text = Null2String(RS!Grade2)
        txt2nd_Ins.Text = Null2String(RS!SchoolName2)
        txt2nd_Add.Text = Null2String(RS!SchoolAdd2)
        cbo2nd_Month.Text = Null2String(RS!GradMonth2)
        txt2nd_Year.Text = Null2String(RS!GradYear2)



        Call DisplayEmploymentInListView
        Call DisplayTrainInListView
        Call DisplayPapersInListView
        Call DisplayReferenceInListView

    End If
End Sub

Private Sub cboPER_GENDER_Change()
    If cboPER_GENDER.Text = "Male" Then
        lblCAP(35).Caption = "Name Of Wife"
    Else
        lblCAP(35).Caption = "Name Of Spouse"
    End If
End Sub

Private Sub cboPER_GENDER_Click()
    Call cboPER_GENDER_Change
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "ACESS_DELETE", "AIS MASTER APPLICATION") = False Then Exit Sub
    '   gconDMIS.Execute "Delete from tblApplicant where"
End Sub

Private Sub cmdDOC_ADD_Click()
    picSaves.Visible = False
    SAVE_OR_EDIT_PAPERS = "SAVE"
    frmAISApplications.Enabled = False
    frmAISADD_DOC.cmdDelete.Enabled = False
    frmAISADD_DOC.Show
    On Error Resume Next
    frmAISADD_DOC.cboDOC.SetFocus
End Sub

Private Sub cmdEDIT_Click()
    If Function_Access(LOGID, "ACESS_EDIT", "AIS MASTER APPLICATION") = False Then Exit Sub
    If optACTIVE_EMP.Value = 0 Then
        If Not txtPER_LNAME.Text = "" Then
            Label1(1).Visible = True

            SAVE_OR_EDIT = "EDIT"
            Call EnablePicTab(True)

            Call DisplayRequiredFields(True)
            Call DisplayPictureButton(True)

            picAdds.Visible = False
            picSaves.Visible = True
            cmdSave.Caption = "&UPDATE"
            cmdEXAM.Visible = True

            optACTIVE_EMP.Enabled = False
        End If
    End If
End Sub

Private Sub cmdEMP_ADD_Click()
    If Not lsvEMP.ListItems.Count = 5 Then
        picSaves.Visible = False
        SAVE_OR_EDIT_EMP = "SAVE"
        frmAISApplications.Enabled = False
        frmAISEMPLOYMENT.Show
        frmAISEMPLOYMENT.cmdDelete.Enabled = False
        On Error Resume Next
        frmAISEMPLOYMENT.txtCOMP_Name.SetFocus
    End If
End Sub

Private Sub cmdEXAM_Click()
    If SAVE_OR_EDIT = "EDIT" Then
        frmAISApplications.Enabled = False
        frmAISEXAM_DISPLAY.Show

        Call frmAISEXAM_DISPLAY.DisplayApplicantInfoOnEXAMDISPLAY
    End If
End Sub

Private Sub cmdFind_Click()
    txtSEARCH.Text = ""
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "ACESS_ADD", "AIS MASTER APPLICATION") = False Then Exit Sub
    Label1(1).Visible = True
    LoadPic imgAPP, ""

    SAVE_OR_EDIT = "SAVE"
    Call CleanFormApplication
    Call GenerateApplicationNumber
    Call CleanAllUnsaveRecords
    Call DisplayRequiredFields(True)

    Call EnablePicTab(True)
    Call DisplayPictureButton(True)

    picSEARCH.Enabled = False

    cmdEXAM.Visible = False
    picAdds.Visible = False
    picSaves.Visible = True

    cmdSave.Caption = "&Save"
    tbcApplication.SelectedItem = 0
    On Error Resume Next
    txtPOSITION.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Label1(1).Visible = False

    SAVE_OR_EDIT = "SAVE"
    optACTIVE_EMP.Visible = True

    Call DisplayRequiredFields(False)
    Call EnablePicTab(False)
    Call DisplayPictureButton(False)

    picSaves.Visible = False
    picAdds.Visible = True
    cmdSave.Caption = "&SAVE"

    picSEARCH.Enabled = True
    cmdUPLOAD.Visible = False
    cmdEXAM.Visible = False
    optACTIVE_EMP.Enabled = True

    Call lsvSEARCH_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    frmMain.MousePointer = 11

    RS.MoveFirst
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next

    frmMain.MousePointer = 11

    RS.MoveLast
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    frmMain.MousePointer = 11

    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub cmdPIC_CHANGE_Click()
    lblPIC_LOC.Caption = ShowInsertpic(CDPIC, imgAPP)
End Sub

Private Sub cmdPIC_DELETE_Click()
    If MsgBox("Delete Applicant Picture", vbQuestion + vbYesNo + vbDefaultButton2, "Are you Sure") = vbYes Then
        lblPIC_LOC.Caption = ""
        LoadPic imgAPP, ""
        gconDMIS.Execute ("Delete From HRMS_APPLICANT_IMAGE_LOCATION Where Applicant_ID = " & APPLICANT_ID & "")
    End If
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    frmMain.MousePointer = 11

    RS.MovePrevious
    If RS.BOF Then
        RS.MoveNext
        ShowFirstRecordMsg
    End If
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub cmdPRINT_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_PRINT", "AIS MASTER APPLICATION") = False Then Exit Sub
    frmMain.MousePointer = 11

    Call PrintSQLReport(rptAPP, AIS_REPORT_PATH & "HRMS_ApplicationForm.rpt", "{PERSONAL.Applicant_ID} = " & APPLICANT_ID, AIS_REPORT_Connection, 1)

    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:14
Private Sub cmdSave_Click()
    On Error GoTo Errorcode

    Select Case SAVE_OR_EDIT
        Case "SAVE":    'Insert
            If MsgBox("Save New Applicant", vbQuestion + vbYesNo + vbDefaultButton1, "Are you Sure") = vbYes Then
                If Not CheckErrorEntry Then
                    frmMain.MousePointer = 11

                    Call SaveUpdateApplicant

                    If optINACTIVE_EMP Then Call DisplayAllApplicantInListView("NO")
                    If Not optINACTIVE_EMP Then Call DisplayAllApplicantInListView("YES")

                    cmdUPLOAD.Visible = False                 'Upload Button
                    picSaves.Visible = False
                    picAdds.Visible = True
                    SAVE_OR_EDIT = "EDIT"

                    Call EnablePicTab(False)
                    Call DisplayPictureButton(False)
                    Label1(1).Visible = False
                    optACTIVE_EMP.Enabled = True

                    picSEARCH.Enabled = True
                    tbcApplication.SelectedItem = 0

                    frmMain.MousePointer = 0
                End If
            End If

        Case "EDIT":    'Update
            If MsgBox("Save Edited Applicant", vbQuestion + vbYesNo + vbDefaultButton1, "Are you Sure") = vbYes Then
                If Not CheckErrorEntry Then
                    frmMain.MousePointer = 11

                    Call SaveUpdateApplicant
                    If optINACTIVE_EMP Then Call DisplayAllApplicantInListView("NO")
                    If Not optINACTIVE_EMP Then Call DisplayAllApplicantInListView("YES")

                    cmdSave.Caption = "SAVE"
                    optACTIVE_EMP.Visible = True

                    picSaves.Visible = False
                    picAdds.Visible = True

                    cmdUPLOAD.Visible = False                 'Upload Button
                    cmdEXAM.Visible = False

                    Call EnablePicTab(False)
                    Call DisplayPictureButton(False)
                    Label1(1).Visible = False
                    optACTIVE_EMP.Enabled = True

                    picSEARCH.Enabled = True
                    tbcApplication.SelectedItem = 0
                    frmMain.MousePointer = 0
                End If
            End If
    End Select
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub cmdTRAIN_ADD_Click()
    If Not lsvTRAIN.ListItems.Count = 5 Then
        picSaves.Visible = False
        SAVE_OR_EDIT_TRAINING = "SAVE"
        frmAISApplications.Enabled = False
        frmAISADD_TRAIN.Show
        frmAISADD_TRAIN.cmdTRAIN_DELETE.Enabled = False
        On Error Resume Next
        frmAISADD_TRAIN.txtTRAIN_TRAIN.SetFocus
    End If
End Sub

Private Sub cmdREF_ADD_Click()
    If Not LsvREF.ListItems.Count = 5 Then
        picSaves.Visible = False
        SAVE_OR_EDIT_REF = "SAVE"
        frmAISApplications.Enabled = False
        frmAISADD_REF.Show
        frmAISADD_REF.cmdREF_DELETE.Enabled = False
        On Error Resume Next
        frmAISADD_REF.txtREF_NAME.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If picSEARCH.Enabled = True Then
            txtSEARCH.Text = ""
            On Error Resume Next
            txtSEARCH.SetFocus
        End If
        Exit Sub
    End If
    If KeyCode = vbKeyF2 Then
        If picSaves.Visible = True Then
            Call cmdSave_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    frmMain.MousePointer = 11

    Call rsrefresh
    Call FillGenderandMaritalStatusPosition
    Call FillDegree
    Call FillFields
    Call FillMonth
    Call FillCityMunicipality

    Call optINACTIVE_EMP_Click
    Call EnablePicTab(False)
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub FillMonth()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_MONTHS")
    cbo1st_Month.Clear
    cbo2nd_Month.Clear
    cbo1st_Month.AddItem "-"
    cbo2nd_Month.AddItem "-"
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cbo1st_Month.AddItem RSTMP!Months
            cbo2nd_Month.AddItem RSTMP!Months

            RSTMP.MoveNext
        Loop
    End If
    cbo1st_Month.ListIndex = 0
    cbo2nd_Month.ListIndex = 0
End Sub

Private Sub FillDegree()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_DEGREE Order By Degree ASC")
    cbo1st_Level.Clear
    cbo2nd_Level.Clear
    cbo1st_Level.AddItem "-"
    cbo2nd_Level.AddItem "-"
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cbo1st_Level.AddItem RSTMP!DEGREE
            cbo2nd_Level.AddItem RSTMP!DEGREE

            RSTMP.MoveNext
        Loop
    End If
    cbo1st_Level.ListIndex = 0
    cbo2nd_Level.ListIndex = 0
End Sub

Private Sub FillFields()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_FIELDS Order By Fields ASC")
    cbo1st_Field.Clear
    cbo2nd_Field.Clear
    cbo1st_Field.AddItem "-"
    cbo2nd_Field.AddItem "-"
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cbo1st_Field.AddItem RSTMP!FIELDS
            cbo2nd_Field.AddItem RSTMP!FIELDS
            RSTMP.MoveNext
        Loop
    End If
    cbo1st_Field.ListIndex = 0
    cbo2nd_Field.ListIndex = 0
End Sub

Private Sub DisplayPictureButton(COND As Boolean)
    picTRAIN.Visible = COND
    picDOC.Visible = COND
    picEMP.Visible = COND
    PicREF.Visible = COND
End Sub

Private Sub EnablePicTab(COND As Boolean)
    picTAB_EDU.Enabled = COND
    picTAB_FAMILY.Enabled = COND
    picTAB_PER.Enabled = COND
    picTAB_TRAIN.Enabled = COND
    picTAB_EMP.Enabled = COND
    picTAB_DOC.Enabled = COND
    PicTAB_REF.Enabled = COND
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmAISADD_TRAIN
End Sub

Private Sub lsvDOC_DblClick()
    Dim INDEX                                                         As Integer
    If Not lsvDOC.ListItems.Count = 0 Then
        INDEX = CInt(lsvDOC.SelectedItem.INDEX)
        With lsvDOC
            picSaves.Visible = False
            SAVE_OR_EDIT_PAPERS = "EDIT"
            PAPERS_ENTRY_ID = CInt(.ListItems(INDEX).SubItems(1))

            frmAISApplications.Enabled = False
            frmAISADD_DOC.Show
            frmAISADD_DOC.cboDOC.Text = .ListItems(INDEX).Text
            On Error Resume Next
            frmAISADD_DOC.cboDOC.SetFocus
        End With
    End If
End Sub

Private Sub lsvEMP_DblClick()
    Dim INDEX                                                         As Integer

    If Not lsvEMP.ListItems.Count = 0 Then
        SAVE_OR_EDIT_EMP = "EDIT"
        INDEX = lsvEMP.SelectedItem.INDEX
        With lsvEMP
            picSaves.Visible = False
            EMP_ENTRY_ID = .ListItems(INDEX).SubItems(4)
            frmAISApplications.Enabled = False
            frmAISEMPLOYMENT.Show
            frmAISEMPLOYMENT.txtCOMP_Name.Text = .ListItems(INDEX).Text
            frmAISEMPLOYMENT.txtCOMP_ADD.Text = .ListItems(INDEX).SubItems(1)
            frmAISEMPLOYMENT.txtEMP_POS.Text = .ListItems(INDEX).SubItems(2)
            frmAISEMPLOYMENT.txtEMP_FROM.Text = .ListItems(INDEX).SubItems(3)
            On Error Resume Next
            frmAISEMPLOYMENT.txtCOMP_Name.SetFocus
        End With
    End If
End Sub

Private Sub LsvREF_DblClick()
    Dim INDEX                                                         As Long
    If Not LsvREF.ListItems.Count = 0 Then
        INDEX = LsvREF.SelectedItem.INDEX

        With LsvREF
            picSaves.Visible = False
            REFERENCE_ENTRY_ID = .ListItems(INDEX).SubItems(4)
            SAVE_OR_EDIT_REF = "EDIT"

            frmAISApplications.Enabled = False
            frmAISADD_REF.Show
            frmAISADD_REF.txtREF_NAME.Text = .ListItems(INDEX).Text
            frmAISADD_REF.txtREF_ADD.Text = .ListItems(INDEX).SubItems(1)
            frmAISADD_REF.txtREF_POS.Text = .ListItems(INDEX).SubItems(2)
            frmAISADD_REF.txtREF_TEL.Text = .ListItems(INDEX).SubItems(3)
            On Error Resume Next
            frmAISADD_REF.txtREF_NAME.SetFocus
        End With
    End If
End Sub

Private Sub lsvSEARCH_Click()
    Dim INDEX                                                         As Integer

    If Not lsvSEARCH.ListItems.Count = 0 Then
        frmMain.MousePointer = 11

        INDEX = lsvSEARCH.SelectedItem.INDEX
        With lsvSEARCH
            APPLICANT_ID = CLng(.ListItems(INDEX).SubItems(1))
            lblAPPNO.Caption = CLng(.ListItems(INDEX).SubItems(1))

            Call DisplayAllInformation
        End With

        frmMain.MousePointer = 0
    End If
End Sub

Private Sub SaveUpdateApplicant()
    Dim vtxtPosition                                                  As String
    Dim vtxtPER_LNAME As String, vtxtPER_FNAME As String, vtxtPER_MNAME As String
    Dim vdtpPER_BDATE                                                 As String
    Dim vcboPER_GENDER As String, vtxtPER_BPLACE As String, vtxtPER_RELIGION As String
    Dim vtxtPER_HEIGHT As String, vtxtPER_WEIGHT As String, vtxtPER_CITIZENSHIP As String
    Dim vHIRED As String, vtxtPER_CITY As String, vcboPER_CITY        As String
    Dim vtxtPER_ADDRESS As String, vtxtEMAIL As String, vtxtPER_TELNO As String
    Dim vtxtPER_AGE As Integer, VtxtFAMILY_FAGE As String, VtxtFAMILY_MAGE As String, vtxtFAMILY_SAGE As String
    Dim VTYPE                                                         As String

    Dim vcboPER_CSTATUS                                               As String

    VTYPE = N2Str2Null("")
    Dim VtxtFAMILY_SNAME As String, VtxtFAMILY_SOCCU As String, VtxtFAMILY_FNAME As String
    Dim VtxtFAMILY_FOCCU As String, VtxtFAMILY_MNAME                  As String
    Dim VtxtFAMILY_MOCCU                                              As String

    Dim vcbo1st_LEVEL As String, vcbo1st_FIELD As String, vtxt1st_Major As String, vtxt1st_GRADE As String, vtxt1st_INST As String, vtxt1st_ADD As String, vcbo1st_GRADMON As String, vtxt1st_YEAR As String
    Dim vcbo2nd_LEVEL As String, vcbo2nd_FIELD As String, vtxt2nd_Major As String, vtxt2nd_GRADE As String, vtxt2nd_INST As String, vtxt2nd_ADD As String, vcbo2nd_GRADMON As String, vtxt2nd_YEAR As String

    Dim vlblIMAGE_LOCATION                                            As String

    vtxtPER_AGE = DateDiff("yyyy", dtpPER_BirthDate, Date)
    VtxtFAMILY_FAGE = N2Str2Null(txtFAMILY_FAGE)
    VtxtFAMILY_MAGE = N2Str2Null(txtFAMILY_MAGE)
    vtxtFAMILY_SAGE = N2Str2Null(txtFAMILY_SAGE)

    vtxtPER_LNAME = N2Str2Null(txtPER_LNAME)
    vtxtPER_FNAME = N2Str2Null(txtPER_FNAME)
    vtxtPER_MNAME = N2Str2Null(txtPER_MNAME)
    vtxtPER_TELNO = N2Str2Null(txtPER_CNO)
    vdtpPER_BDATE = N2Str2Null(dtpPER_BirthDate)
    vcboPER_GENDER = N2Str2Null(cboPER_GENDER.Text)
    vcboPER_CSTATUS = N2Str2Null(cboPER_CSTATUS.Text)
    vtxtPER_BPLACE = N2Str2Null(txtPER_BPlace)
    vtxtPER_HEIGHT = N2Str2Null(txtPER_Height)
    vtxtPER_WEIGHT = N2Str2Null(txtPER_Weight)
    vtxtPER_RELIGION = N2Str2Null(txtPER_Religion)
    vtxtPER_CITIZENSHIP = N2Str2Null(txtPER_Citizenship)
    vHIRED = N2Str2Null("NO")

    vtxtPER_ADDRESS = N2Str2Null(Trim(txtPER_ADD.Text))
    'vtxtPER_CITY = N2Str2Null(txtPER_CITY.Text)
    vcboPER_CITY = N2Str2Null(cboPER_CITY.Text)
    vtxtPER_TELNO = N2Str2Null(txtPER_CNO.Text)
    vtxtEMAIL = N2Str2Null(txtPER_EMAIL.Text)
    vtxtPosition = N2Str2Null(txtPOSITION.Text)

    vcbo1st_LEVEL = N2Str2Null(cbo1st_Level)
    vcbo1st_FIELD = N2Str2Null(cbo1st_Field)
    vtxt1st_Major = N2Str2Null(txt1st_Major)
    vtxt1st_GRADE = N2Str2Null(txt1st_Grade)
    vtxt1st_INST = N2Str2Null(txt1st_Ins)
    vtxt1st_ADD = N2Str2Null(txt1st_Add)
    vcbo1st_GRADMON = N2Str2Null(cbo1st_Month)
    vtxt1st_YEAR = N2Str2Null(txt1st_Year)

    vcbo2nd_LEVEL = N2Str2Null(cbo2nd_Level)
    vcbo2nd_FIELD = N2Str2Null(cbo2nd_Field)
    vtxt2nd_Major = N2Str2Null(txt2nd_Major)
    vtxt2nd_GRADE = N2Str2Null(txt2nd_grade)
    vtxt2nd_INST = N2Str2Null(txt2nd_Ins)
    vtxt2nd_ADD = N2Str2Null(txt2nd_Add)
    vcbo2nd_GRADMON = N2Str2Null(cbo2nd_Month)
    vtxt2nd_YEAR = N2Str2Null(txt2nd_Year)

    VtxtFAMILY_SNAME = N2Str2Null(txtFAMILY_SNAME)
    VtxtFAMILY_SOCCU = N2Str2Null(txtFAMILY_SOCCU)
    VtxtFAMILY_FNAME = N2Str2Null(txtFAMILY_FNAME)

    VtxtFAMILY_FOCCU = N2Str2Null(txtFAMILY_FOCCU)
    VtxtFAMILY_MNAME = N2Str2Null(txtFAMILY_MNAME)

    VtxtFAMILY_MOCCU = N2Str2Null(txtFAMILY_MOCCU)

    vlblIMAGE_LOCATION = N2Str2Null(lblPIC_LOC)

    On Error GoTo ERROR_ON_SAVING

    If SAVE_OR_EDIT = "SAVE" Then
        gconDMIS.Execute ("Insert Into HRMS_APPLICANT_PERSONAL Values(" & APPLICANT_ID & _
                          "," & vtxtPER_FNAME & "," & vtxtPER_MNAME & "," & vtxtPER_LNAME & _
                          "," & vtxtPER_ADDRESS & "," & vcboPER_CITY & "," & vtxtPER_TELNO & "," & vtxtEMAIL & _
                          "," & vdtpPER_BDATE & "," & vtxtPER_AGE & _
                          "," & vcboPER_GENDER & "," & vcboPER_CSTATUS & _
                          "," & vtxtPER_BPLACE & _
                          "," & vtxtPER_HEIGHT & "," & vtxtPER_WEIGHT & _
                          "," & vtxtPER_RELIGION & "," & vtxtPER_CITIZENSHIP & _
                          "," & vHIRED & "," & VTYPE & _
                          "," & VtxtFAMILY_SNAME & "," & vtxtFAMILY_SAGE & "," & VtxtFAMILY_SOCCU & _
                          "," & VtxtFAMILY_FNAME & "," & VtxtFAMILY_FAGE & "," & VtxtFAMILY_FOCCU & _
                          "," & VtxtFAMILY_MNAME & "," & VtxtFAMILY_MAGE & "," & VtxtFAMILY_MOCCU & _
                          "," & vcbo1st_LEVEL & "," & vcbo1st_FIELD & "," & vtxt1st_Major & "," & vtxt1st_GRADE & "," & vtxt1st_INST & "," & vtxt1st_ADD & "," & vcbo1st_GRADMON & "," & vtxt1st_YEAR & _
                          "," & vcbo2nd_LEVEL & "," & vcbo2nd_FIELD & "," & vtxt2nd_Major & "," & vtxt2nd_GRADE & "," & vtxt2nd_INST & "," & vtxt2nd_ADD & "," & vcbo2nd_GRADMON & "," & vtxt2nd_YEAR & _
                          "," & vtxtPosition & ",'" & Date & "')")

        gconDMIS.Execute ("Insert Into HRMS_APPLICANT_IMAGE_LOCATION Values(" & APPLICANT_ID & _
                          "," & vlblIMAGE_LOCATION & ")")
    Else
        gconDMIS.Execute ("Update HRMS_APPLICANT_PERSONAL Set FirstName = " & vtxtPER_FNAME & ",LastName = " & vtxtPER_LNAME & ",MiddleName = " & vtxtPER_MNAME & _
                          ",Address = " & vtxtPER_ADDRESS & ",City = " & vcboPER_CITY & ",ContactNo = " & vtxtPER_TELNO & ",EmailAdd = " & vtxtEMAIL & _
                          ",Birthdate = " & vdtpPER_BDATE & ",Age = " & vtxtPER_AGE & ",BirthPlace = " & vtxtPER_BPLACE & _
                          ",Gender = " & vcboPER_GENDER & ",CivilStatus = " & vcboPER_CSTATUS & _
                          ",Height = " & vtxtPER_HEIGHT & ",Weight = " & vtxtPER_WEIGHT & _
                          ",Religion = " & vtxtPER_RELIGION & ",CitizenShip = " & vtxtPER_CITIZENSHIP & _
                          ",SpouseName = " & VtxtFAMILY_SNAME & ",SpouseAge = " & vtxtFAMILY_SAGE & ",SpouseOccupation = " & VtxtFAMILY_SOCCU & _
                          ",FatherName = " & VtxtFAMILY_FNAME & ",FatherAge = " & VtxtFAMILY_FAGE & ",FatherOccupation = " & VtxtFAMILY_FOCCU & _
                          ",Mothername = " & VtxtFAMILY_MNAME & ",MotherAge = " & VtxtFAMILY_MAGE & ",MotherOccupation = " & VtxtFAMILY_MOCCU & _
                          ",HighestLevel1 = " & vcbo1st_LEVEL & ",StudyFields1 = " & vcbo1st_FIELD & ",Major1 = " & vtxt1st_Major & ",Grade1 = " & vtxt1st_GRADE & ",SchoolName1 = " & vtxt1st_INST & ",SchoolAdd1 = " & vtxt1st_ADD & ",GradMonth1 = " & vcbo1st_GRADMON & ",GradYear1 = " & vtxt1st_YEAR & _
                          ",HighestLevel2 = " & vcbo2nd_LEVEL & ",StudyFields2 = " & vcbo2nd_FIELD & ",Major2 = " & vtxt2nd_Major & ",Grade2 = " & vtxt2nd_GRADE & ",SchoolName2 = " & vtxt2nd_INST & ",SchoolAdd2 = " & vtxt2nd_ADD & ",GradMonth2 = " & vcbo2nd_GRADMON & ",GradYear2 = " & vtxt2nd_YEAR & _
                          ",PositionDesired = " & vtxtPosition & _
                        " Where Applicant_ID = " & APPLICANT_ID & "")

        gconDMIS.Execute ("Update HRMS_APPLICANT_IMAGE_LOCATION Set ImageLocation = " & _
                          vlblIMAGE_LOCATION & " Where Applicant_ID = " & APPLICANT_ID & "")
    End If

    Exit Sub

ERROR_ON_SAVING:
    ShowVBError
    Exit Sub
End Sub

Private Sub lsvSEARCH_DblClick()
    If Not lsvSEARCH.ListItems.Count = 0 Then
        frmMain.MousePointer = 11
        If optACTIVE_EMP.Value = True Then

        Else
            optACTIVE_EMP.Visible = False
            Call cmdEDIT_Click

            tbcApplication.SelectedItem = 0
            On Error Resume Next
            txtPOSITION.SetFocus
        End If
        frmMain.MousePointer = 0
    End If

    frmMain.MousePointer = 0
End Sub

Private Sub lsvSEARCH_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call lsvSEARCH_Click
End Sub

Private Sub lsvSEARCH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvSEARCH_DblClick
End Sub

Private Sub lsvTRAIN_DblClick()
    Dim INDEX                                                         As Integer

    If Not lsvTRAIN.ListItems.Count = 0 Then
        INDEX = lsvTRAIN.SelectedItem.INDEX
        With lsvTRAIN
            picSaves.Visible = False

            SAVE_OR_EDIT_TRAINING = "EDIT"
            TRAINING_ENTRY_ID = .ListItems(INDEX).SubItems(4)
            frmAISApplications.Enabled = False
            frmAISADD_TRAIN.Show
            frmAISADD_TRAIN.txtTRAIN_TRAIN.Text = .ListItems(INDEX).Text
            frmAISADD_TRAIN.txtTRAIN_MonthYear.Text = .ListItems(INDEX).SubItems(1)
            frmAISADD_TRAIN.txtTRAIN_PLACE.Text = .ListItems(INDEX).SubItems(2)
            frmAISADD_TRAIN.txtTRAIN_SPONSOR.Text = .ListItems(INDEX).SubItems(3)
            On Error Resume Next
            frmAISADD_TRAIN.txtTRAIN_TRAIN.SetFocus
        End With
    End If
End Sub

Private Sub optACTIVE_EMP_Click()
    If optINACTIVE_EMP.Value Then optINACTIVE_EMP.Value = False
    optACTIVE_EMP = True
    cmdUPLOAD.Visible = False

    Call DisplayAllApplicantInListView("YES")
    Call lsvSEARCH_Click
End Sub

Private Sub optINACTIVE_EMP_Click()
    frmMain.MousePointer = 11

    If optACTIVE_EMP.Value Then optACTIVE_EMP.Value = False
    optINACTIVE_EMP = True

    Call DisplayAllApplicantInListView("NO")
    Call lsvSEARCH_Click

    frmMain.MousePointer = 0
End Sub

Private Sub TXTSEARCH_Change()
    Dim Keyword                                                       As String
    Dim RSTMP                                                         As ADODB.Recordset
    Dim Item                                                          As ListItem

    lsvSEARCH.Enabled = False
    Keyword = Trim(txtSEARCH.Text)
    frmMain.MousePointer = 11
    If optACTIVE_EMP.Value = True Then
        On Error GoTo ERROR
        Set RSTMP = gconDMIS.Execute("Select Applicant_ID,LastName,FirstName From HRMS_APPLICANT_PERSONAL Where Hired = '" & "YES" & _
                                     "' And (LastName Like '%" & Keyword & "%' or Firstname Like '%" & Keyword & "%') Order By LastName ASC")
        lsvSEARCH.ListItems.Clear
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            Do While Not RSTMP.EOF
                Set Item = lsvSEARCH.ListItems.Add(, , RSTMP!lastname & ", " & RSTMP!FIRSTNAME)
                Item.SubItems(1) = RSTMP!APPLICANT_ID

                RSTMP.MoveNext
            Loop
        Else
            lsvSEARCH.ListItems.Clear
        End If

        If txtSEARCH.Text = "" Then lsvSEARCH.ListItems.Clear
    Else
        On Error GoTo ERROR
        Set RSTMP = gconDMIS.Execute("Select Applicant_ID,LastName,FirstName From HRMS_APPLICANT_PERSONAL Where Hired = '" & "NO" & _
                                     "' And (LastName Like '%" & Keyword & "%' Or FirstName Like '%" & Keyword & "%') Order By LastName ASC")
        lsvSEARCH.ListItems.Clear
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            Do While Not RSTMP.EOF
                Set Item = lsvSEARCH.ListItems.Add(, , RSTMP!lastname & ", " & RSTMP!FIRSTNAME)
                Item.SubItems(1) = RSTMP!APPLICANT_ID

                RSTMP.MoveNext
            Loop
        Else
            lsvSEARCH.ListItems.Clear
        End If
    End If
    frmMain.MousePointer = 0
    Exit Sub

ERROR:
    frmMain.MousePointer = 0
End Sub


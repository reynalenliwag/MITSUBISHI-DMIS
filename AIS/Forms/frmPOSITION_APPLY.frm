VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmAISPOSITION_APPLY 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPLOAD APPLICANT"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOSITION_APPLY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10290
   Begin VB.Frame Frame2 
      Caption         =   "APPLICANT INFORMATION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   3270
      TabIndex        =   8
      Top             =   3000
      Width           =   6915
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5970
         Picture         =   "frmPOSITION_APPLY.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Exit Window"
         Top             =   3900
         Width           =   855
      End
      Begin VB.CommandButton cmdUPLOAD 
         Caption         =   "&Upload"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5130
         Picture         =   "frmPOSITION_APPLY.frx":6DA4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Upload Applicants"
         Top             =   3900
         Width           =   855
      End
      Begin XtremeSuiteControls.TabControl tcAPP 
         Height          =   3525
         Left            =   30
         TabIndex        =   9
         Top             =   270
         Width           =   6825
         _Version        =   655364
         _ExtentX        =   12039
         _ExtentY        =   6218
         _StockProps     =   64
         AllowReorder    =   -1  'True
         Appearance      =   9
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   3
         SelectedItem    =   2
         Item(0).Caption =   "Personal Information"
         Item(0).Tooltip =   "Personal Information"
         Item(0).ControlCount=   12
         Item(0).Control(0)=   "lblCAP(23)"
         Item(0).Control(1)=   "lblCAP(24)"
         Item(0).Control(2)=   "lblCAP(25)"
         Item(0).Control(3)=   "lblCAP(26)"
         Item(0).Control(4)=   "imgAPP"
         Item(0).Control(5)=   "lblAPP(0)"
         Item(0).Control(6)=   "lblAPP(1)"
         Item(0).Control(7)=   "lblAPP(2)"
         Item(0).Control(8)=   "lblAPP(3)"
         Item(0).Control(9)=   "Shape1"
         Item(0).Control(10)=   "lblCAP(14)"
         Item(0).Control(11)=   "lblAPP(10)"
         Item(1).Caption =   "Educational Background"
         Item(1).Tooltip =   "Educational Background"
         Item(1).ControlCount=   12
         Item(1).Control(0)=   "lblAPP(4)"
         Item(1).Control(1)=   "lblAPP(5)"
         Item(1).Control(2)=   "lblAPP(6)"
         Item(1).Control(3)=   "lblAPP(7)"
         Item(1).Control(4)=   "lblCAP(6)"
         Item(1).Control(5)=   "lblCAP(7)"
         Item(1).Control(6)=   "lblCAP(8)"
         Item(1).Control(7)=   "lblCAP(9)"
         Item(1).Control(8)=   "lblCAP(10)"
         Item(1).Control(9)=   "lblCAP(11)"
         Item(1).Control(10)=   "lblAPP(8)"
         Item(1).Control(11)=   "lblAPP(9)"
         Item(2).Caption =   "Papers Pass"
         Item(2).Tooltip =   "Papers Pass"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "lsvDOC"
         Item(2).Control(1)=   "lsvEXP"
         Begin MSComctlLib.ListView lsvDOC 
            Height          =   2565
            Left            =   60
            TabIndex        =   16
            Top             =   450
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   4524
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
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Documents"
               Object.Width           =   8819
            EndProperty
         End
         Begin MSComctlLib.ListView lsvEXP 
            Height          =   3795
            Left            =   -69900
            TabIndex        =   17
            Top             =   500
            Width           =   6705
            _ExtentX        =   11827
            _ExtentY        =   6694
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   10
            Left            =   -68350
            TabIndex        =   54
            Top             =   3090
            Visible         =   0   'False
            Width           =   4965
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Applying for:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   -69760
            TabIndex        =   53
            Top             =   3150
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   9
            Left            =   -68530
            TabIndex        =   38
            Top             =   2940
            Visible         =   0   'False
            Width           =   5265
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   8
            Left            =   -68530
            TabIndex        =   37
            Top             =   2550
            Visible         =   0   'False
            Width           =   5265
         End
         Begin VB.Label lblCAP 
            BackStyle       =   0  'Transparent
            Caption         =   "2nd Highest Level"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   11
            Left            =   -69670
            TabIndex        =   36
            Top             =   2130
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Study Fields"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   -69670
            TabIndex        =   35
            Top             =   2640
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Major"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   -69100
            TabIndex        =   34
            Top             =   3030
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lblCAP 
            BackStyle       =   0  'Transparent
            Caption         =   "1st Highest Level"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   8
            Left            =   -69610
            TabIndex        =   33
            Top             =   600
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Study Fields"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   -69700
            TabIndex        =   32
            Top             =   1110
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Major"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   -69160
            TabIndex        =   31
            Top             =   1470
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   7
            Left            =   -68530
            TabIndex        =   30
            Top             =   2130
            Visible         =   0   'False
            Width           =   5265
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   6
            Left            =   -68530
            TabIndex        =   29
            Top             =   1380
            Visible         =   0   'False
            Width           =   5265
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   5
            Left            =   -68530
            TabIndex        =   28
            Top             =   1020
            Visible         =   0   'False
            Width           =   5265
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   4
            Left            =   -68530
            TabIndex        =   27
            Top             =   630
            Visible         =   0   'False
            Width           =   5265
         End
         Begin VB.Image imgAPP 
            Height          =   1155
            Left            =   -64750
            Stretch         =   -1  'True
            Top             =   570
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Shape Shape1 
            Height          =   1335
            Left            =   -64840
            Top             =   480
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   3
            Left            =   -68350
            TabIndex        =   24
            Top             =   2700
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   2
            Left            =   -68350
            TabIndex        =   23
            Top             =   2340
            Visible         =   0   'False
            Width           =   4965
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   -68350
            TabIndex        =   22
            Top             =   1950
            Visible         =   0   'False
            Width           =   4965
         End
         Begin VB.Label lblAPP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   -68350
            TabIndex        =   21
            Top             =   1590
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   -68980
            TabIndex        =   13
            Top             =   2790
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   -69640
            TabIndex        =   12
            Top             =   2430
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   -69640
            TabIndex        =   11
            Top             =   2040
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Applicant ID:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   -69790
            TabIndex        =   10
            Top             =   1680
            Visible         =   0   'False
            Width           =   1125
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "CHOOSE APPLICANT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   330
      TabIndex        =   25
      Top             =   3000
      Width           =   2805
      Begin VB.TextBox txtSEARCH 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   390
         Width           =   2655
      End
      Begin MSComctlLib.ListView lsvAPP 
         Height          =   3285
         Left            =   90
         TabIndex        =   5
         Top             =   870
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5794
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
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   9
         EndProperty
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - SEARCH"
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
         Index           =   12
         Left            =   120
         TabIndex        =   51
         Top             =   4200
         Width           =   1110
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "REQUIRED "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   4500
      TabIndex        =   14
      Top             =   120
      Width           =   5685
      Begin XtremeSuiteControls.TabControl tcPOSITION 
         Height          =   2445
         Left            =   30
         TabIndex        =   15
         Top             =   270
         Width           =   5565
         _Version        =   655364
         _ExtentX        =   9816
         _ExtentY        =   4313
         _StockProps     =   64
         AllowReorder    =   -1  'True
         Appearance      =   9
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   3
         Item(0).Caption =   "Personality Requirement"
         Item(0).Tooltip =   "Personality Requirement"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "Frame5"
         Item(0).Control(1)=   "Frame6"
         Item(1).Caption =   "Required Education"
         Item(1).Tooltip =   "Required Education"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lsvREQ_EDU"
         Item(2).Caption =   "Required Papers"
         Item(2).Tooltip =   "Required Papers"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "lsvREQ_DOC"
         Item(2).Control(1)=   "lsvREQ_EXP"
         Begin VB.Frame Frame6 
            Caption         =   "Position Info."
            Height          =   1125
            Left            =   90
            TabIndex        =   45
            Top             =   390
            Width           =   4635
            Begin VB.Label lblINFO 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   2040
               TabIndex        =   49
               Top             =   720
               Width           =   1635
            End
            Begin VB.Label lblINFO 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   2040
               TabIndex        =   48
               Top             =   330
               Width           =   1635
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vacant Position"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   330
               TabIndex        =   47
               Top             =   750
               Width           =   1305
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Needed Applicant"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   46
               Top             =   360
               Width           =   1485
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Required Age"
            Height          =   795
            Left            =   90
            TabIndex        =   40
            Top             =   1560
            Width           =   4665
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "To"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   2880
               TabIndex        =   44
               Top             =   390
               Width           =   210
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "From"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   300
               TabIndex        =   43
               Top             =   360
               Width           =   435
            End
            Begin VB.Label lblREQ 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   0
               Left            =   1380
               TabIndex        =   42
               Top             =   330
               Width           =   945
            End
            Begin VB.Label lblREQ 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   1
               Left            =   3420
               TabIndex        =   41
               Top             =   330
               Width           =   915
            End
         End
         Begin MSComctlLib.ListView lsvREQ_DOC 
            Height          =   1935
            Left            =   -69955
            TabIndex        =   18
            Top             =   435
            Visible         =   0   'False
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   3413
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
               Text            =   "Documents"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Notes"
               Object.Width           =   3528
            EndProperty
         End
         Begin MSComctlLib.ListView lsvREQ_EXP 
            Height          =   3765
            Left            =   -1.39900e5
            TabIndex        =   19
            Top             =   500
            Visible         =   0   'False
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   6641
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lsvREQ_EDU 
            Height          =   1935
            Left            =   -69955
            TabIndex        =   20
            Top             =   435
            Visible         =   0   'False
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   3413
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Degree"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Course"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Note"
               Object.Width           =   3528
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CHOOSE POSITION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   90
      TabIndex        =   7
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox cboSAL 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmPOSITION_APPLY.frx":74D5
         Left            =   1740
         List            =   "frmPOSITION_APPLY.frx":74DF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1620
         Width           =   2235
      End
      Begin VB.ComboBox cboDept 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmPOSITION_APPLY.frx":74F3
         Left            =   1740
         List            =   "frmPOSITION_APPLY.frx":74F5
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1170
         Width           =   2265
      End
      Begin VB.ComboBox cboTYPE 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmPOSITION_APPLY.frx":74F7
         Left            =   1755
         List            =   "frmPOSITION_APPLY.frx":74F9
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
         Width           =   2235
      End
      Begin VB.ComboBox cboPOSITION 
         Appearance      =   0  'Flat
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   990
         TabIndex        =   55
         Top             =   480
         Width           =   660
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   585
         TabIndex        =   52
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   630
         TabIndex        =   50
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   26
         Top             =   840
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmAISPOSITION_APPLY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub tmp()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim POS_ID                                                        As Integer
    Dim ITEM                                                          As ListItem

    POS_ID = CInt(Right(cboPosition, 3))
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION Where POS_ID = " & POS_ID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        lblINFO(0).Caption = RSTMP!PositionAvailable
        lblINFO(1).Caption = RSTMP!PositionTaken
        lblREQ(0).Caption = RSTMP!fromAge
        lblREQ(1).Caption = RSTMP!toAge
    End If

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION_EDUCATION Where POS_ID = " & POS_ID & "")
    lsvREQ_EDU.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvREQ_EDU.ListItems.Add(, , Null2String(RSTMP!DEGREE))
            ITEM.SubItems(1) = RSTMP!GRADE
            ITEM.SubItems(2) = Null2String(RSTMP!NOTES)

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION_DOCUMENTS Where POS_ID = " & POS_ID & "")
    lsvREQ_DOC.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvREQ_DOC.ListItems.Add(, , Null2String(RSTMP!DocumentType))
            ITEM.SubItems(1) = Null2String(RSTMP!NOTES)

            RSTMP.MoveNext
        Loop
    End If
End Sub

Sub FillCboDepartment()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim ZEROS                                                         As String

    Set RSTMP = gconDMIS.Execute("Select * from HRMS_Department Order by ID ASC")
    cboDept.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            If Len(RSTMP!ID) = 1 Then ZEROS = "0"

            cboDept.AddItem RSTMP!DEPTNAME & " - " & ZEROS & RSTMP!ID
            RSTMP.MoveNext
        Loop
        cboDept.ListIndex = 0
    End If

    Set RSTMP = Nothing
End Sub

Sub FillCBOType()
    cboType.AddItem "Contractual"
    cboType.ItemData(cboType.NewIndex) = 0
    cboType.AddItem "Allowance Base"
    cboType.ItemData(cboType.NewIndex) = 1
    cboType.AddItem "Probationary"
    cboType.ItemData(cboType.NewIndex) = 2
    cboType.ListIndex = 0
End Sub

Sub FillCboPosition()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim SZERO                                                         As String

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION Where PositionAvailable > PositionTaken And DateInactive <= '" & Date & "' Order By PositionDesc ASC")
    cboPosition.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            If Len(RSTMP!POS_ID) = 1 Then SZERO = "0"

            cboPosition.AddItem RSTMP!PositionDesc & " - " & SZERO & RSTMP!POS_ID
            RSTMP.MoveNext
        Loop
        cboPosition.ListIndex = 0
    End If
End Sub

Private Sub cboPOSITION_Change()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim POS_ID                                                        As Integer
    Dim ITEM                                                          As ListItem

    frmMain.MousePointer = 11
    POS_ID = CInt(Right(cboPosition, 3))
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION Where POS_ID = " & POS_ID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        lblINFO(0).Caption = RSTMP!PositionAvailable
        lblINFO(1).Caption = RSTMP!PositionTaken
        lblREQ(0).Caption = RSTMP!fromAge
        lblREQ(1).Caption = RSTMP!toAge
    End If

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION_EDUCATION Where POS_ID = " & POS_ID & "")
    lsvREQ_EDU.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvREQ_EDU.ListItems.Add(, , Null2String(RSTMP!DEGREE))
            ITEM.SubItems(1) = RSTMP!FIELDS
            ITEM.SubItems(2) = Null2String(RSTMP!NOTES)

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION_DOCUMENTS Where POS_ID = " & POS_ID & "")
    lsvREQ_DOC.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvREQ_DOC.ListItems.Add(, , Null2String(RSTMP!DocumentType))
            ITEM.SubItems(1) = Null2String(RSTMP!NOTES)

            RSTMP.MoveNext
        Loop
    End If
    frmMain.MousePointer = 0
End Sub

Private Sub cboPOSITION_Click()
    Call cboPOSITION_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUPLOAD_Click()
    Dim Index                                                         As Integer
    Dim RSTMP                                                         As ADODB.Recordset
    Dim rsINT                                                         As ADODB.Recordset
    Dim rsPOS                                                         As ADODB.Recordset
    Dim ETYPE                                                         As String

    If Function_Access(LOGID, "ACESS_PROCESS", "UPLOAD APPLICANT") = False Then Exit Sub
    On Error GoTo Errorcode:

    frmMain.MousePointer = 11

    If lblINFO(1).Caption = lblINFO(0).Caption Then
        MsgBox "No Vacancy for this Job", vbInformation, "Upload Applicant"
        cboPosition.SetFocus
        frmMain.MousePointer = 0
        Exit Sub
    End If

    If Not lblAPP(0).Caption = "" Then
        If Not cboType.Text = "" Then
            If MsgBox("Upload Applicant", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                If cboType.Text = "Contractual" Then ETYPE = "C"
                If cboType.Text = "Allowance Base" Then ETYPE = "A"
                If cboType.Text = "Probationary" Then ETYPE = "E"

                Set rsINT = gconDMIS.Execute("Select * From HRMS_APPLICANT_INTERVIEW_SCHEDULE Where Remarks = '" & _
                                             "Passed" & "' And Applicant_ID = " & lblAPP(0).Caption & "")
                If Not (rsINT.BOF And rsINT.EOF) Then
                    frmMain.MousePointer = 11

                    gconDMIS.Execute ("Update HRMS_APPLICANT_PERSONAL Set Type = '" & ETYPE & "',Hired = '" & "YES" & "' Where Applicant_ID = " & lblAPP(0).Caption & "")
                    Call UploadApplicantToEmployee

                    GoTo UPDATE_POSITION

                    frmMain.MousePointer = 0
                Else
                    If MsgBox("Applicant Not Yet Pass the Interview, Continue", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                        frmMain.MousePointer = 11

                        gconDMIS.Execute ("Update HRMS_APPLICANT_PERSONAL Set Type = '" & ETYPE & "', Hired = '" & "YES" & "' Where Applicant_ID = " & lblAPP(0).Caption & "")
                        Call UploadApplicantToEmployee

                        GoTo UPDATE_POSITION

                        frmMain.MousePointer = 0
                    End If
                End If
            End If
        Else
            MsgBox "Choose a Employee Type", vbExclamation, "Upload Applicant"
            On Error Resume Next
            cboType.SetFocus
        End If
    Else
        MsgBox "Choose A Applicant to Upload", vbExclamation, "Upload Applicant"
        On Error Resume Next
        TXTSEARCH.SetFocus
    End If

    frmMain.MousePointer = 0
    Exit Sub


UPDATE_POSITION:

    Set rsPOS = gconDMIS.Execute("Select * From HRMS_POSITION Where POS_ID = " & Right(cboPosition, 2) & "")
    If Not (rsPOS.BOF And rsPOS.EOF) Then
        gconDMIS.Execute ("Update HRMS_POSITION Set PositionAvailable = " & rsPOS!PositionAvailable & _
                          ",PositionTaken = " & rsPOS!PositionTaken + 1 & _
                        " Where Pos_ID = " & Right(cboPosition, 2) & "")
    End If

    Call FillCboPosition
    Call txtSEARCH_Change

    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub UploadApplicantToEmployee()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim rsAPP                                                         As New ADODB.Recordset
    Dim rsPIC                                                         As New ADODB.Recordset
    Dim EMPNO                                                         As String
    Dim ID                                                            As Double
    Dim IMAGE_LOCATION                                                As String
    Dim EMPSTATUS                                                     As String
    Dim DEPTNO                                                        As Integer
    Dim lastname, FIRSTNAME, MIDDLENAME                               As String
    Dim ADDRESS, TELEPHONE, BIRTHDATE, SEX, STATUS, BIRTHPLACE        As String
    Dim HEIGHT, WEIGHT, RELIGION, CITIZEN, MYPOSITION                 As String
    Dim SPOUSE, SPOUSEAGE, SOCCUPATION                                As String
    Dim FATHER, FATHERAGE, FOCCUPATION, MOTHER, MOTHERAGE, MOCCUPATION, PICFILNAME, EMPLEVEL As String

    If cboSAL.Text = "Monthly" Then EMPSTATUS = "M"
    If cboSAL.Text = "Daily" Then EMPSTATUS = "D"

    'Get the last EMPNO--------------------------------------------------------------
    Set RSTMP = gconDMIS.Execute("Select EmpNo,ID From HRMS_EmpInfo Order By EmpNo ASC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveLast
    End If
    EMPNO = CDbl(Right(RSTMP!EMPNO, 4)) + 1
    If Len(EMPNO) = 1 Then EMPNO = "00" & EMPNO
    If Len(EMPNO) = 2 Then EMPNO = "0" & EMPNO
    ID = RSTMP!ID + 1
    'Get the last EMPNO--------------------------------------------------------------

    'Personal Info---------------------------------------------------------------------------
    Set rsAPP = gconDMIS.Execute("Select * From HRMS_Applicant_Personal Where Applicant_Id = " & lblAPP(0).Caption & "")
    If Not (rsAPP.BOF And rsAPP.EOF) Then
        lastname = N2Str2Null(rsAPP!lastname)
        FIRSTNAME = N2Str2Null(rsAPP!FIRSTNAME)
        MIDDLENAME = N2Str2Null(rsAPP!MIDDLENAME)
        ADDRESS = N2Str2Null(rsAPP!ADDRESS)
        TELEPHONE = N2Str2Null(rsAPP!ContactNo)
        BIRTHDATE = N2Str2Null(rsAPP!BIRTHDATE)
        If rsAPP!GENDER = "Male" Then SEX = "M"
        If rsAPP!GENDER = "Female" Then SEX = "F"
        STATUS = N2Str2Null(rsAPP!CIVILSTATUS)
        BIRTHPLACE = N2Str2Null(rsAPP!BIRTHPLACE)
        HEIGHT = N2Str2Null(rsAPP!HEIGHT)
        WEIGHT = N2Str2Null(rsAPP!WEIGHT)
        RELIGION = N2Str2Null(rsAPP!RELIGION)
        CITIZEN = N2Str2Null(rsAPP!Citizenship)
        SPOUSE = N2Str2Null(rsAPP!SpouseName)
        SPOUSEAGE = N2Str2Null(rsAPP!SPOUSEAGE)
        SOCCUPATION = N2Str2Null(rsAPP!SpouseOccupation)
        FATHER = N2Str2Null(rsAPP!FatherName)
        FATHERAGE = N2Str2Null(rsAPP!FATHERAGE)
        FOCCUPATION = N2Str2Null(rsAPP!FatherOccupation)
        MOTHER = N2Str2Null(rsAPP!MotherName)
        MOTHERAGE = N2Str2Null(rsAPP!MOTHERAGE)
        MOCCUPATION = N2Str2Null(rsAPP!MotherOccupation)
    End If
    '--------------------------------------------------------------------------------
    'Image Location--------------------------------------------------------------------------------
    Set rsPIC = gconDMIS.Execute("Select * From HRMS_Applicant_Image_Location Where Applicant_Id = " & lblAPP(0).Caption & "")
    If Not (rsPIC.BOF And rsPIC.EOF) Then
        IMAGE_LOCATION = N2Str2Null(rsPIC!ImageLocation)
    End If
    '--------------------------------------------------------------------------------
    DEPTNO = Right(cboDept, 2)
    MYPOSITION = Left(cboPosition, Len(cboPosition) - 5)

    gconDMIS.Execute ("Insert Into HRMS_EmpInfo (empno,deptcode,empLevel,lastname,firstname,middlename,address,telephone,birthdate,sex,status,birthplace" & _
                      ",height,weight,religion,citizen,[position]" & _
                      ",empstatus,datehired,spouse,spouseage,soccupation" & _
                      ",father,fatherage,foccupation,mother,motherage,moccupation,ActiveInactive,picfilname)" & _
                    " Values('" & EMPNO & "','" & DEPTNO & "','" & "E" & _
                      "'," & lastname & "," & FIRSTNAME & "," & MIDDLENAME & "," & ADDRESS & "," & TELEPHONE & "," & BIRTHDATE & ",'" & SEX & "'," & STATUS & "," & BIRTHPLACE & _
                      "," & HEIGHT & "," & WEIGHT & "," & RELIGION & "," & CITIZEN & ",'" & MYPOSITION & _
                      "','" & EMPSTATUS & "','" & Date & "'," & SPOUSE & "," & SPOUSEAGE & "," & SOCCUPATION & _
                      "," & FATHER & "," & FATHERAGE & "," & FOCCUPATION & "," & MOTHER & "," & MOTHERAGE & "," & MOCCUPATION & ",'" & "A" & "'," & IMAGE_LOCATION & ")")


    Set rsAPP = Nothing
    Set RSTMP = Nothing

    Dim vTRAINING                                                     As String
    Dim vMONYEAR                                                      As String
    Dim vPLACE                                                        As String
    Dim vSPONSOR                                                      As String

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_Applicant_Train Where Applicant_ID = " & lblAPP(0).Caption & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            vTRAINING = Null2String(RSTMP!Training)
            vMONYEAR = Null2String(RSTMP!monthYear)
            vPLACE = Null2String(RSTMP!Place)
            vSPONSOR = Null2String(RSTMP!Sponsor)

            gconDMIS.Execute ("Insert Into HRMS_Training (EMPLEVEL,EMPNO,TRAINING,MONYEAR,PLACE,SPONSOR,UESRCODE,ID) Values ('" & "E" & _
                              "'," & EMPNO & "," & vTRAINING & _
                              "," & vMONYEAR & "," & vPLACE & _
                              "," & vSPONSOR & ",'" & LOGCODE & "'," & RSTMP!Enr_ID & ")")

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3:
            TXTSEARCH.Text = ""
            On Error Resume Next
            TXTSEARCH.Text = ""
            TXTSEARCH.SetFocus
    End Select
End Sub

Private Sub Form_Load()

    Call CenterMe(frmMain, Me, 1)

    frmMain.MousePointer = 11

    Call FillCBOType
    Call FillCboPosition
    Call FillCboDepartment

    tcPOSITION.SelectedItem = 0
    tcAPP.SelectedItem = 0
    cboSAL.ListIndex = 0

    frmMain.MousePointer = 0
    On Error Resume Next
    cboPosition.SetFocus
End Sub

Private Sub LsvAPP_DblClick()
    Dim Index                                                         As Long
    Dim rsPER As ADODB.Recordset, rsDOC                               As ADODB.Recordset
    Dim rsPIC                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    LsvAPP.Enabled = False

    If Not LsvAPP.ListItems.count = 0 Then
        Index = LsvAPP.SelectedItem.Index
        With LsvAPP
            lblAPP(0).Caption = .ListItems(Index).SubItems(1)
            Set rsPER = gconDMIS.Execute("Select * From HRMS_APPLICANT_PERSONAL Where Applicant_ID = " & .ListItems(Index).SubItems(1) & "")
            If Not (rsPER.BOF And rsPER.EOF) Then
                lblAPP(1).Caption = rsPER!lastname
                lblAPP(2).Caption = rsPER!FIRSTNAME
                lblAPP(3).Caption = rsPER!AGE
                lblAPP(4).Caption = Null2String(rsPER!HighestLevel1)
                lblAPP(5).Caption = Null2String(rsPER!StudyFields1)
                lblAPP(6).Caption = Null2String(rsPER!Major1)
                lblAPP(7).Caption = Null2String(rsPER!HighestLevel2)
                lblAPP(8).Caption = Null2String(rsPER!StudyFields2)
                lblAPP(9).Caption = Null2String(rsPER!Major2)
                lblAPP(10).Caption = Null2String(rsPER!PositionDesired)
            End If
            Set rsPER = Nothing

            Set rsDOC = gconDMIS.Execute("Select * From HRMS_APPLICANT_PAPER Where Applicant_ID = " & .ListItems(Index).SubItems(1) & "")
            lsvDOC.ListItems.Clear
            If Not (rsDOC.BOF And rsDOC.EOF) Then
                Do While Not rsDOC.EOF
                    Set ITEM = lsvDOC.ListItems.Add(, , Null2String(rsDOC!PaperPass))

                    rsDOC.MoveNext
                Loop
            End If
            Set rsDOC = Nothing

            Set rsPIC = gconDMIS.Execute("Select * From HRMS_APPLICANT_IMAGE_LOCATION Where Applicant_ID = " & .ListItems(Index).SubItems(1) & "")
            If Not (rsPIC.BOF And rsPIC.EOF) Then
                If Null2String(rsPIC!ImageLocation) <> "" Then
                    On Error Resume Next
                    LoadPic imgAPP, Null2String(rsPIC!ImageLocation)
                Else
                    LoadPic imgAPP, ""
                End If
            Else
                LoadPic imgAPP, ""
            End If
            Set rsPIC = Nothing

            On Error Resume Next
            cboPosition.SetFocus
        End With
    End If
    LsvAPP.Enabled = True
End Sub

Private Sub LsvAPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then LsvAPP_DblClick
End Sub

Private Sub txtSEARCH_Change()
    Dim RSTMP As ADODB.Recordset, rsINT                               As ADODB.Recordset
    Dim Keyword                                                       As String
    Dim ITEM                                                          As ListItem

    On Error GoTo Errorcode:

    LsvAPP.Enabled = False
    Keyword = Trim(TXTSEARCH.Text)

    Set RSTMP = gconDMIS.Execute("Select LastName,FirstName,Applicant_ID From HRMS_APPLICANT_PERSONAL Where (LastName Like '%" & Keyword & "%' Or FirstName Like '%" & Keyword & _
                                 "%') And Hired = '" & "NO" & "' Order by Applicant_ID ASC")
    LsvAPP.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = LsvAPP.ListItems.Add(, , RSTMP!lastname & ", " & RSTMP!FIRSTNAME)
            ITEM.SubItems(1) = RSTMP!APPLICANT_ID

            RSTMP.MoveNext
        Loop
    Else
        LsvAPP.ListItems.Clear
    End If

    If TXTSEARCH.Text = "" Then LsvAPP.ListItems.Clear
    LsvAPP.Enabled = True
    Exit Sub
Errorcode:
    ShowVBError
End Sub


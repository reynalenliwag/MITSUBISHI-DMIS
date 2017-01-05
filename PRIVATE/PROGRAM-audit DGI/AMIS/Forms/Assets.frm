VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISDATAAssets 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assets Registry"
   ClientHeight    =   5190
   ClientLeft      =   1260
   ClientTop       =   435
   ClientWidth     =   10980
   ForeColor       =   &H00FFC0C0&
   Icon            =   "Assets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   10980
   Visible         =   0   'False
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H00FF8080&
      Height          =   1785
      Left            =   30
      ScaleHeight     =   1725
      ScaleWidth      =   3015
      TabIndex        =   45
      Top             =   3390
      Width           =   3075
      Begin wizButton.cmd cmd1 
         Height          =   465
         Left            =   2070
         TabIndex        =   51
         Top             =   1170
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   820
         TX              =   "Ok"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Assets.frx":08CA
      End
      Begin VB.CheckBox chk1 
         BackColor       =   &H00FF8080&
         Caption         =   "All Assets"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   1
         Left            =   150
         TabIndex        =   49
         Top             =   690
         Width           =   2745
      End
      Begin VB.CheckBox chk1 
         BackColor       =   &H00FF8080&
         Caption         =   "Per Assets Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   150
         TabIndex        =   48
         Top             =   360
         Width           =   2745
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H000080FF&
         Height          =   345
         Left            =   -60
         ScaleHeight     =   285
         ScaleWidth      =   4365
         TabIndex        =   46
         Top             =   0
         Width           =   4425
         Begin VB.CommandButton Command1 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2730
            TabIndex        =   50
            Top             =   0
            Width           =   315
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Printing Option"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   47
            Top             =   30
            Width           =   2175
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5340
      ScaleHeight     =   855
      ScaleWidth      =   5940
      TabIndex        =   36
      ToolTipText     =   "Save this Record"
      Top             =   4320
      Width           =   5940
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
         Left            =   4860
         MouseIcon       =   "Assets.frx":08E6
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":0A38
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
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
         Left            =   4170
         MouseIcon       =   "Assets.frx":0D9E
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Print this Record"
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
         Left            =   3480
         MouseIcon       =   "Assets.frx":1256
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":13A8
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Left            =   2790
         MouseIcon       =   "Assets.frx":16D3
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":1825
         Style           =   1  'Graphical
         TabIndex        =   42
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
         Left            =   2100
         MouseIcon       =   "Assets.frx":1B81
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":1CD3
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
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
         Left            =   1410
         MouseIcon       =   "Assets.frx":1FE6
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":2138
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
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
         Left            =   720
         MouseIcon       =   "Assets.frx":2432
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":2584
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
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
         Left            =   30
         MouseIcon       =   "Assets.frx":28DC
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":2A2E
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3045
      Left            =   150
      ScaleHeight     =   3045
      ScaleWidth      =   10500
      TabIndex        =   15
      Top             =   1140
      Width           =   10500
      Begin VB.TextBox txtVoucher_No 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   5430
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   60
         Width           =   1005
      End
      Begin VB.ComboBox cboSupCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   360
         Left            =   1560
         TabIndex        =   3
         Text            =   "cboSupCode"
         Top             =   1290
         Width           =   4875
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2460
         Width           =   4875
      End
      Begin VB.TextBox txtSerialNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1680
         Width           =   4875
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   780
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "Assets.frx":2D8D
         Top             =   480
         Width           =   4875
      End
      Begin VB.TextBox txtAssetCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   60
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   2955
         Left            =   6525
         TabIndex        =   26
         Top             =   0
         Width           =   3885
         Begin MSMask.MaskEdBox txtAccumDep 
            Height          =   360
            Left            =   2070
            TabIndex        =   12
            Top             =   1920
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   8421504
            ForeColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtBookBalance 
            Height          =   360
            Left            =   2070
            TabIndex        =   11
            Top             =   1500
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   8421504
            ForeColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtAccCost 
            Height          =   360
            Left            =   2070
            TabIndex        =   9
            Top             =   690
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtEsti_Life 
            Height          =   360
            Left            =   2070
            TabIndex        =   8
            Top             =   270
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMonthly_Dep 
            Height          =   360
            Left            =   2070
            TabIndex        =   13
            Top             =   2340
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   8421504
            ForeColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtLife_Rem 
            Height          =   360
            Left            =   2070
            TabIndex        =   10
            Top             =   1080
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Months"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2730
            TabIndex        =   53
            Top             =   1140
            Width           =   915
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Life Remaining"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -210
            TabIndex        =   52
            Top             =   1140
            Width           =   2055
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Book Balance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   180
            TabIndex        =   28
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Accumulated Dep."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1980
            Width           =   1815
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Acquisition Cost"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   27
            Top             =   750
            Width           =   1695
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated Life"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -210
            TabIndex        =   30
            Top             =   330
            Width           =   2055
         End
         Begin VB.Label Label11 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Months"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2730
            TabIndex        =   31
            Top             =   330
            Width           =   915
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Dep."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -120
            TabIndex        =   32
            Top             =   2370
            Width           =   2055
         End
      End
      Begin MSMask.MaskEdBox txtSal_Value 
         Height          =   360
         Left            =   4830
         TabIndex        =   6
         Top             =   2070
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   7347754
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpDate_Purch 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   2100
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131858433
         CurrentDate     =   38216
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3030
         TabIndex        =   20
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         TabIndex        =   21
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3900
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Acquired"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -150
         TabIndex        =   23
         Top             =   2130
         Width           =   1635
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   22
         Top             =   1710
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Asset Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salvage Value"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   2130
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9465
      ScaleHeight     =   885
      ScaleWidth      =   1620
      TabIndex        =   33
      Top             =   4320
      Width           =   1620
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
         MouseIcon       =   "Assets.frx":2D93
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":2EE5
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Cancel Entry"
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
         MouseIcon       =   "Assets.frx":3223
         MousePointer    =   99  'Custom
         Picture         =   "Assets.frx":3375
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4230
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7461
      _Version        =   393216
      Style           =   1
      Tabs            =   12
      Tab             =   6
      TabsPerRow      =   6
      TabHeight       =   882
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Land"
      TabPicture(0)   =   "Assets.frx":36C5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Building"
      TabPicture(1)   =   "Assets.frx":36E1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Leasehold and Improvements"
      TabPicture(2)   =   "Assets.frx":36FD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Service Tools and Equipment"
      TabPicture(3)   =   "Assets.frx":3719
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Company Vehicle"
      TabPicture(4)   =   "Assets.frx":3735
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Furnitures and Fixtures"
      TabPicture(5)   =   "Assets.frx":3751
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Office Equipment"
      TabPicture(6)   =   "Assets.frx":376D
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Computing Hardware"
      TabPicture(7)   =   "Assets.frx":3789
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Computing Software"
      TabPicture(8)   =   "Assets.frx":37A5
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Airconditioning Equipment"
      TabPicture(9)   =   "Assets.frx":37C1
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Communication Equipment"
      TabPicture(10)  =   "Assets.frx":37DD
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Other Equipment"
      TabPicture(11)  =   "Assets.frx":37F9
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
   End
   Begin Crystal.CrystalReport rptAssets 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "List of Assets"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmAMISDATAAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAssets                                                As ADODB.Recordset
Dim rsChartAccount                                          As ADODB.Recordset
Dim rsVENDOR                                                As ADODB.Recordset
Dim AddorEdit                                               As String
Dim PPE_TYPE                                                As String

Function SetAccCode(Acc As String) As String
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "select * from AMIS_ChartAccount where description = " & N2Str2Null(Acc), gconDMIS
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccCode = Null2String(rsChartAccount!AcctCode)
    Else
        SetAccCode = ""
    End If
End Function

Function SetAccType(Acc As String) As String
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "select * from AMIS_ChartAccount where Acctcode = " & N2Str2Null(Acc), gconDMIS
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccType = Null2String(rsChartAccount!DESCRIPTION)
    Else
        SetAccType = "Not Defined"
    End If
End Function

Function SetVendorCode(Acc As String) As String
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "select * from ALL_Vendor where nameofvendor = " & N2Str2Null(Acc), gconDMIS
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorCode = Null2String(rsVENDOR!CODE)
    Else
        SetVendorCode = ""
    End If
End Function

Function SetVendorName(Acc As String) As String
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "select * from ALL_Vendor where code = " & N2Str2Null(Acc), gconDMIS
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = "Not Defined"
    End If
End Function

Sub initMemvars()
    Frame1.Enabled = True
    txtAssetCode.Text = ""
    txtVoucher_No.Text = ""
    txtDescription.Text = ""
    txtSerialNo.Text = ""
    dtpDate_Purch = LOGDATE
    txtSal_Value.Text = 0#
    txtLocation.Text = ""

    txtAccCost.Text = 0#
    txtEsti_Life.Text = 0#
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "select * from ALL_Vendor order by code asc", gconDMIS
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        rsVENDOR.MoveFirst
        cboSupCode.Clear: cboSupCode.AddItem "Not Defined"
        Do While Not rsVENDOR.EOF
            cboSupCode.AddItem Null2String(rsVENDOR!nameofvendor)
            rsVENDOR.MoveNext
        Loop
    End If
End Sub

Sub PPE_Refresh(XXX As String)
    Dim rsPPE_Refresh                                       As ADODB.Recordset
    Set rsPPE_Refresh = New ADODB.Recordset
    rsPPE_Refresh.Open "select * from AMIS_Assets where assetname = '" & UCase(Trim(XXX)) & "' order by date_purch asc,assetcode asc", gconDMIS, adOpenKeyset
    Set rsAssets = New ADODB.Recordset
    Set rsAssets = rsPPE_Refresh.Clone
    If Not rsPPE_Refresh.EOF And Not rsPPE_Refresh.BOF Then
        rsAssets.Bookmark = rsPPE_Refresh.Bookmark
    End If
    StoreMemVars
    Set rsPPE_Refresh = Nothing
End Sub

Sub rsRefresh()
    Set rsAssets = New ADODB.Recordset
    rsAssets.Open "select * from AMIS_Assets where assetname = '" & PPE_TYPE & "' order by date_purch asc,AssetCode asc", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemVars()
    If Not rsAssets.EOF And Not rsAssets.BOF Then
        Frame1.Enabled = False
        labID.Caption = rsAssets!ID
        txtVoucher_No.Text = Null2String(rsAssets!voucher_no)
        txtAssetCode.Text = Null2String(rsAssets!assetCode)
        'txtAssetName.Text = Null2String(rsAssets!AssetName)
        txtDescription.Text = Null2String(rsAssets!DESCRIPTION)
        cboSupCode.Text = SetVendorName(Null2String(rsAssets!supcode))
        txtSerialNo.Text = Null2String(rsAssets!serialno)
        txtAccCost.Text = N2Str2Zero(rsAssets!acq_cost)
        txtSal_Value.Text = N2Str2Zero(rsAssets!sal_value)
        If Null2Date(rsAssets!date_purch) = "" Then
            dtpDate_Purch = LOGDATE
        Else
            dtpDate_Purch = Null2Date(rsAssets!date_purch)
        End If
        txtEsti_Life.Text = N2Str2Zero(rsAssets!esti_life)
        txtLife_Rem.Text = N2Str2Zero(rsAssets!life_rem)
        txtLocation.Text = N2Str2Zero(rsAssets!Location)
        cmdPrevious.Enabled = True: cmdNext.Enabled = True
        cmdEdit.Enabled = True: cmdDelete.Enabled = True: cmdFind.Enabled = True
    Else
        'MsgBox "No Such Record!"
        'cmdAdd.Value = True
        initMemvars
        cmdPrevious.Enabled = False: cmdNext.Enabled = False
        cmdEdit.Enabled = False: cmdDelete.Enabled = False: cmdFind.Enabled = False
    End If
End Sub

Sub Compute()
    If NumericVal(txtLife_Rem.Text) > 0 Then
        txtBookBalance.Text = Round((NumericVal(txtAccCost) / NumericVal(txtEsti_Life.Text)) * NumericVal(txtLife_Rem.Text), 2)
    Else
        txtBookBalance.Text = 0
    End If
    txtAccumDep.Text = Round(NumericVal(txtAccCost.Text) - NumericVal(txtBookBalance.Text), 2)
    If NumericVal(txtEsti_Life.Text) > 0 Then
        txtMonthly_Dep.Text = Round(NumericVal(txtAccCost.Text) / NumericVal(txtEsti_Life.Text), 2)
    Else
        txtMonthly_Dep.Text = 0
    End If
End Sub

Private Sub cmd1_Click()
'Update By BTT - 07222008
' Per Asset
    If chk1(0).Value = 1 Then
        On Error GoTo ErrorCode:
        If Function_Access(LOGID, "Acess_Print", "ASSETS REGISTRY") = False Then Exit Sub
        Screen.MousePointer = 11
        Dim rsProfile                                       As ADODB.Recordset
        rptAssets.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAssets.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAssets.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAssets.ReportTitle = "LIST OF ASSETS"
        End If
        PrintReport rptAssets, AMIS_REPORT_PATH & "\Files\Assets.rpt", "{Assets.AssetName} = '" & UCase(SSTab1.Caption) & "'", 1
        'PrintReport rptAssets, AMIS_REPORT_PATH & "\Files\Assets.rpt", "", 1
        LogAudit "V", "ASSETS REGISTRY", "Asset Code: " & txtAssetCode
        Screen.MousePointer = 0
    End If
    'All Assets
    If chk1(1).Value = 1 Then
        On Error GoTo ErrorCode:
        If Function_Access(LOGID, "Acess_Print", "ASSETS REGISTRY") = False Then Exit Sub
        Screen.MousePointer = 11
        Dim rsProfile1                                      As ADODB.Recordset
        rptAssets.Reset
        Set rsProfile1 = New ADODB.Recordset
        Set rsProfile1 = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile1.EOF And rsProfile1.BOF) Then
            rptAssets.Formulas(3) = "CompanyName = '" & Null2String(rsProfile1!CompanyName) & "'"
            rptAssets.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile1!Companyaddress) & "'"
            rptAssets.ReportTitle = "LIST OF ASSETS"
        End If
        'PrintReport rptAssets, AMIS_REPORT_PATH & "\Files\ListOfAssets.rpt", "{Assets.AssetName} = '" & UCase(SSTab1.Caption) & "'", 1
        PrintReport rptAssets, AMIS_REPORT_PATH & "\Files\Assets.rpt", "", 1
        LogAudit "V", "ASSETS REGISTRY", "Asset Code: " & txtAssetCode
        Screen.MousePointer = 0
    End If

ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:01
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "ASSETS REGISTRY") = False Then Exit Sub

    AddorEdit = "ADD"
    initMemvars
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtAssetCode.SetFocus
    SSTab1.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    SSTab1.Enabled = True
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200713:02
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "ASSETS REGISTRY") = False Then Exit Sub
    If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
        gconDMIS.Execute "delete from AMIS_Assets where id = " & labID.Caption
    End If
    Call PPE_Refresh(Trim(SSTab1.Caption))
    StoreMemVars
    LogAudit "X", "ASSETS REGISTRY", "Asset Code: " & txtAssetCode
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:02
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "ASSETS REGISTRY") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtVoucher_No.SetFocus
    SSTab1.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                                             As String
    findStr = InputBox("Please Input Assets ...", "Find")
    If findStr <> "" Then
        On Error Resume Next
        rsAssets.Bookmark = rsFind(rsAssets.Clone, "AssetCode", findStr).Bookmark
        If Err.Number = 3021 Then
            On Error Resume Next
            rsAssets.Bookmark = rsFind(rsAssets.Clone, "AssetCode", findStr).Bookmark
            If Err.Number = 3021 Then
                On Error GoTo ErrorCode
                rsAssets.Bookmark = rsFind(rsAssets.Clone, "Description", findStr).Bookmark
            End If
        End If
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.Number = 3021 Then
        'MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
        MessagePop RecNotFound, "Not Found", "Can't find " & findStr
        Resume Next
    End If
End Sub

Private Sub cmdNext_Click()
    rsAssets.MoveNext
    If rsAssets.EOF Then
        rsAssets.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsAssets.MovePrevious
    If rsAssets.BOF Then
        rsAssets.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200713:02
Private Sub cmdPrint_Click()
    picPrinting.Visible = True
End Sub

'Upating Code       : AXP-0707200713:02
Private Sub cmdSave_Click()

    Dim VtxtVoucher_No                                      As String
    Dim VtxtAssetCode, VtxtAssetName, vtxtDescription       As String
    Dim VcboSupCode, VtxtSerialNo                           As String
    Dim VtxtAcq_Cost, VtxtSal_Value                         As Double
    Dim Vdtpdate_purch                                      As String
    Dim VtxtEsti_Life                                       As Double
    Dim VtxtLocation                                        As String

    On Error GoTo ErrorCode:

    If txtAssetCode.Text = "" Then
        MsgBox "Invalid Asset Code", vbInformation, "Not Allowed"
        Exit Sub
    End If
    VtxtVoucher_No = N2Str2Null(txtVoucher_No.Text)
    VtxtAssetCode = N2Str2Null(txtAssetCode.Text)
    VtxtAssetName = N2Str2Null(UCase(SSTab1.Caption))
    vtxtDescription = N2Str2Null(txtDescription.Text)
    VcboSupCode = N2Str2Null(SetVendorCode(cboSupCode.Text))
    VtxtSerialNo = N2Str2Null(txtSerialNo.Text)
    VtxtAcq_Cost = NumericVal(txtAccCost.Text)
    VtxtSal_Value = NumericVal(txtSal_Value.Text)
    Vdtpdate_purch = N2Date2Null(dtpDate_Purch)
    VtxtEsti_Life = NumericVal(txtEsti_Life.Text)
    VtxtLocation = N2Str2Null(txtLocation.Text)

    If AddorEdit = "ADD" Then
        Dim rsAssetDup                                      As ADODB.Recordset
        Set rsAssetDup = New ADODB.Recordset
        rsAssetDup.Open "select Assetcode from AMIS_Assets where Assetcode = " & VtxtAssetCode, gconDMIS
        If Not rsAssetDup.EOF And Not rsAssetDup.BOF Then
            'MsgBox "Assets Code Already Exist!", vbCritical, "Duplicate Code Not Allowed"
            MessagePop RecSaveError, "Duplicate Entry", "Assets Code Already Exist!"
            Exit Sub
        End If

        SQL_STATEMENT = "Insert into AMIS_Assets " & _
                        "(Voucher_no,AssetCode,AssetName,Description,SupCode,SerialNo,Acq_cost,Sal_Value,Date_Purch,Esti_Life,Life_Rem,Location) " & _
                        " values (" & VtxtVoucher_No & ", " & VtxtAssetCode & ", " & VtxtAssetName & _
                        ", " & vtxtDescription & ", " & VcboSupCode & ", " & VtxtSerialNo & ", " & VtxtAcq_Cost & ", " & VtxtSal_Value & _
                        ", " & Vdtpdate_purch & ", " & VtxtEsti_Life & "," & NumericVal(txtLife_Rem.Text) & ", " & VtxtLocation & ")"
        gconDMIS.Execute SQL_STATEMENT

        NEW_LogAudit "A", "ASSETS REGISTRY", SQL_STATEMENT, labID.Caption, "", VtxtVoucher_No, "", N2Str2Null(VtxtAssetCode)
    Else
        SQL_STATEMENT = "update AMIS_Assets set" & _
                        " Voucher_No = " & VtxtVoucher_No & "," & _
                        " AssetCode = " & VtxtAssetCode & "," & _
                        " assetname = " & VtxtAssetName & "," & _
                        " description = " & vtxtDescription & "," & _
                        " supcode = " & VcboSupCode & "," & _
                        " serialno = " & VtxtSerialNo & "," & _
                        " acq_cost = " & VtxtAcq_Cost & "," & _
                        " sal_value = " & VtxtSal_Value & "," & _
                        " date_purch = " & Vdtpdate_purch & "," & _
                        " esti_life = " & VtxtEsti_Life & "," & _
                        " Life_Rem = " & NumericVal(txtLife_Rem.Text) & "," & _
                        " location = " & VtxtLocation & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT

        NEW_LogAudit "E", "ASSETS REGISTRY", SQL_STATEMENT, labID.Caption, "", VtxtVoucher_No, "", N2Str2Null(VtxtAssetCode)
    End If
    Call PPE_Refresh(Trim(SSTab1.Caption))
    On Error Resume Next
    rsAssets.Find "assetcode = " & VtxtAssetCode
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Command1_Click()
    picPrinting.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = Me.Caption
        Call frmALL_AuditInquiry.DisplayHistory(labID, "ASSETS REGISTRY")
    Case vbKeyEscape
        'fraDetails.ZOrder 1
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    picPrinting.Visible = False: picPrinting.Left = 7830: picPrinting.Top = 2490
    initMemvars
    Call PPE_Refresh(Trim(SSTab1.Caption))
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call PPE_Refresh(Trim(SSTab1.Caption))
End Sub

Private Sub txtEsti_Life_LostFocus()
    Compute
End Sub

Private Sub txtLife_Rem_Change()
    Compute
End Sub

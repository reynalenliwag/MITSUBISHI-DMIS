VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMSOvertime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overtime Entry"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9705
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Overtime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   9705
   Begin Crystal.CrystalReport rptOvertime 
      Left            =   6960
      Top             =   2010
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Overtime"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   5595
      Left            =   60
      Picture         =   "Overtime.frx":030A
      ScaleHeight     =   5595
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   30
      Width           =   2475
      Begin VB.TextBox txtSearch 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   0
         MaxLength       =   35
         TabIndex        =   1
         Top             =   0
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstOverTime 
         Height          =   5175
         Left            =   0
         TabIndex        =   2
         Top             =   420
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Overtime.frx":14067
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox picOvertime 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   3150
      ScaleHeight     =   4395
      ScaleWidth      =   4875
      TabIndex        =   13
      Top             =   300
      Width           =   4905
      Begin VB.ComboBox cboQuensina 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "Overtime.frx":141C9
         Left            =   90
         List            =   "Overtime.frx":141CB
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Text            =   "cboQuensina"
         Top             =   1290
         Width           =   3195
      End
      Begin VB.TextBox txtNoHours_Computed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1110
         TabIndex        =   22
         Top             =   2490
         Width           =   1125
      End
      Begin VB.TextBox txtNoMin_Computed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   23
         Top             =   2490
         Width           =   1125
      End
      Begin VB.TextBox txtJustif 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   150
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   3270
         Width           =   4665
      End
      Begin VB.ComboBox cboOT 
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   630
         Width           =   4335
      End
      Begin VB.ComboBox cboYear2 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3750
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox cboDay2 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3840
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   315
         Left            =   1920
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtTotalHr 
         Height          =   315
         Left            =   1920
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   6600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtHoliday 
         Height          =   315
         Left            =   3300
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3570
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
      Begin VB.ComboBox cboMonth2 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3840
         Visible         =   0   'False
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dt_ot 
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Top             =   1950
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54788097
         CurrentDate     =   40179
      End
      Begin VB.Label Label17 
         Caption         =   "Cut-Off"
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
         Left            =   120
         TabIndex        =   69
         Top             =   1050
         Width           =   915
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   65
         Top             =   -30
         Width           =   5475
         _Version        =   655364
         _ExtentX        =   9657
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "  Edit Over Time"
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
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   90
         X2              =   4710
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes"
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
         Height          =   285
         Left            =   2400
         TabIndex        =   42
         Top             =   2550
         Width           =   645
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Hours"
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
         Height          =   285
         Left            =   -150
         TabIndex        =   37
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Justification"
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   3030
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Computed Hour(s)"
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
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   47
         Top             =   6630
         Width           =   1590
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   4740
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Label lblHR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
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
         Height          =   195
         Left            =   2940
         TabIndex        =   16
         Top             =   420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday Amount"
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
         Height          =   195
         Left            =   2460
         TabIndex        =   43
         Top             =   3600
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Select OT Description"
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
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   420
         Width           =   1845
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "OT Rate"
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
         Height          =   195
         Index           =   3
         Left            =   3000
         TabIndex        =   17
         Top             =   5550
         Width           =   690
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ADD"
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
         Height          =   195
         Index           =   0
         Left            =   5910
         TabIndex        =   41
         Top             =   2100
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Hour"
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
         Height          =   195
         Index           =   2
         Left            =   5430
         TabIndex        =   46
         Top             =   2430
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblRPH 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6450
         TabIndex        =   45
         Top             =   2340
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblOTR 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6450
         TabIndex        =   40
         Top             =   2040
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Rate"
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
         Height          =   195
         Index           =   1
         Left            =   5430
         TabIndex        =   36
         Top             =   1770
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblEmpRate 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6450
         TabIndex        =   35
         Top             =   1680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblRate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3030
         TabIndex        =   18
         Top             =   5790
         Width           =   765
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1650
         Width           =   525
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime Amount"
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
         Height          =   195
         Left            =   300
         TabIndex        =   38
         Top             =   6300
         Width           =   1515
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Height          =   285
         Left            =   3960
         TabIndex        =   34
         Top             =   3870
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Height          =   285
         Left            =   240
         TabIndex        =   25
         Top             =   5580
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Height          =   285
         Left            =   1260
         TabIndex        =   26
         Top             =   5550
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Height          =   285
         Left            =   2250
         TabIndex        =   27
         Top             =   5550
         Width           =   465
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Left            =   240
         TabIndex        =   28
         Top             =   3720
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Height          =   285
         Left            =   390
         TabIndex        =   32
         Top             =   3960
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin wizButton.cmd cmdOvertime 
      Height          =   4515
      Left            =   3120
      TabIndex        =   50
      Top             =   270
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   7964
      TX              =   "cmd1"
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
      MICON           =   "Overtime.frx":141CD
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   3345
      Left            =   2610
      ScaleHeight     =   3345
      ScaleWidth      =   7095
      TabIndex        =   51
      Top             =   1350
      Width           =   7095
      Begin MSFlexGridLib.MSFlexGrid grdOvertime 
         Height          =   2985
         Left            =   0
         TabIndex        =   52
         Top             =   270
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5265
         _Version        =   393216
         Cols            =   7
         FixedCols       =   2
         ForeColor       =   0
         ForeColorFixed  =   0
         BackColorSel    =   -2147483633
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1335
      Left            =   2610
      ScaleHeight     =   1335
      ScaleWidth      =   7095
      TabIndex        =   4
      Top             =   90
      Width           =   7095
      Begin VB.TextBox txtTOT 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5460
         TabIndex        =   12
         Top             =   900
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtPosition 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   900
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   5985
      End
      Begin VB.TextBox txtName 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   900
         TabIndex        =   3
         Top             =   60
         Width           =   5985
      End
      Begin VB.TextBox txtYTDOvertime 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1500
         TabIndex        =   10
         Top             =   900
         Visible         =   0   'False
         Width           =   1695
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   0
         TabIndex        =   66
         Top             =   480
         Width           =   6915
         _Version        =   655364
         _ExtentX        =   12197
         _ExtentY        =   661
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total no. Of Overtime"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   3390
         TabIndex        =   11
         Top             =   990
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   5
         Top             =   150
         Width           =   540
      End
      Begin VB.Label labID 
         Caption         =   "Label6"
         Height          =   285
         Left            =   1530
         TabIndex        =   8
         Top             =   510
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   510
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Overtime"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   1020
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4110
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   56
      Top             =   4740
      Width           =   5580
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
         MouseIcon       =   "Overtime.frx":141E9
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":1433B
         Style           =   1  'Graphical
         TabIndex        =   64
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
         MouseIcon       =   "Overtime.frx":146A1
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":147F3
         Style           =   1  'Graphical
         TabIndex        =   63
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
         MouseIcon       =   "Overtime.frx":14B59
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":14CAB
         Style           =   1  'Graphical
         TabIndex        =   62
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
         MouseIcon       =   "Overtime.frx":14FD6
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":15128
         Style           =   1  'Graphical
         TabIndex        =   61
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
         MouseIcon       =   "Overtime.frx":15484
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":155D6
         Style           =   1  'Graphical
         TabIndex        =   60
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
         MouseIcon       =   "Overtime.frx":158E9
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":15A3B
         Style           =   1  'Graphical
         TabIndex        =   59
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
         MouseIcon       =   "Overtime.frx":15D35
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":15E87
         Style           =   1  'Graphical
         TabIndex        =   58
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
         MouseIcon       =   "Overtime.frx":161DF
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":16331
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   8235
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   53
      Top             =   4725
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
         MouseIcon       =   "Overtime.frx":16690
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":167E2
         Style           =   1  'Graphical
         TabIndex        =   55
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
         MouseIcon       =   "Overtime.frx":16B20
         MousePointer    =   99  'Custom
         Picture         =   "Overtime.frx":16C72
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      Index           =   2
      Left            =   2790
      TabIndex        =   68
      Top             =   1560
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label LABMONTH 
      BackColor       =   &H000000FF&
      Height          =   345
      Left            =   2310
      TabIndex        =   67
      Top             =   6570
      Width           =   3285
   End
End
Attribute VB_Name = "frmHRMSOvertime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo                                                           As ADODB.Recordset
Dim rsOvertime                                                          As ADODB.Recordset
Dim rsSalaryGrade                                                       As ADODB.Recordset
Dim rsSETUPDEDUCTION                                                    As ADODB.Recordset
Dim DAYS_OF_WORK                                                        As Integer
Dim ADDOREDIT, Diyt, Diyt2                                              As String
Dim Obertaym, Halidey                                                   As Double
Attribute Obertaym.VB_VarUserMemId = 1073938438
Dim EMPLIVIL                                                            As String
Attribute EMPLIVIL.VB_VarUserMemId = 1073938440
Dim LOC_NAME                                                            As String
Dim LOC_LEVEL                                                           As String
Public Event SelectionMade()

Public Sub SetVariable(xNAME As String, XLEVEL As String)
    LOC_NAME = xNAME
    LOC_LEVEL = XLEVEL
End Sub

Function StoreEntry(ByVal ID As Variant)
    Dim MM                                                            As String
    Dim DD                                                            As String
    Dim YY                                                            As String
    Dim MM2                                                           As String
    Dim DD2                                                           As String
    Dim YY2                                                           As String
    Dim TheDeyt                                                       As String
    Dim TheDeyt2                                                      As String
    Dim totmin
    Set rsOvertime = New ADODB.Recordset
    rsOvertime.Open "SELECT * FROM HRMS_OVERTIME WHERE ID = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsOvertime.EOF And Not rsOvertime.BOF Then
        LABID.Caption = rsOvertime!ID
        TheDeyt = Null2Date(rsOvertime!DEYT)
        TheDeyt2 = Null2Date(rsOvertime!deyt2)
        If Null2String(rsOvertime!OCODE) = "R" Then
        End If
        If Null2String(rsOvertime!OCODE) = "RH" Then
        End If
        If Null2String(rsOvertime!OCODE) = "S" Then
        End If
        If Null2String(rsOvertime!OCODE) = "SH" Then
        End If
        If rsOvertime!CUT_OFF = 1 Then
            cboQuensina.Text = "1st Cut-Off"
        End If
        If rsOvertime!CUT_OFF = 2 Then
            cboQuensina.Text = "2nd Cut-Off"
        End If

'JBF: 08/31/2010 not applicable to datepicker
'        DD = Day(TheDeyt)
'        MM = The_month(MONTH(TheDeyt))
'        YY = YEAR(TheDeyt)
'        cboDay.Text = DD
'        cboMonth.Text = MM
'        cboYear.Text = YY
        
        
        dt_ot = TheDeyt
        txtTotalHr.Text = N2Str2Zero(rsOvertime!TotalHR)
        txtAmount.Text = N2Str2Zero(rsOvertime!AMOUNT)
        txtJustif.Text = Null2String(rsOvertime!Justification)
        totmin = NumericVal(rsOvertime!TotalHR) * 60
        txtNoMin_Computed = totmin Mod 60
        txtNoHours_Computed = totmin \ 60
        FindOTCodeThenPut rsOvertime!OCODE
    End If
End Function

Function FindOTCodeDescription(OTCODE As Integer)
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT PAY_DESC FROM HRMS_OTCODES WHERE PAY_CODE = " & OTCODE & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindOTCodeDescription = Null2String(RSTMP!PAY_DESC)
    Else
        FindOTCodeDescription = ""
    End If
    Set RSTMP = Nothing
End Function

Function GetBasicPay()
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT BASICSALARY, EMPSTATUS FROM HRMS_EMPINFO WHERE EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!EMPSTATUS) = "M" Then
            GetBasicPay = N2Str2Zero(RSTMP!BASICSALARY)
        ElseIf Null2String(RSTMP!EMPSTATUS) = "D" Then
            GetBasicPay = (N2Str2Zero(RSTMP!BASICSALARY) * DAYS_OF_WORK) / 12
        End If
    End If
    Set RSTMP = Nothing
End Function

Sub FindOTRate(HOLID_REGUL As String)
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim IDOT                                                          As String
    IDOT = Right(cboOT, 3)
    Set RSTMP = gconDMIS.Execute("SELECT ISHOLIDAY, PAY_RATE, PAY_CODE FROM HRMS_OTCODES WHERE PAY_CODE = '" & IDOT & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If IsNull(RSTMP!ISHOLIDAY) = False Then
            If RSTMP!ISHOLIDAY = True Then
                HOLID_REGUL = "HOLI"
            Else
                HOLID_REGUL = "REGU"
            End If
            lblHR.Caption = HOLID_REGUL
            lblRate.Caption = RSTMP!pay_rate
        Else
            HOLID_REGUL = "REGU"
            lblRate.Caption = RSTMP!pay_rate
        End If
    End If
    Set RSTMP = Nothing
End Sub

Sub EnabledPics(COND As Boolean)
    picSearch.Enabled = COND
    Picture4.Enabled = COND
End Sub

Sub FindOTCodeThenPut(vOCODE As String)
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim X                                                             As String
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_OTCODES WHERE PAY_CODE = '" & vOCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        cboOT.Text = Null2String(Trim(RSTMP!PAY_DESC)) & " - " & RSTMP!PAY_CODE
    End If
    Set RSTMP = Nothing
End Sub

Sub rsSETUP()
    Set rsSETUPDEDUCTION = gconDMIS.Execute("SELECT WORKING_DAY FROM HRMS_SETUPDEDUCTION")
    If Not (rsSETUPDEDUCTION.EOF And rsSETUPDEDUCTION.BOF) Then
        DAYS_OF_WORK = N2Str2Zero(rsSETUPDEDUCTION!WORKING_DAY)
    End If
    Set rsSETUPDEDUCTION = Nothing
End Sub

Sub FillCboOT()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim ZEROS                                                         As String
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_OTCODES ORDER BY PAY_DESC ASC")
    cboOT.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboOT.AddItem Null2String(Trim(RSTMP!PAY_DESC)) & " - " & ZEROS & RSTMP!PAY_CODE
            RSTMP.MoveNext
        Loop
        cboOT.ListIndex = 0
    End If
    Set RSTMP = Nothing
End Sub

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & EMPINFOEMPNO.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & frmHRMSEmpInfo.LABID.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set rsEmpInfo = New ADODB.Recordset
        
        'jbf: include confidentials employee in the list
        rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        'rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE RESIGNED IS NULL ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    
    End If
End Sub

Sub InitGrid()
    With grdOvertime
        .Rows = 2
        .ColWidth(0) = 1200
        .ColWidth(1) = 0
        .ColWidth(2) = 1300
        .ColWidth(3) = 1000
        .ColWidth(4) = 1100
        .ColWidth(5) = 0
        .ColWidth(6) = 1
        .Row = 0
        .Col = 0
        .Text = "From Date"
        .Col = 1
        .Text = "To Date"
        .Col = 2
        .Text = "Type"
        .Col = 3
        .Text = "Total Hours"
        .Col = 4
        .Text = "Amount"
        .Col = 5
        .Text = "Holiday"
        .Col = 6
        .Text = "ID"
    End With
End Sub

'Sub AddMonthName()
'    Dim X As Integer
'    cboMOnth.Clear
'    For X = 1 To 12
'        cboMOnth.AddItem MonthName(X)
'    Next
'End Sub

Sub InitMemvars()
    Dim rsCutoff                                                      As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM HRMS_PAYROLLSETUP")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            cboQuensina.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            cboQuensina.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        'COMMENT BY : MJP 11072008
            'cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        'COMMENT BY : MJP 11072008
        
        'UPDATE BY : MJP 11072008
            ' JBF: Call AddMonthName
            'JBF: cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
            LABMONTH.Caption = MonthName(Null2String(rsCutoff!PERIODMONTH))
        'UPDATE BY : MJP 11072008
        'JBF: cboYear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If
    
    
    'JBF ************************
'    fillcboDay cboDay
'    cboDay.Text = Day(Now)
'    fillcboDay cboDay2
'    cboDay.Text = Day(LOGDATE)
'    cboDay2.Text = Day(LOGDATE)
    'JBF ************************
    
    
    
    
    'txtTotalHr.Text = 0
    'txtAmount.Text = 0
    'txtNoMin_Computed = 0
    'txtNoHours_Computed = 0
End Sub

Sub StoreMemVars()
    'On Error GoTo Errorcode
    Dim CNT                                                           As Integer
    Dim Kode                                                          As String
    Dim VYTDOvertime                                                  As Double
    Dim vTOT                                                          As Double
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Set rsOvertime = New ADODB.Recordset
        rsOvertime.Open "SELECT * FROM HRMS_OVERTIME WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND CUT_OFF = '" & CUTTOFF_CODE & "' AND PAY_MONTH = '" & PAY_MONTH & "' AND PAY_YEAR = '" & PAY_YEAR & "' ORDER BY DEYT DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        CNT = 0
        VYTDOvertime = 0
        If Not rsOvertime.EOF And Not rsOvertime.BOF Then
            rsOvertime.MoveFirst
            cleargrid grdOvertime
            grdOvertime.Rows = grdOvertime.Rows
            Do While Not rsOvertime.EOF
                CNT = CNT + 1
                LABID.Caption = rsOvertime!ID
                Kode = FindOTCodeDescription(Null2String(rsOvertime!OCODE))
                grdOvertime.AddItem Null2Date(rsOvertime!DEYT) & _
                    Chr(9) & Null2Date(rsOvertime!deyt2) & _
                    Chr(9) & Kode & _
                    Chr(9) & N2Str2Zero(rsOvertime!TotalHR) & _
                    Chr(9) & N2Str2Zero(rsOvertime!AMOUNT) & _
                    Chr(9) & N2Str2Zero(rsOvertime!HOLIDAY) & _
                    Chr(9) & rsOvertime!ID
                    
                If YEAR(Null2Date(rsOvertime!DEYT)) = YEAR(LOGDATE) Then
                    VYTDOvertime = VYTDOvertime + N2Str2Zero(rsOvertime!AMOUNT) + N2Str2Zero(rsOvertime!HOLIDAY)
                    vTOT = vTOT + N2Str2Zero(rsOvertime!TotalHR)
                End If
                rsOvertime.MoveNext
            Loop
            grdOvertime.RemoveItem 1
        Else
            cleargrid grdOvertime
        End If
        txtYTDOvertime.Text = Format(N2Str2Zero(VYTDOvertime), MAXIMUM_DIGIT)
        txtTOT.Text = Format(N2Str2Zero(vTOT), MAXIMUM_DIGIT)
        txtPosition.Text = Null2String(rsEmpInfo!Position)
        txtName.Text = Cap1st(Null2String(rsEmpInfo!lastname)) & ", " & Cap1st(Null2String(rsEmpInfo!FIRSTNAME)) & " " & Cap1st(Null2String(rsEmpInfo!MIDDLENAME))
        
        'COMMENT BY  : MJP080109 1141AM
        'DESCRIPTION : TO NOT LOCK THE SEARCH PICTURE
            If EMPINFOSHOW = True Then
                picSearch.Enabled = False
            Else
                picSearch.Enabled = True
            End If
        'COMMENT BY  : MJP080109 1141AM
        
'Debug.Print EMPINFOSHOW
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub UpdateAmount_NEW(HOLID_REGUL As String)
    Dim TaymWork                                                      As Double
    Dim RperHar                                                       As Double
    Dim DailyRate                                                     As Double
    Obertaym = 0
    Halidey = 0
    TaymWork = NumericVal(txtTotalHr.Text)
    DailyRate = (GetBasicPay * 12) / DAYS_OF_WORK
    RperHar = DailyRate / 8
    If lblRate.Caption = "" Then
        lblRate.Caption = "0"
    End If
    lblOTR.Caption = ""
    lblEmpRate.Caption = DailyRate
    lblRPH.Caption = RperHar
    Obertaym = (TaymWork * RperHar) * CDbl(lblRate.Caption)
    lblOTR.Caption = ((TaymWork * RperHar) * CDbl(lblRate.Caption))
    If HOLID_REGUL = "HOLI" Then
        txtHoliday.Text = Format(Obertaym, MAXIMUM_DIGIT)
        txtAmount.Text = "0"
    End If
    If HOLID_REGUL = "REGU" Then
        txtAmount.Text = Format(Obertaym, MAXIMUM_DIGIT)
        txtHoliday.Text = "0"
    End If
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lstOverTime.Sorted = False
    lstOverTime.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME,EMPNO ,ID FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    
    'Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno ,ID from HRMS_EmpInfo WHERE RESIGNED IS NULL order by lastname+', '+firstname asc")
    
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lstOverTime.ListItems, rsEMPINFO2
        lstOverTime.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lstOverTime.Sorted = False
    lstOverTime.ListItems.Clear

    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO,ID  FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL AND LASTNAME+', '+FIRSTNAME LIKE'" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    'Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO,ID  FROM HRMS_EMPINFO WHERE RESIGNED IS NULL AND LASTNAME+', '+FIRSTNAME LIKE'" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lstOverTime.ListItems, rsEMPINFO2
        lstOverTime.Refresh
    End If
End Sub

Sub ComputeHours()
    'On Error Resume Next
    txtTotalHr = NumericVal(txtNoHours_Computed) + (NumericVal(txtNoMin_Computed)) / 60
End Sub

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Private Sub cboDay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub cboOT_Change()
    Dim HOLID_REGUL                                                   As String
    FindOTRate HOLID_REGUL
    UpdateAmount_NEW HOLID_REGUL
End Sub

Private Sub cboOT_Click()
    Dim HOLID_REGUL                                                   As String
    FindOTRate HOLID_REGUL
    UpdateAmount_NEW HOLID_REGUL
End Sub

Private Sub cboOT_LostFocus()
    Dim HOLID_REGUL                                                   As String
    FindOTRate HOLID_REGUL
    UpdateAmount_NEW HOLID_REGUL
End Sub

Private Sub chkSHoliday_Click()
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        'Call UpdateAmount
    End If
End Sub

Private Sub cmdAdd_Click()
    'On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Add", "EMPLOYEE MAINTAIN OVERTIME") = False Then Exit Sub
    
    ADDOREDIT = "ADD"
    picOvertime.Visible = True
    cmdOvertime.Visible = True
    picOvertime.Enabled = True
    cmdOvertime.ZOrder 0
    picOvertime.ZOrder 0
    Picture1.Visible = False
    Picture2.Visible = True
    txtJustif.Text = ""
    txtNoHours_Computed = ""
    txtNoMin_Computed = ""
    EnabledPics False
    InitMemvars
    'JBF
    dt_ot = LOGDATE
    '***
    Exit Sub
    
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    lstOverTime.Enabled = True
    ADDOREDIT = ""
    picOvertime.Visible = False
    cmdOvertime.Visible = False
    picOvertime.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    cmdOvertime.ZOrder 1
    picOvertime.ZOrder 1
    EnabledPics True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    'On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Delete", "EMPLOYEE MAINTAIN OVERTIME") = False Then Exit Sub
    grdOvertime.Col = 6
    If grdOvertime.Text <> "" Then
        If MsgBox("Delete this selected record, Are you sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        
        gconDMIS.Execute "DELETE FROM HRMS_OVERTIME WHERE ID = " & grdOvertime.Text
        LogAudit "X", "DELETE OVERTIME LIST OF THE EMPLOYEE", EMPLOYEE_NO
        ShowDeletedMsg
    Else
        ShowNothingToDeleteMsg
    End If
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    'On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE MAINTAIN OVERTIME") = False Then Exit Sub
    Dim fild                                                          As String
    grdOvertime.Row = grdOvertime.Row
    grdOvertime.Col = 6
    fild = grdOvertime.Text

    If fild <> "" Then
        lstOverTime.Enabled = False
        ADDOREDIT = "EDIT"
        cmdOvertime.Visible = True
        cmdOvertime.ZOrder 0
        picOvertime.Visible = True
        picOvertime.ZOrder 0
        picOvertime.Enabled = True
        Picture1.Visible = False
        Picture2.Visible = True
        EnabledPics False
        StoreEntry fild
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    'txtSearch.Enabled = True
    On Error Resume Next
    txtsearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsEmpInfo.MoveNext
    If rsEmpInfo.EOF Then
        rsEmpInfo.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsEmpInfo.MovePrevious
    If rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    'On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Print", "EMPLOYEE MAINTAIN OVERTIME") = False Then Exit Sub
    Screen.MousePointer = 11
    rptOvertime.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptOvertime.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptOvertime.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptOvertime.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
    PrintSQLReport rptOvertime, HRMS_REPORT_PATH & "OVERTIME.RPT", "{OVERTIME.EMPNO} = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND {OVERTIME.CUT_OFF} = '" & CUTTOFF_CODE & "' AND {OVERTIME.PAY_MONTH} = " & PAY_MONTH & " AND {OVERTIME.PAY_YEAR} = " & PAY_YEAR & "", DMIS_REPORT_Connection, 1
    LogAudit "V", "PRINT EMPLOYEE OVERTIME LIST", EMPLOYEE_NO
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
'JBF not applicable to datepicker 8/31/2010
'    If IsDate(cboMonth & "/" & cboDay & "/" & cboYear) = False Then
'        MsgBox "Invalid Date Format", vbInformation, "HRMS"
'        cboDay.SetFocus
'        Exit Sub
'    End If
    
    Dim MM, DD, YY, MM2, DD2, YY2                                     As String
    Dim OvertimeCode                                                  As String
    Dim HDAMOUNT                                                      As Currency
    Dim RDAMOUNT                                                      As Currency
    Dim JUSTIF                                                        As String
    Dim vCUTOFF                                                       As Integer
    If cboQuensina.Text = "" Then
        ShowIsRequiredMsg "Choose a Cut-Off"
        cboQuensina.SetFocus
        Exit Sub
    End If
    JUSTIF = N2Str2Null(txtJustif.Text)
    OvertimeCode = Right(cboOT, 3)
    If cboQuensina.Text = "1st Cut-Off" Then
        vCUTOFF = 1
    End If
    If cboQuensina.Text = "2nd Cut-Off" Then
        vCUTOFF = 2
    End If
    
   'JBF not applicable to datepicker 8/31/2010
'    MM = What_month(cboMonth)
'    YY = cboYear.Text
'    DD = cboDay.Text
'    Diyt = DateSerial(YY, MM, DD)
    '****************************************
    
    
    MM = MONTH(dt_ot)
    YY = YEAR(dt_ot)
    DD = Day(dt_ot)
    Diyt = DateSerial(YY, MM, DD)
    
        
    If ADDOREDIT = "ADD" Then
        'COMMENT BY : MJP 11072008
            'gconDMIS.Execute "INSERT INTO HRMS_OVERTIME " & _
                             "(EMPLEVEL, EMPNO, OCODE, DEYT, DEYT2, TOTALHR, AMOUNT, HOLIDAY, JUSTIFICATION, CUT_OFF, PAY_MONTH, PAY_YEAR, TEMP_MONTH) " & _
                           " VALUES (" & EMPLIVIL & "," & N2Str2Null(rsEmpInfo!EMPNO) & ", '" & OvertimeCode & "', " & N2Date2Null(Diyt) & _
                             ", " & N2Date2Null(Diyt2) & _
                             ", " & NumericVal(txtTotalHr.Text) & _
                             ", " & NumericVal(txtAmount.Text) & _
                             "," & txtHoliday.Text & _
                             "," & JUSTIF & _
                             "," & vCUTOFF & _
                             "," & MM & _
                             "," & YY & "," &  & " ")"
        'COMMENT BY : MJP 11072008
        
        'UPDATE BY   : MJP 11072008
        'DESCRIPTION :
            gconDMIS.Execute "INSERT INTO HRMS_OVERTIME " & _
                         "(EMPLEVEL, EMPNO, OCODE, DEYT, DEYT2, TOTALHR, AMOUNT, HOLIDAY, JUSTIFICATION, CUT_OFF, PAY_MONTH, PAY_YEAR, MANUAL) " & _
                       " VALUES (" & EMPLIVIL & "," & N2Str2Null(rsEmpInfo!EMPNO) & ", '" & OvertimeCode & "', " & N2Date2Null(Diyt) & _
                         ", " & N2Date2Null(Diyt2) & _
                         ", " & NumericVal(txtTotalHr.Text) & _
                         ", " & NumericVal(txtAmount.Text) & _
                         "," & txtHoliday.Text & _
                         "," & JUSTIF & _
                         "," & vCUTOFF & _
                         "," & What_month(LABMONTH) & _
                         "," & PAY_YEAR & ",'Y')"
        'UPDATE BY   : MJP 11072008
        
        LogAudit "A", "ADD OVERTIME LIST ON A EMPLOYEE", EMPLOYEE_NO & "-" & OVERTIME_CODES
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "UPDATE HRMS_OVERTIME SET" & _
                       " EMPLEVEL = " & EMPLIVIL & "," & _
                       " EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & "," & _
                       " OCODE = '" & OvertimeCode & "'," & _
                       " DEYT = " & N2Date2Null(Diyt) & "," & _
                       " DEYT2 = " & N2Date2Null(Diyt2) & "," & _
                       " TOTALHR = " & NumericVal(txtTotalHr.Text) & "," & _
                       " AMOUNT = " & NumericVal(txtAmount.Text) & "," & _
                       " HOLIDAY  = " & NumericVal(txtHoliday.Text) & "," & _
                       " JUSTIFICATION = " & JUSTIF & "," & _
                       " CUT_OFF = " & vCUTOFF & "," & _
                       " PAY_MONTH = " & What_month(LABMONTH) & "," & _
                       " PAY_YEAR = " & PAY_YEAR & _
                       " WHERE ID = " & LABID.Caption

        LogAudit "E", "EDIT OVERTIME LIST ON A EMPLOYEE", EMPLOYEE_NO & "-" & OVERTIME_CODES
        ShowSuccessFullyUpdated
    End If
    cmdCancel.Value = True
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            cmdAdd.Value = True
        Case vbKeyEscape
            cmdCancel.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then
        EMPLIVIL = "'C'"
    End If
    If EMP_TYPE = "ALLOWANCE BASE" Then
        EMPLIVIL = "'A'"
    End If
    
    Call rsSETUP
    
    
    Call rsrefresh
    Call txtsearch_Change
    Call InitGrid
    Call InitMemvars
    Call cmdCancel_Click
    Call FillGrid
    Call FillCboOT
    
    'DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub grdOvertime_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub chkHoliday_Click()
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
    End If
End Sub

Private Sub optRegular_Click()
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
    End If
End Sub

Private Sub optSunday_Click()
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
    End If
End Sub

Private Sub txtJustif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtNoMin_Computed_Change()
    If NumericVal(txtNoMin_Computed) > 59 Then
        txtNoMin_Computed = 59
    End If
    ComputeHours
End Sub

Private Sub txtNoMin_Computed_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub txtNoHours_Computed_Change()
    ComputeHours
End Sub

Private Sub txtNoHours_Computed_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtHoliday_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtTotalHr_Change()
    Dim HOLID_REGUL                                                   As String
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        FindOTRate HOLID_REGUL
        UpdateAmount_NEW HOLID_REGUL
    End If
End Sub

Private Sub txtTotalHr_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtTotalHr_LostFocus()
    Dim HOLID_REGUL                                                   As String
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        UpdateAmount_NEW HOLID_REGUL
    End If
End Sub

Private Sub lstOverTime_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    'On Error Resume Next
    'rsEmpInfo.MoveFirst
    'rsEmpInfo.Find ("id=" & ITEM.ListSubItems(2).Text)
    'StoreMemVars

    rsEmpInfo.Requery
    rsEmpInfo.Find ("EMPNO=" & ITEM.ListSubItems(1).Text)
    StoreMemVars


End Sub

Private Sub lstOverTime_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOverTime
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstOverTime_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtsearch_Change()
    If Trim(txtsearch.Text) = "" Then
        Call FillGrid
    Else
        Call FillSearchGrid(txtsearch.Text)
    End If
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMS_OT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overtime Entry"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5490
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Overtimex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5490
   Begin Crystal.CrystalReport rptOvertime 
      Left            =   90
      Top             =   4920
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
      Height          =   1125
      Left            =   10680
      ScaleHeight     =   1125
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   6240
      Width           =   435
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
         Width           =   315
      End
      Begin MSComctlLib.ListView lstOverTime 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
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
      Height          =   4785
      Left            =   30
      ScaleHeight     =   4755
      ScaleWidth      =   5385
      TabIndex        =   13
      Top             =   0
      Width           =   5415
      Begin VB.ComboBox cboDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   21
         Text            =   "cboDay"
         Top             =   1890
         Width           =   945
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4230
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   22
         Text            =   "cboYear"
         Top             =   1890
         Width           =   945
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   690
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Text            =   "cboMonth"
         Top             =   1890
         Width           =   1455
      End
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
         Left            =   120
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   19
         Text            =   "cboQuensina"
         Top             =   1290
         Width           =   2385
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
         Left            =   990
         TabIndex        =   23
         Top             =   2460
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
         Left            =   3210
         TabIndex        =   24
         Top             =   2460
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
         Height          =   1305
         Left            =   150
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   3210
         Width           =   4965
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
         TabIndex        =   18
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
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3780
         Visible         =   0   'False
         Width           =   495
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
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3780
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   315
         Left            =   3180
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3750
         Width           =   345
         _ExtentX        =   609
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
         Left            =   420
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3450
         Width           =   375
         _ExtentX        =   661
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
         Left            =   420
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3780
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3780
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   90
         X2              =   5220
         Y1              =   1770
         Y2              =   1770
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
         Left            =   3660
         TabIndex        =   64
         Top             =   1980
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
         Left            =   30
         TabIndex        =   63
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Left            =   2250
         TabIndex        =   62
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label Label6 
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
         Height          =   225
         Left            =   90
         TabIndex        =   61
         Top             =   1080
         Width           =   1185
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   -30
         Width           =   5475
         _Version        =   655364
         _ExtentX        =   9657
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "  Add Over Time"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.24
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
         X2              =   5250
         Y1              =   2910
         Y2              =   2910
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
         TabIndex        =   38
         Top             =   2520
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
         TabIndex        =   33
         Top             =   2520
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
         Left            =   90
         TabIndex        =   43
         Top             =   2940
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
         Left            =   840
         TabIndex        =   42
         Top             =   3540
         Width           =   1590
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   5250
         Y1              =   2340
         Y2              =   2340
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
         Left            =   2970
         TabIndex        =   15
         Top             =   390
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   14
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
         Left            =   2460
         TabIndex        =   16
         Top             =   3540
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
         TabIndex        =   37
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   36
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         Left            =   2760
         TabIndex        =   17
         Top             =   3750
         Width           =   345
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
         Left            =   2460
         TabIndex        =   34
         Top             =   3540
         Width           =   1515
      End
   End
   Begin wizButton.cmd cmdOvertime 
      Height          =   4575
      Left            =   -30
      TabIndex        =   45
      Top             =   30
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   8070
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
      MICON           =   "Overtimex.frx":628A
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   3645
      TabIndex        =   46
      Top             =   7530
      Width           =   3645
      Begin MSFlexGridLib.MSFlexGrid grdOvertime 
         Height          =   705
         Left            =   450
         TabIndex        =   47
         Top             =   60
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1244
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
      Height          =   2145
      Left            =   4620
      ScaleHeight     =   2145
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   7080
      Width           =   4635
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
         Left            =   5880
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   465
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
         Left            =   6420
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   465
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
         Left            =   600
         TabIndex        =   5
         Top             =   1350
         Visible         =   0   'False
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
         Left            =   5400
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   60
         TabIndex        =   49
         Top             =   540
         Visible         =   0   'False
         Width           =   1155
         _Version        =   655364
         _ExtentX        =   2037
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
         Left            =   3690
         TabIndex        =   11
         Top             =   1170
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
         Left            =   -240
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label labID 
         Caption         =   "Label6"
         Height          =   285
         Left            =   1530
         TabIndex        =   8
         Top             =   1050
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
         Left            =   1200
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
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   -60
      ScaleHeight     =   975
      ScaleWidth      =   5820
      TabIndex        =   51
      Top             =   7530
      Width           =   5820
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
         MouseIcon       =   "Overtimex.frx":62A6
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":63F8
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Visible         =   0   'False
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
         MouseIcon       =   "Overtimex.frx":6757
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":68A9
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Visible         =   0   'False
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
         MouseIcon       =   "Overtimex.frx":6C01
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":6D53
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Find a Record"
         Top             =   30
         Visible         =   0   'False
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
         MouseIcon       =   "Overtimex.frx":704D
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":719F
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Add Record"
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
         MouseIcon       =   "Overtimex.frx":74B2
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":7604
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Edit Selected Record"
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
         MouseIcon       =   "Overtimex.frx":7960
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":7AB2
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Delete Selected Record"
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
         MouseIcon       =   "Overtimex.frx":7DDD
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":7F2F
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
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
         MouseIcon       =   "Overtimex.frx":8295
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":83E7
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   3450
      ScaleHeight     =   1155
      ScaleWidth      =   2490
      TabIndex        =   60
      Top             =   4710
      Width           =   2490
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Approved"
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
         Left            =   210
         MouseIcon       =   "Overtimex.frx":874D
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":889F
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Save Entry"
         Top             =   150
         Width           =   855
      End
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
         Left            =   1080
         MouseIcon       =   "Overtimex.frx":8BEF
         MousePointer    =   99  'Custom
         Picture         =   "Overtimex.frx":8D41
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cancel"
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.Label LABMONTH 
      BackColor       =   &H000000FF&
      Height          =   345
      Left            =   11280
      TabIndex        =   50
      Top             =   6480
      Width           =   165
   End
End
Attribute VB_Name = "frmHRMS_OT"
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
Dim XEMPNO                                                              As String

Dim xdate                                                               As String
Dim XMONTH                                                              As String
Dim xin                                                                 As String
Dim xout                                                                As String

Dim BegDayFirstCutOff                                                   As Integer
Dim EndDayFirstCutOff                                                   As Integer
Dim BegDaySecondCutOff                                                  As Integer
Dim EndDaySecondCutOff                                                  As Integer




Public Event SelectionMade()

Public Sub SelectSQl(XXX As String, xxempno As String, xxdate As String, xxmonth As String, xxin As String, xxout As String)
    
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open XXX, gconDMIS, adOpenForwardOnly, adLockReadOnly
    XEMPNO = xxempno
    xdate = xxdate
    XMONTH = xxmonth
    xin = xxin
    xout = xxout
    
             
End Sub

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
        labID.Caption = rsOvertime!ID
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
        DD = Day(TheDeyt)
        
        MM = The_month(MONTH(TheDeyt))
        YY = YEAR(TheDeyt)
        cboDay.Text = DD
        cboMonth.Text = MM
        cboYear.Text = YY
        
        'dt_ot = Null2Date(rsOvertime!DEYT)
        
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
        rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & EmpInfoEmpno.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & frmHRMSEmpInfo.labID.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
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

Sub AddMonthName()
    Dim X As Integer
    cboMonth.Clear
    For X = 1 To 12
        cboMonth.AddItem MonthName(X)
    Next
End Sub

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
            Call AddMonthName
            cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
            LABMONTH.Caption = MonthName(Null2String(rsCutoff!PERIODMONTH))
        'UPDATE BY : MJP 11072008
        cboYear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If
    fillcboDay cboDay
    cboDay.Text = Day(Now)
    fillcboDay cboDay2
    cboDay.Text = Day(LOGDATE)
    cboDay2.Text = Day(LOGDATE)
    
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
                labID.Caption = rsOvertime!ID
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
    'Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME,EMPNO ,ID FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno ,ID from HRMS_EmpInfo WHERE RESIGNED IS NULL order by lastname+', '+firstname asc")
    
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
    'If Function_Access(LOGID, "Acess_Add", "EMPLOYEE OVERTIME DEDUCTION") = False Then Exit Sub
    
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
    'InitMemvars
    
    initadd
    
    Exit Sub
    
'Errorcode:
    'ShowVBError
End Sub

Sub initadd()

        Dim hours                                                            As Integer
        Dim minutes                                                          As Integer
        Dim total_ot                                                        As Integer
        
        
        Dim rsPayrollSetup                                            As ADODB.Recordset
        Set rsPayrollSetup = New ADODB.Recordset
        Set rsPayrollSetup = gconDMIS.Execute("Select * from HRMS_PayrollSetup")
        If Not rsPayrollSetup.EOF And Not rsPayrollSetup.BOF Then
            BegDayFirstCutOff = N2Str2Zero(rsPayrollSetup!FROMDATE1)
            BegDaySecondCutOff = N2Str2Zero(rsPayrollSetup!FROMDATE2)
            EndDayFirstCutOff = N2Str2Zero(rsPayrollSetup!TODATE1)
            EndDaySecondCutOff = N2Str2Zero(rsPayrollSetup!TODATE2)
            
            If CDate(Format(xdate, "mm/dd/yyyy")) <= CDate(Format(DateSerial(cboYear, What_month(RTrim(LTrim(XMONTH))), EndDayFirstCutOff), "mm/dd/yyyy")) Then
                cboQuensina.Text = "1st Cut-Off"
            Else
                cboQuensina.Text = "2nd Cut-Off"
            End If
            
                Call AddMonthName
                
                cboMonth.Text = Format(xdate, "mmmm")
                LABMONTH.Caption = Format(xdate, "mmmm")
                cboDay.Text = Day(xdate)
                cboYear.Text = YEAR(xdate)
            
            
             If xin = "8:00:00 AM" And xout = "5:00:00 PM" Then
                    txtNoHours_Computed.Text = ""
                    txtNoMin_Computed.Text = ""
        
             ElseIf xout > "5:00:00 PM" Then
            
                    hours = Hour(xout)
                    minutes = Minute(xout)
                    total_ot = hours - 17
                    txtNoHours_Computed.Text = total_ot
                    txtNoMin_Computed.Text = minutes
             Else
                'do nothing
            End If
        End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me

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
    txtSearch.SetFocus
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

Private Sub cmdPreviousx_Click()

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
    Dim RSCHECK                                                       As ADODB.Recordset
    Dim MM, DD, YY, MM2, DD2, YY2                                     As String
    Dim OvertimeCode                                                  As String
    Dim HDAMOUNT                                                      As Currency
    Dim RDAMOUNT                                                      As Currency
    Dim JUSTIF                                                        As String
    Dim vCUTOFF                                                       As Integer
   
   
    
    MM = What_month(cboMonth)
    YY = cboYear.Text
    DD = cboDay.Text
    Diyt = DateSerial(YY, MM, DD)
    
   
   If txtJustif.Text = "" Then
        MsgBox "Specify the reason for overtime", vbInformation, "HRMS"
        txtJustif.SetFocus
        Exit Sub
   End If
    

   If ADDOREDIT = "ADD" Then
   
        Set RSCHECK = New ADODB.Recordset
        RSCHECK.Open "select deyt from hrms_overtime where empno = '" & XEMPNO & "' and deyt = '" & Diyt & "'", gconDMIS
        If Not RSCHECK.BOF And Not RSCHECK.EOF Then
            MsgSpeechBox "The Date that has been selected has already have data!" & vbCrLf & _
                          "Check the Details in Overtime Module. "
            On Error Resume Next
                txtNoHours_Computed.SetFocus
            Exit Sub
        End If
    Else
    
        ' do nothing
   End If
      
    
    If IsDate(cboMonth & "/" & cboDay & "/" & cboYear) = False Then
        MsgBox "Invalid Date Format", vbInformation, "HRMS"
        cboDay.SetFocus
        Exit Sub
    End If

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
                       " WHERE ID = " & labID.Caption

        LogAudit "E", "EDIT OVERTIME LIST ON A EMPLOYEE", EMPLOYEE_NO & "-" & OVERTIME_CODES
        ShowSuccessFullyUpdated
    End If
    'CmdCancel.Value = True
    
            'lstOverTime.Enabled = True
            'ADDOREDIT = ""
'    picOvertime.Visible = True
'    cmdOvertime.Visible = True
'    picOvertime.Enabled = True
'            'Picture1.Visible = True
'    Picture2.Visible = True
'    cmdOvertime.ZOrder 0
'    picOvertime.ZOrder 0
            'EnabledPics True
            'StoreMemVars
    
    Unload Me
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Command8_Click()

End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            cmdAdd.Value = True
        Case vbKeyEscape
            CmdCancel.Value = True
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
    
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPNO  = '" & XEMPNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    
    Call InitGrid
    Call InitMemvars
            'Call cmdCancel_Click
            'Call FillGrid
    Call FillCboOT
    DrawXPCtl Me

    Call cmdAdd_Click
    
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
'    'On Error Resume Next
'    rsEmpInfo.MoveFirst
'    rsEmpInfo.Find ("id=" & ITEM.ListSubItems(2).Text)
'    StoreMemVars


    On Error Resume Next
    rsEmpInfo.Bookmark = rsFIND(rsEmpInfo.Clone, "empno", lstOverTime.SelectedItem.SubItems(1)).Bookmark
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
    If Trim(txtSearch.Text) = "" Then
        Call FillGrid
    Else
        Call FillSearchGrid(txtSearch.Text)
    End If
End Sub

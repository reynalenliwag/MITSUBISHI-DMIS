VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "WIZBOX.OCX"
Object = "{E6BE8522-29DC-4EDD-813C-BAA34BBA1069}#2.0#0"; "WIZMACFORM.OCX"
Begin VB.Form frmHRMSLedger 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Employee Ledger"
   ClientHeight    =   6915
   ClientLeft      =   300
   ClientTop       =   480
   ClientWidth     =   11610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Haettenschweiler"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Ledger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Ledger.frx":030A
   ScaleHeight     =   6915
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab TabSSS 
      Height          =   4365
      Left            =   2640
      TabIndex        =   26
      Top             =   1560
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   7699
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   1058
      BackColor       =   14606302
      MouseIcon       =   "Ledger.frx":3046
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Payroll"
      TabPicture(0)   =   "Ledger.frx":3062
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picPayroll"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDetails"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Loans"
      TabPicture(1)   =   "Ledger.frx":34B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "SSS/PHIC."
      TabPicture(2)   =   "Ledger.frx":3906
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pag-Ibig / TAX"
      TabPicture(3)   =   "Ledger.frx":4870
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "YTD Details"
      TabPicture(4)   =   "Ledger.frx":4B8A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture6"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3645
         Left            =   60
         TabIndex        =   91
         Top             =   630
         Width           =   8775
         Begin MSFlexGridLib.MSFlexGrid grdPayroll 
            Height          =   3435
            Left            =   60
            TabIndex        =   73
            Top             =   150
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   6059
            _Version        =   393216
            Cols            =   26
            FixedCols       =   2
            ForeColor       =   0
            BackColorFixed  =   14606302
            ForeColorFixed  =   0
            BackColorSel    =   14606302
            ForeColorSel    =   0
            BackColorBkg    =   14606302
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
      Begin VB.PictureBox Picture6 
         Height          =   3585
         Left            =   -74970
         ScaleHeight     =   3525
         ScaleWidth      =   8775
         TabIndex        =   124
         Top             =   690
         Width           =   8835
         Begin TabDlg.SSTab TabYTD 
            Height          =   3525
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   6218
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BackColor       =   14606302
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Previous Payrolls YTD Details"
            TabPicture(0)   =   "Ledger.frx":4EA4
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "cmdEditPrevYTD"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdUpdatePrevYTD"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "picPrevYTD"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Picture10"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Generated Payrolls YTD Details"
            TabPicture(1)   =   "Ledger.frx":4EC0
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "picYTD"
            Tab(1).ControlCount=   1
            Begin VB.PictureBox Picture10 
               BackColor       =   &H00DEDFDE&
               Height          =   2925
               Left            =   5580
               ScaleHeight     =   2865
               ScaleWidth      =   2985
               TabIndex        =   136
               Top             =   450
               Width           =   3045
               Begin VB.Image Image1 
                  Height          =   6960
                  Left            =   30
                  Picture         =   "Ledger.frx":4EDC
                  Top             =   30
                  Width           =   9915
               End
            End
            Begin VB.PictureBox picPrevYTD 
               BackColor       =   &H00DEDFDE&
               Enabled         =   0   'False
               Height          =   2295
               Left            =   210
               Picture         =   "Ledger.frx":18C39
               ScaleHeight     =   2235
               ScaleWidth      =   5115
               TabIndex        =   129
               Top             =   450
               Width           =   5175
               Begin MSMask.MaskEdBox txtPYTDSSS 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   69
                  Top             =   780
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  BackColor       =   16777215
                  ForeColor       =   0
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
               Begin MSMask.MaskEdBox txtPYTDGross 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   67
                  Top             =   60
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  BackColor       =   16777215
                  ForeColor       =   0
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
               Begin MSMask.MaskEdBox txtPYTDTax 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   68
                  Top             =   420
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  BackColor       =   16777215
                  ForeColor       =   0
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
               Begin MSMask.MaskEdBox txtPYTDPHIC 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   70
                  Top             =   1140
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  BackColor       =   16777215
                  ForeColor       =   0
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
               Begin MSMask.MaskEdBox txtPYTDPagIbig 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   71
                  Top             =   1500
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  BackColor       =   16777215
                  ForeColor       =   0
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
               Begin MSMask.MaskEdBox txtPYTDMidYear 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   72
                  Top             =   1860
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  Appearance      =   0
                  BackColor       =   16777215
                  ForeColor       =   0
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
               Begin VB.Label Label78 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "YTD PAG-IBIG CONTRIBUTION"
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
                  Height          =   315
                  Left            =   90
                  TabIndex        =   135
                  Top             =   1500
                  Width           =   3075
               End
               Begin VB.Label Label77 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "MID YEAR PAY"
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
                  Height          =   315
                  Left            =   90
                  TabIndex        =   134
                  Top             =   1860
                  Width           =   3075
               End
               Begin VB.Label Label74 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "YTD GROSS"
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
                  Height          =   315
                  Left            =   90
                  TabIndex        =   133
                  Top             =   60
                  Width           =   3075
               End
               Begin VB.Label Label73 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "YTD WITHHOLDING TAX"
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
                  Height          =   315
                  Left            =   90
                  TabIndex        =   132
                  Top             =   420
                  Width           =   3075
               End
               Begin VB.Label Label72 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "YTD SSS CONTRIBUTION"
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
                  Height          =   315
                  Left            =   90
                  TabIndex        =   131
                  Top             =   780
                  Width           =   3075
               End
               Begin VB.Label Label8 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "YTD PHIC CONTRIBUTION"
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
                  Height          =   315
                  Left            =   90
                  TabIndex        =   130
                  Top             =   1140
                  Width           =   3075
               End
            End
            Begin VB.PictureBox picYTD 
               Height          =   3075
               Left            =   -74940
               Picture         =   "Ledger.frx":1B975
               ScaleHeight     =   3015
               ScaleWidth      =   8595
               TabIndex        =   128
               Top             =   360
               Width           =   8655
               Begin MSMask.MaskEdBox txtYTDCommission 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   150
                  Top             =   630
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
               Begin MSMask.MaskEdBox txtYTDRemSal 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   151
                  Top             =   1230
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
               Begin MSMask.MaskEdBox txtYTDGross 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   152
                  Top             =   2130
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
               Begin MSMask.MaskEdBox txtYTDBasicPay 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   153
                  Top             =   30
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
               Begin MSMask.MaskEdBox txtYTDOvertime 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   154
                  Top             =   330
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
               Begin MSMask.MaskEdBox txtYTDSSSPAGIBIGPHIC 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   155
                  Top             =   2430
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
               Begin MSMask.MaskEdBox txtTaxExemp 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   156
                  Top             =   2730
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
               Begin MSMask.MaskEdBox txtYTDTaxDue 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   157
                  Top             =   600
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
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox txtYTDTaxWithHeld 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   158
                  Top             =   900
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
               Begin MSMask.MaskEdBox txtYTDTaxRefund 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   159
                  Top             =   1200
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
               Begin MSMask.MaskEdBox txt13thMonth 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   160
                  Top             =   1800
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
               Begin MSMask.MaskEdBox txtTaxRefund 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   161
                  Top             =   2100
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
               Begin MSMask.MaskEdBox txtAdjSalary 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   162
                  Top             =   1500
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
               Begin MSMask.MaskEdBox txtTotalPay 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   163
                  Top             =   2700
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
               Begin MSMask.MaskEdBox txtTaxableIncome 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   164
                  Top             =   30
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
               Begin MSMask.MaskEdBox txtMidYear 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   165
                  Top             =   930
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
               Begin MSMask.MaskEdBox txtYTDRemOT 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   166
                  Top             =   1530
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
               Begin MSMask.MaskEdBox txtYTDRemWTax 
                  Height          =   255
                  Left            =   2850
                  TabIndex        =   167
                  Top             =   1830
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
               Begin MSMask.MaskEdBox txtYTDRemDed 
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   168
                  Top             =   2400
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
               Begin VB.Label Label71 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "REMAINING DEDUCTION"
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
                  Left            =   4470
                  TabIndex        =   187
                  Top             =   2400
                  Width           =   2295
               End
               Begin VB.Label Label70 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "REMAINING TAX WITHHELD"
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
                  Left            =   90
                  TabIndex        =   186
                  Top             =   1830
                  Width           =   2445
               End
               Begin VB.Label Label69 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "REMAINING OVERTIME PAY"
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
                  Left            =   90
                  TabIndex        =   185
                  Top             =   1530
                  Width           =   2445
               End
               Begin VB.Label Label68 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "MID YEAR PAY"
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
                  Left            =   90
                  TabIndex        =   184
                  Top             =   930
                  Width           =   2175
               End
               Begin VB.Label Label67 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAXABLE INCOME"
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
                  Left            =   4470
                  TabIndex        =   183
                  Top             =   30
                  Width           =   1845
               End
               Begin VB.Label Label66 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL PAY"
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
                  Left            =   4470
                  TabIndex        =   182
                  Top             =   2700
                  Width           =   1845
               End
               Begin VB.Label Label65 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ADJUSTED SALARY"
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
                  Left            =   4470
                  TabIndex        =   181
                  Top             =   1500
                  Width           =   1845
               End
               Begin VB.Label Label64 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX REFUND"
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
                  Left            =   4470
                  TabIndex        =   180
                  Top             =   2100
                  Width           =   1845
               End
               Begin VB.Label Label58 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "13TH MONTH PAY"
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
                  Left            =   4470
                  TabIndex        =   179
                  Top             =   1800
                  Width           =   1845
               End
               Begin VB.Label Label63 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "(TAX REFUND / TAX PAYABLE)"
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
                  Left            =   4470
                  TabIndex        =   178
                  Top             =   1200
                  Width           =   2685
               End
               Begin VB.Label Label62 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX WITHHELD"
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
                  Left            =   4470
                  TabIndex        =   177
                  Top             =   900
                  Width           =   1845
               End
               Begin VB.Label Label61 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX DUE"
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
                  Left            =   4470
                  TabIndex        =   176
                  Top             =   600
                  Width           =   1845
               End
               Begin VB.Label Label54 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX EXEMPTION"
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
                  Left            =   900
                  TabIndex        =   175
                  Top             =   2730
                  Width           =   1695
               End
               Begin VB.Label Label53 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "LESS:     SSS/PAG-IBIG /PHIC"
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
                  Left            =   90
                  TabIndex        =   174
                  Top             =   2430
                  Width           =   2685
               End
               Begin VB.Label Label60 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "COMMISSION"
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
                  Left            =   90
                  TabIndex        =   173
                  Top             =   630
                  Width           =   2175
               End
               Begin VB.Label Label59 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "REMAINING SALARY"
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
                  Left            =   90
                  TabIndex        =   172
                  Top             =   1230
                  Width           =   2175
               End
               Begin VB.Label Label57 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "GROSS PAY"
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
                  Left            =   90
                  TabIndex        =   171
                  Top             =   2130
                  Width           =   2445
               End
               Begin VB.Label Label56 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "OVERTIME/OTHERS"
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
                  Left            =   90
                  TabIndex        =   170
                  Top             =   330
                  Width           =   2175
               End
               Begin VB.Label Label55 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "BASIC"
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
                  Left            =   90
                  TabIndex        =   169
                  Top             =   30
                  Width           =   2175
               End
            End
            Begin wizButton.cmd cmdUpdatePrevYTD 
               Height          =   495
               Left            =   2850
               TabIndex        =   137
               Top             =   2850
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   873
               TX              =   "Update Previous YTD"
               ENAB            =   0   'False
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
               MICON           =   "Ledger.frx":1E6B1
            End
            Begin wizButton.cmd cmdEditPrevYTD 
               Height          =   495
               Left            =   630
               TabIndex        =   138
               Top             =   2850
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   873
               TX              =   "Edit Previous YTD"
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
               MICON           =   "Ledger.frx":1E6CD
            End
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00DEDFDE&
         Height          =   3555
         Left            =   -74940
         Picture         =   "Ledger.frx":1E6E9
         ScaleHeight     =   3495
         ScaleWidth      =   8775
         TabIndex        =   109
         Top             =   690
         Width           =   8835
         Begin VB.Frame fraPagIbigTIN 
            Appearance      =   0  'Flat
            BackColor       =   &H00DEDFDE&
            ForeColor       =   &H80000008&
            Height          =   1665
            Left            =   30
            TabIndex        =   110
            Top             =   -60
            Width           =   8685
            Begin VB.TextBox txtPagIbigNo 
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
               Left            =   1440
               TabIndex        =   59
               Top             =   180
               Width           =   2145
            End
            Begin VB.TextBox txtTINNo 
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
               Left            =   5730
               TabIndex        =   63
               Top             =   180
               Width           =   2265
            End
            Begin MSMask.MaskEdBox txtPagIbigMonthly 
               Height          =   315
               Left            =   2070
               TabIndex        =   60
               Top             =   540
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
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
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtTINMonthly 
               Height          =   315
               Left            =   6300
               TabIndex        =   64
               Top             =   540
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
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
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPagIbigStarted 
               Height          =   315
               Left            =   2070
               TabIndex        =   61
               Top             =   900
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtTINStarted 
               Height          =   315
               Left            =   6300
               TabIndex        =   65
               Top             =   900
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPagIbigLast 
               Height          =   315
               Left            =   2070
               TabIndex        =   62
               Top             =   1260
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtTINLast 
               Height          =   315
               Left            =   6300
               TabIndex        =   66
               Top             =   1260
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label28 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Date of Cont."
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
               Left            =   4380
               TabIndex        =   118
               Top             =   1290
               Width           =   1965
            End
            Begin VB.Label Label27 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Date of Cont."
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
               Left            =   90
               TabIndex        =   117
               Top             =   1290
               Width           =   1875
            End
            Begin VB.Label Label26 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Started"
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
               Left            =   4380
               TabIndex        =   116
               Top             =   930
               Width           =   1395
            End
            Begin VB.Label Label25 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Started"
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
               Left            =   90
               TabIndex        =   115
               Top             =   930
               Width           =   1365
            End
            Begin VB.Label Label24 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Ded."
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
               Left            =   4380
               TabIndex        =   114
               Top             =   570
               Width           =   1395
            End
            Begin VB.Label Label23 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Ded."
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
               Left            =   90
               TabIndex        =   113
               Top             =   570
               Width           =   1365
            End
            Begin VB.Label Label15 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Pag-Ibig No."
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
               Left            =   90
               TabIndex        =   112
               Top             =   210
               Width           =   1365
            End
            Begin VB.Label Label18 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "TIN No."
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
               Left            =   4380
               TabIndex        =   111
               Top             =   210
               Width           =   1395
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdPagIbig 
            Height          =   1815
            Left            =   60
            TabIndex        =   77
            Top             =   1650
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   4
            BackColorFixed  =   14606302
            ForeColorFixed  =   0
            BackColorSel    =   14606302
            ForeColorSel    =   0
            BackColorBkg    =   14606302
            Appearance      =   0
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
         Begin MSFlexGridLib.MSFlexGrid grdTIN 
            Height          =   1815
            Left            =   4380
            TabIndex        =   78
            Top             =   1650
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   3
            BackColorFixed  =   14606302
            ForeColorFixed  =   0
            BackColorSel    =   14606302
            ForeColorSel    =   0
            BackColorBkg    =   14606302
            Appearance      =   0
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
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00DEDFDE&
         Height          =   3555
         Left            =   -74940
         Picture         =   "Ledger.frx":21425
         ScaleHeight     =   3495
         ScaleWidth      =   8775
         TabIndex        =   99
         Top             =   690
         Width           =   8835
         Begin VB.Frame fraSSSMED 
            Appearance      =   0  'Flat
            BackColor       =   &H00DEDFDE&
            ForeColor       =   &H80000008&
            Height          =   1665
            Left            =   30
            TabIndex        =   100
            Top             =   -60
            Width           =   8685
            Begin VB.TextBox txtSSSNo 
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
               Left            =   1440
               TabIndex        =   51
               Top             =   180
               Width           =   2145
            End
            Begin VB.TextBox txtPhilHealthNo 
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
               Left            =   5730
               TabIndex        =   55
               Top             =   180
               Width           =   2265
            End
            Begin MSMask.MaskEdBox txtSSSMonthly 
               Height          =   315
               Left            =   2070
               TabIndex        =   52
               Top             =   540
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
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
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPHMonthly 
               Height          =   315
               Left            =   6300
               TabIndex        =   56
               Top             =   540
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
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
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtSSSStarted 
               Height          =   315
               Left            =   2070
               TabIndex        =   53
               Top             =   900
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPHStarted 
               Height          =   315
               Left            =   6300
               TabIndex        =   57
               Top             =   900
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtSSSLast 
               Height          =   315
               Left            =   2070
               TabIndex        =   54
               Top             =   1260
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPHLast 
               Height          =   315
               Left            =   6300
               TabIndex        =   58
               Top             =   1260
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label22 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Date of Cont."
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
               Left            =   4380
               TabIndex        =   108
               Top             =   1290
               Width           =   1875
            End
            Begin VB.Label Label21 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Date of Cont."
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
               Left            =   90
               TabIndex        =   107
               Top             =   1290
               Width           =   1965
            End
            Begin VB.Label Label20 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Started"
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
               Left            =   4380
               TabIndex        =   106
               Top             =   930
               Width           =   1455
            End
            Begin VB.Label Label19 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Started"
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
               Left            =   90
               TabIndex        =   105
               Top             =   930
               Width           =   1395
            End
            Begin VB.Label Label17 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Ded."
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
               Left            =   4380
               TabIndex        =   104
               Top             =   570
               Width           =   1455
            End
            Begin VB.Label Label16 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Ded."
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
               Left            =   90
               TabIndex        =   103
               Top             =   570
               Width           =   1395
            End
            Begin VB.Label Label11 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "SSS No."
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
               Left            =   90
               TabIndex        =   102
               Top             =   210
               Width           =   1395
            End
            Begin VB.Label Label9 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "PH No."
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
               Left            =   4380
               TabIndex        =   101
               Top             =   210
               Width           =   1455
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdPhilHealth 
            Height          =   1815
            Left            =   4380
            TabIndex        =   76
            Top             =   1650
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   4
            BackColorFixed  =   14606302
            ForeColorFixed  =   0
            BackColorSel    =   14606302
            ForeColorSel    =   0
            BackColorBkg    =   14606302
            Appearance      =   0
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
         Begin MSFlexGridLib.MSFlexGrid grdSSS 
            Height          =   1815
            Left            =   60
            TabIndex        =   75
            Top             =   1650
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   4
            BackColorFixed  =   14606302
            ForeColorFixed  =   0
            BackColorSel    =   14606302
            ForeColorSel    =   0
            BackColorBkg    =   14606302
            Appearance      =   0
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
      End
      Begin VB.PictureBox picPayroll 
         Height          =   405
         Left            =   240
         ScaleHeight     =   345
         ScaleWidth      =   2880
         TabIndex        =   82
         Top             =   3120
         Width           =   2940
         Begin VB.CommandButton cmdPrintPayroll 
            Caption         =   "&Print"
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
            Left            =   2160
            TabIndex        =   86
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdDeletePayroll 
            Caption         =   "&Delete"
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
            Left            =   1440
            TabIndex        =   85
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdEditPayroll 
            Caption         =   "&Edit"
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
            Left            =   720
            TabIndex        =   83
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdAddPayroll 
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
            Height          =   360
            Left            =   0
            TabIndex        =   84
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00D8E9EC&
         Height          =   3585
         Left            =   -74940
         Picture         =   "Ledger.frx":24161
         ScaleHeight     =   3525
         ScaleWidth      =   8775
         TabIndex        =   92
         Top             =   690
         Width           =   8835
         Begin MSFlexGridLib.MSFlexGrid grdLoanMas 
            Height          =   3495
            Left            =   30
            TabIndex        =   74
            Top             =   0
            Width           =   8685
            _ExtentX        =   15319
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   10
            BackColorFixed  =   14606302
            ForeColorFixed  =   0
            BackColorSel    =   14606302
            ForeColorSel    =   0
            BackColorBkg    =   14606302
            Appearance      =   0
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
         Begin VB.Frame fraLoanMas 
            Appearance      =   0  'Flat
            BackColor       =   &H00DEDFDE&
            ForeColor       =   &H80000008&
            Height          =   3465
            Left            =   1050
            TabIndex        =   93
            Top             =   -30
            Width           =   6555
            Begin VB.TextBox txtSMonthlyDed 
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
               Left            =   2640
               TabIndex        =   44
               Text            =   "Text1"
               Top             =   2730
               Width           =   1305
            End
            Begin VB.TextBox txtMonthlyDed 
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
               Left            =   2640
               TabIndex        =   43
               Text            =   "Text1"
               Top             =   2370
               Width           =   1305
            End
            Begin VB.TextBox txtOtherTypeDed 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   4050
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               Text            =   "Ledger.frx":2641A
               Top             =   2010
               Width           =   2385
            End
            Begin VB.OptionButton Opt2 
               BackColor       =   &H00DEDFDE&
               Caption         =   "Every 30th Only"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4170
               TabIndex        =   47
               Top             =   600
               Width           =   2175
            End
            Begin VB.OptionButton Opt3 
               BackColor       =   &H00DEDFDE&
               Caption         =   "Every 15th and 30th"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4170
               TabIndex        =   48
               Top             =   930
               Value           =   -1  'True
               Width           =   2175
            End
            Begin VB.OptionButton Opt1 
               BackColor       =   &H00DEDFDE&
               Caption         =   "Every 15th Only"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4170
               TabIndex        =   46
               Top             =   270
               Width           =   2175
            End
            Begin VB.ComboBox cboLoanType 
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
               Left            =   1380
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   180
               Width           =   2565
            End
            Begin MSMask.MaskEdBox txtAmountLoaned 
               Height          =   315
               Left            =   2640
               TabIndex        =   42
               Top             =   2010
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
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
            Begin MSMask.MaskEdBox txtLoanBalance 
               Height          =   315
               Left            =   2640
               TabIndex        =   45
               Top             =   3090
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
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
            Begin MSMask.MaskEdBox txtAcctNo 
               Height          =   315
               Left            =   1380
               TabIndex        =   38
               Top             =   570
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   556
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
            Begin MSMask.MaskEdBox txtDateGranted 
               Height          =   315
               Left            =   2640
               TabIndex        =   39
               Top             =   930
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtMaturityDate 
               Height          =   315
               Left            =   2640
               TabIndex        =   41
               Top             =   1650
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtDateStarted 
               Height          =   315
               Left            =   2640
               TabIndex        =   40
               Top             =   1290
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtOtherTypeDedAmount 
               Height          =   315
               Left            =   4050
               TabIndex        =   50
               Top             =   3060
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   556
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
            Begin wizBox.Box Box1 
               Height          =   1065
               Left            =   4050
               Top             =   210
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   1879
            End
            Begin VB.Label Label81 
               BackColor       =   &H00D8E9EC&
               BackStyle       =   0  'Transparent
               Caption         =   "Deduction Amount"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4050
               TabIndex        =   141
               Top             =   2760
               Width           =   2415
            End
            Begin VB.Label Label79 
               BackColor       =   &H00D8E9EC&
               BackStyle       =   0  'Transparent
               Caption         =   "Other Type of Deduction Deducted to Loan Balance"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   4050
               TabIndex        =   140
               Top             =   1350
               Width           =   2415
            End
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Started"
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
               Left            =   90
               TabIndex        =   139
               Top             =   1320
               Width           =   2445
            End
            Begin VB.Label Label14 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Loan Type"
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
               Left            =   90
               TabIndex        =   122
               Top             =   210
               Width           =   1815
            End
            Begin VB.Label Label13 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Account No."
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
               Left            =   90
               TabIndex        =   121
               Top             =   540
               Width           =   1845
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Granted"
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
               Left            =   90
               TabIndex        =   120
               Top             =   960
               Width           =   2445
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Maturity Date"
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
               Left            =   90
               TabIndex        =   119
               Top             =   1680
               Width           =   2445
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Loan Balance"
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
               Left            =   90
               TabIndex        =   97
               Top             =   3120
               Width           =   2445
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Semi-Monthly Deduction"
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
               Left            =   90
               TabIndex        =   96
               Top             =   2730
               Width           =   2445
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Deduction"
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
               Left            =   90
               TabIndex        =   95
               Top             =   2400
               Width           =   2445
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Amount Loaned"
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
               Left            =   90
               TabIndex        =   94
               Top             =   2040
               Width           =   2445
            End
         End
         Begin VB.CommandButton cmdLoanMas 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Command1"
            Height          =   3495
            Left            =   990
            Picture         =   "Ledger.frx":26420
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   0
            Width           =   6705
         End
         Begin MSFlexGridLib.MSFlexGrid grdLoanMasDet 
            Height          =   2115
            Left            =   30
            TabIndex        =   98
            Top             =   1350
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   3731
            _Version        =   393216
            Cols            =   4
            BackColorFixed  =   14606302
            ForeColorFixed  =   0
            BackColorSel    =   14606302
            ForeColorSel    =   0
            BackColorBkg    =   14606302
            Appearance      =   0
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
      End
   End
   Begin VB.PictureBox picAddEditPayroll 
      Height          =   3465
      Left            =   2760
      Picture         =   "Ledger.frx":286D9
      ScaleHeight     =   3405
      ScaleWidth      =   8595
      TabIndex        =   192
      Top             =   2310
      Width           =   8655
      Begin Crystal.CrystalReport rptPayroll 
         Left            =   60
         Top             =   2220
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSMask.MaskEdBox txtRate 
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtOvertime 
         Height          =   315
         Left            =   1590
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSSS 
         Height          =   315
         Left            =   4800
         TabIndex        =   10
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMed 
         Height          =   315
         Left            =   4800
         TabIndex        =   11
         Top             =   660
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPagIbig 
         Height          =   315
         Left            =   4800
         TabIndex        =   12
         Top             =   1050
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSalLoan 
         Height          =   315
         Left            =   4800
         TabIndex        =   14
         Top             =   1830
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCalLoan 
         Height          =   315
         Left            =   4800
         TabIndex        =   15
         Top             =   2220
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtAbsent 
         Height          =   315
         Left            =   7080
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTAX 
         Height          =   315
         Left            =   4800
         TabIndex        =   13
         Top             =   1440
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFrom 
         Height          =   315
         Left            =   540
         TabIndex        =   0
         Top             =   90
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
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
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTo 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   90
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
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
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoliday 
         Height          =   315
         Left            =   1590
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCommission 
         Height          =   315
         Left            =   1590
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtUndertime 
         Height          =   315
         Left            =   7080
         TabIndex        =   20
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTelBill 
         Height          =   315
         Left            =   7080
         TabIndex        =   22
         Top             =   1830
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtOthers 
         Height          =   315
         Left            =   7080
         TabIndex        =   23
         Top             =   2220
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNetPay 
         Height          =   315
         Left            =   7080
         TabIndex        =   25
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtGrossWage 
         Height          =   315
         Left            =   1590
         TabIndex        =   9
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTotDed 
         Height          =   315
         Left            =   7080
         TabIndex        =   24
         Top             =   2610
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDailyRate 
         Height          =   315
         Left            =   1590
         TabIndex        =   3
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTaxableAdj 
         Height          =   315
         Left            =   1590
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtPagMPLoan 
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Top             =   2610
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPagHLLoan 
         Height          =   315
         Left            =   4800
         TabIndex        =   17
         Top             =   3000
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNonTaxableAdj 
         Height          =   315
         Left            =   1590
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox txtSalaryAdvance 
         Height          =   315
         Left            =   7080
         TabIndex        =   19
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBLLoan 
         Height          =   315
         Left            =   7080
         TabIndex        =   18
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label46 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
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
         Height          =   255
         Left            =   6390
         TabIndex        =   218
         Top             =   2220
         Width           =   1755
      End
      Begin VB.Label Label45 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. Bill"
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
         Height          =   255
         Left            =   6330
         TabIndex        =   217
         Top             =   1860
         Width           =   1815
      End
      Begin VB.Label Label44 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "UT/Late"
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
         Height          =   255
         Left            =   6300
         TabIndex        =   216
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   3120
         X2              =   3120
         Y1              =   -30
         Y2              =   3420
      End
      Begin VB.Label Label42 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Commission"
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
         Height          =   255
         Left            =   390
         TabIndex        =   215
         Top             =   2670
         Width           =   1995
      End
      Begin VB.Label Label41 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday"
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
         Height          =   255
         Left            =   840
         TabIndex        =   214
         Top             =   2280
         Width           =   1875
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   3120
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label39 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Height          =   255
         Left            =   60
         TabIndex        =   213
         Top             =   90
         Width           =   1455
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   255
         Left            =   1650
         TabIndex        =   212
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TAX"
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
         Height          =   255
         Left            =   4140
         TabIndex        =   211
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Absent"
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
         Height          =   255
         Left            =   6360
         TabIndex        =   210
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS Cal. Loan"
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
         Height          =   255
         Left            =   3210
         TabIndex        =   209
         Top             =   2250
         Width           =   1605
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS Sal. Loan"
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
         Height          =   255
         Left            =   3210
         TabIndex        =   208
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Pag-Ibig"
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
         Height          =   255
         Left            =   3810
         TabIndex        =   207
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label32 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PHIC"
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
         Height          =   255
         Left            =   4080
         TabIndex        =   206
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS"
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
         Height          =   255
         Left            =   4110
         TabIndex        =   205
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime"
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
         Height          =   255
         Left            =   690
         TabIndex        =   204
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label Label29 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Height          =   255
         Left            =   1110
         TabIndex        =   203
         Top             =   510
         Width           =   1455
      End
      Begin VB.Label Label47 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Pay"
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
         Height          =   255
         Left            =   6270
         TabIndex        =   202
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label48 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Wage"
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
         Height          =   255
         Left            =   390
         TabIndex        =   201
         Top             =   3000
         Width           =   1995
      End
      Begin VB.Label Label49 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ded."
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
         Height          =   255
         Left            =   6120
         TabIndex        =   200
         Top             =   2640
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Rate"
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
         Height          =   255
         Left            =   600
         TabIndex        =   199
         Top             =   870
         Width           =   1875
      End
      Begin VB.Label Label40 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable Adj."
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
         Height          =   255
         Left            =   390
         TabIndex        =   198
         Top             =   1230
         Width           =   1875
      End
      Begin VB.Label Label50 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PI MP Loan"
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
         Height          =   255
         Left            =   3210
         TabIndex        =   197
         Top             =   2640
         Width           =   1755
      End
      Begin VB.Label Label51 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PI Housing Loan"
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
         Height          =   255
         Left            =   3210
         TabIndex        =   196
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label52 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Non-Tax. Adj."
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
         Height          =   255
         Left            =   270
         TabIndex        =   195
         Top             =   1590
         Width           =   2265
      End
      Begin VB.Label Label75 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sal. Adv."
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
         Height          =   255
         Left            =   6210
         TabIndex        =   194
         Top             =   690
         Width           =   1845
      End
      Begin VB.Label Label76 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "BL Loan"
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
         Height          =   255
         Left            =   6270
         TabIndex        =   193
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00DEDFDE&
      Height          =   1155
      Left            =   9900
      ScaleHeight     =   1095
      ScaleWidth      =   1275
      TabIndex        =   149
      Top             =   360
      Width           =   1335
      Begin VB.Image imgDispPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Left            =   30
         Picture         =   "Ledger.frx":2B415
         Stretch         =   -1  'True
         Top             =   30
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   2640
      Picture         =   "Ledger.frx":3F172
      ScaleHeight     =   1125
      ScaleWidth      =   7185
      TabIndex        =   143
      Top             =   360
      Width           =   7215
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1110
         TabIndex        =   146
         Top             =   150
         Width           =   6015
      End
      Begin VB.TextBox txtPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1110
         TabIndex        =   145
         Top             =   600
         Width           =   6015
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   148
         Top             =   150
         Width           =   885
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   147
         Top             =   600
         Width           =   885
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
         Height          =   255
         Left            =   2160
         TabIndex        =   144
         Top             =   270
         Width           =   615
      End
   End
   Begin wizMacForm.wizMacApp wizMacApp1 
      Height          =   320
      Left            =   0
      TabIndex        =   142
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   556
      MacCaption      =   "Mac Caption"
      Object.ToolTipText     =   "MAC titlebars can even have tooltips"
   End
   Begin wizButton.cmd cmdPayroll 
      Height          =   3615
      Left            =   2700
      TabIndex        =   126
      Top             =   2250
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   6376
      TX              =   "cmd1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Haettenschweiler"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Ledger.frx":4142B
   End
   Begin VB.PictureBox Picture11 
      Height          =   6525
      Left            =   60
      Picture         =   "Ledger.frx":41447
      ScaleHeight     =   6465
      ScaleWidth      =   2445
      TabIndex        =   188
      Top             =   330
      Width           =   2505
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   60
      Picture         =   "Ledger.frx":551A4
      ScaleHeight     =   6465
      ScaleWidth      =   2475
      TabIndex        =   189
      Top             =   360
      Width           =   2505
      Begin VB.TextBox txtSearch 
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
         Left            =   30
         MaxLength       =   35
         TabIndex        =   190
         Text            =   "Text1"
         Top             =   60
         Width           =   2415
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   5985
         Left            =   0
         TabIndex        =   191
         Top             =   450
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   10557
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
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
         MouseIcon       =   "Ledger.frx":57EE0
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
         Picture         =   "Ledger.frx":58042
      End
   End
   Begin VB.CommandButton cmdEditYTD 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit YTD"
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
      Left            =   2850
      Picture         =   "Ledger.frx":6BDAF
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   4110
      Width           =   1185
   End
   Begin VB.PictureBox picPayroll2 
      Height          =   405
      Left            =   6330
      ScaleHeight     =   345
      ScaleWidth      =   1440
      TabIndex        =   88
      Top             =   4080
      Width           =   1500
      Begin VB.CommandButton cmdCancelPayroll 
         Caption         =   "&Cancel"
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
         Left            =   720
         TabIndex        =   89
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSavePayroll 
         Caption         =   "&Save"
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
         Left            =   0
         TabIndex        =   90
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2640
      Picture         =   "Ledger.frx":6C1F1
      ScaleHeight     =   855
      ScaleWidth      =   8895
      TabIndex        =   80
      Top             =   5970
      Width           =   8925
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00DEDFDE&
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
         Left            =   7680
         MouseIcon       =   "Ledger.frx":6E4AA
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":6E5FC
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00DEDFDE&
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
         Left            =   6600
         MouseIcon       =   "Ledger.frx":6E906
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":6EA58
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00DEDFDE&
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
         Left            =   5520
         MouseIcon       =   "Ledger.frx":6F322
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":6F474
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00DEDFDE&
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
         Left            =   4440
         MouseIcon       =   "Ledger.frx":6FD3E
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":6FE90
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00DEDFDE&
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
         Left            =   3360
         MouseIcon       =   "Ledger.frx":7075A
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":708AC
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00DEDFDE&
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
         Left            =   2280
         MouseIcon       =   "Ledger.frx":71176
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":712C8
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00DEDFDE&
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
         Left            =   1200
         MouseIcon       =   "Ledger.frx":71B92
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":71CE4
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00DEDFDE&
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
         Left            =   120
         MouseIcon       =   "Ledger.frx":72126
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":72278
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   30
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2640
      Picture         =   "Ledger.frx":726BA
      ScaleHeight     =   855
      ScaleWidth      =   8895
      TabIndex        =   87
      Top             =   5970
      Width           =   8925
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00DEDFDE&
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
         Left            =   7680
         MouseIcon       =   "Ledger.frx":74973
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":74AC5
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DEDFDE&
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
         Left            =   6600
         MouseIcon       =   "Ledger.frx":75B07
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":75C59
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   30
         Width           =   1095
      End
   End
   Begin VB.Label labPayrollID 
      Caption         =   "TAX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   81
      Top             =   450
      Width           =   1455
   End
   Begin VB.Label LabID 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3990
      TabIndex        =   79
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "frmHRMSLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo, rsPayroll, rsDeductions As ADODB.Recordset
Dim rsLoanMas, rsCalamityLoan, rsLoanmasDet As ADODB.Recordset
Dim rsSSS, rsSSSdet, rsPH, rsPHDet As ADODB.Recordset
Dim rsPagIbig, rsTIN, rsPagibigdet As ADODB.Recordset
Dim rsTINdet, rsYTDDetails, rsSalaryGrade As ADODB.Recordset
Dim AddorEdit As String
Dim ToBeVat As Double
Dim IMPNO, CLID As String

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
EnAbleFrames
If TabSSS.Tab = 0 Then cmdAddPayroll.Value = True
If TabSSS.Tab = 1 Then
   EnAbleTab TabSSS, 1, 3
   InitLoans
End If
If TabSSS.Tab = 2 Then
   EnAbleTab TabSSS, 2, 3
   On Error Resume Next
   txtSSSNo.SetFocus
End If
If TabSSS.Tab = 3 Then
   EnAbleTab TabSSS, 3, 3
   On Error Resume Next
   txtPagIbigNo.SetFocus
End If
Picture1.Visible = False
Picture2.Visible = True
End Sub

Private Sub cmdAddPayroll_Click()
AddorEdit = "ADD"
cmdPayroll.ZOrder 0
picAddEditPayroll.ZOrder 0
txtFrom.Enabled = True
txtTo.Enabled = True
MAHIRO
InitMemVars
End Sub

Private Sub cmdCancel_Click()
cmdLoanMas.ZOrder 1
fraLoanMas.ZOrder 1
cmdPayroll.ZOrder 1
picAddEditPayroll.ZOrder 1
cmdCancelPayroll.Value = True
DisAbleFrames
EnAbleAll TabSSS, 1, 3
Picture1.Visible = True
Picture2.Visible = False
grdStore
End Sub

Private Sub cmdCancelPayroll_Click()
AddorEdit = ""
picAddEditPayroll.Visible = False
cmdPayroll.Visible = False
picAddEditPayroll.Enabled = False
Picture1.Enabled = True
TabSSS.TabEnabled(0) = True
TabSSS.TabEnabled(1) = True
TabSSS.TabEnabled(2) = True
TabSSS.TabEnabled(3) = True
TabSSS.TabEnabled(4) = True
picPayroll2.Visible = False
End Sub

Private Sub cmdDelete_Click()
cmdDeletePayroll.Value = True
End Sub

Private Sub cmdDeletePayroll_Click()
grdPayroll.Col = 25
If grdPayroll.Text <> "" Then
   If ShowConfirmDelete = True Then
      Set rsPayroll = New ADODB.Recordset
          rsPayroll.Open "select * from payroll where id = " & grdPayroll.Text, gconHRMS
      If Not rsPayroll.EOF And Not rsPayroll.BOF Then
         GENTO = rsPayroll!paydateto
         GENFROM = rsPayroll!paydatefrom
         IMPNO = IMPNO
         DelEXIST
         ShowDeletedMsg
      End If
   End If
Else
   ShowNothingToDeleteMsg
End If
StoreMemvars
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
EnAbleFrames
Picture1.Visible = False
Picture2.Visible = True
If TabSSS.Tab = 0 Then cmdEditPayroll.Value = True
If TabSSS.Tab = 4 Then cmdEditYTD.Value = True
End Sub

Private Sub cmdEditPayroll_Click()
Dim fild As String
TabSSS.TabEnabled(0) = False
TabSSS.TabEnabled(1) = False
TabSSS.TabEnabled(2) = False
TabSSS.TabEnabled(3) = False
TabSSS.TabEnabled(4) = False
grdPayroll.Row = grdPayroll.Row
grdPayroll.Col = 25
fild = grdPayroll.Text
If fild <> "" Then MAIDIT (fild)
End Sub

Function MAIDIT(XXX As String)
AddorEdit = "EDIT"
cmdPayroll.ZOrder 0
picAddEditPayroll.ZOrder 0
txtFrom.Enabled = True
txtTo.Enabled = True
MAHIRO
StoreEntry (XXX)
End Function

Sub MAHIRO()
EnAbleFrames
TabSSS.TabEnabled(0) = False
Picture1.Visible = False
Picture2.Visible = True
picAddEditPayroll.Visible = True
cmdPayroll.Visible = True
picAddEditPayroll.Enabled = True
picPayroll2.Visible = True
End Sub

Function StoreEntry(ByVal ID As Variant)
Set rsPayroll = New ADODB.Recordset
    rsPayroll.Open "select * from payroll where id = " & ID, gconHRMS, adOpenForwardOnly, adLockReadOnly
If Not rsPayroll.EOF And Not rsPayroll.BOF Then
   labPayrollID.Caption = rsPayroll!ID
   txtFrom.Text = Null2Date(rsPayroll!paydatefrom)
   txtTo.Text = Null2Date(rsPayroll!paydateto)
   txtRate.Text = N2Str2Zero(rsPayroll!Rate)
   txtDailyRate.Text = N2Str2Zero(rsPayroll!DailyRate)
   txtTaxableAdj.Text = N2Str2IntZero(rsPayroll!taxableadj)
   txtNonTaxableAdj.Text = N2Str2IntZero(rsPayroll!nontaxableadj)
   txtOvertime.Text = N2Str2Zero(rsPayroll!overtime)
   txtHoliday.Text = N2Str2Zero(rsPayroll!holiday)
   txtCommission.Text = N2Str2Zero(rsPayroll!commission)
   txtSSS.Text = N2Str2Zero(rsPayroll!sssE)
   txtMed.Text = N2Str2Zero(rsPayroll!philhealthE)
   txtPagIbig.Text = N2Str2Zero(rsPayroll!pagibig)
   txtTAX.Text = N2Str2Zero(rsPayroll!tax)
   txtSalLoan.Text = N2Str2Zero(rsPayroll!ssssalloan)
   txtCalLoan.Text = N2Str2Zero(rsPayroll!ssscalloan)
   txtPagMPLoan.Text = N2Str2Zero(rsPayroll!pagmploan)
   txtPagHLLoan.Text = N2Str2Zero(rsPayroll!paghlloan)
   txtBLLoan.Text = N2Str2Zero(rsPayroll!blloan)
   txtSalaryAdvance.Text = N2Str2Zero(rsPayroll!SalaryAdvance)
   txtUndertime.Text = N2Str2Zero(rsPayroll!undertime)
   txtAbsent.Text = N2Str2Zero(rsPayroll!absent)
   txtTelBill.Text = N2Str2Zero(rsPayroll!telbill)
   txtOthers.Text = N2Str2Zero(rsPayroll!others)
   showGrossWage
End If
End Function

Sub showGrossWage()
Dim GrossWeyg As Double
GrossWeyg = NumericVal(txtRate.Text) + NumericVal(txtTaxableAdj.Text) + NumericVal(txtOvertime.Text) + NumericVal(txtHoliday.Text)
txtGrossWage.Text = NumericVal(txtRate.Text) + NumericVal(txtTaxableAdj.Text) + NumericVal(txtNonTaxableAdj.Text) + NumericVal(txtOvertime.Text) + NumericVal(txtHoliday.Text)
If wizVar.EncryptAccess(LOGNAME) = "77697A∂ùº∞û¡¡¢¡" Then
   txtTAX.Text = TaxDedSemiMonthly(GrossWeyg, Null2String(rsEmpInfo!exstatus))
   txtSSS.Text = EmployeeSSSshare(NumericVal(txtRate.Text) * 2) / 2
   txtMed.Text = PhilHealthShare(NumericVal(txtRate.Text) * 2) / 2
   txtPagIbig.Text = PagIbigShare(NumericVal(txtRate.Text) * 2) / 2
End If
If AddorEdit = "ADD" Then
   txtTAX.Text = TaxDedSemiMonthly(GrossWeyg, Null2String(rsEmpInfo!exstatus))
   txtSSS.Text = EmployeeSSSshare(NumericVal(txtRate.Text) * 2) / 2
   txtMed.Text = PhilHealthShare(NumericVal(txtRate.Text) * 2) / 2
   txtPagIbig.Text = PagIbigShare(NumericVal(txtRate.Text) * 2) / 2
End If
ShowTotDed
ShowNetPay
End Sub

Sub ShowTotDed()
txtTotDed.Text = NumericVal(txtSSS.Text) + NumericVal(txtMed.Text) + NumericVal(txtPagIbig.Text) + NumericVal(txtTAX.Text) + NumericVal(txtSalLoan.Text) + NumericVal(txtCalLoan.Text) + NumericVal(txtPagMPLoan.Text) + NumericVal(txtPagHLLoan.Text) + NumericVal(txtBLLoan.Text) + NumericVal(txtSalaryAdvance.Text) + NumericVal(txtUndertime.Text) + NumericVal(txtAbsent.Text) + NumericVal(txtTelBill.Text) + NumericVal(txtOthers.Text)
ShowNetPay
End Sub

Sub ShowNetPay()
txtNetPay.Text = NumericVal(txtGrossWage.Text) - NumericVal(txtTotDed.Text)
End Sub

Private Sub cmdEditPrevYTD_Click()
picPrevYTD.Enabled = True
cmdEditPrevYTD.Enabled = False
cmdUpdatePrevYTD.Enabled = True
On Error Resume Next
txtPYTDGross.SetFocus
End Sub

Private Sub cmdEditYTD_Click()
TabSSS.Tab = 4
picYTD.Enabled = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
rsRefresh
picSearch.ZOrder 0
txtSearch.SetFocus
'rsRefresh
'on error resume next
'rsPayroll.Find "id = " & LabID.Caption
'Dim findStr As String
'findStr = InputSpeechBox("Please Input Name ...", txtName.Text)
'If findStr <> "" Then
'   on error resume next
'   rsEmpinfo.Bookmark = rsFind(rsEmpinfo.Clone, "lastname", findStr).Bookmark
'   If Err.Number = 3021 Then
'      On Error GoTo ErrorCode
'      rsEmpinfo.Bookmark = rsFind(rsEmpinfo.Clone, "firstname", findStr).Bookmark
'   End If
'End If
'StoreMemvars
'Exit Sub

'ErrorCode:
'If Err.Number = 3021 Then
'   ShowCantFind findStr
'   Resume Next
'End If
End Sub

Private Sub cmdNext_Click()
rsEmpInfo.MoveNext
If rsEmpInfo.EOF Then
   rsEmpInfo.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsEmpInfo.MovePrevious
If rsEmpInfo.BOF Then
   rsEmpInfo.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrint_Click()
Dim Filter As String
If TabSSS.Tab = 0 Then cmdPrintPayroll.Value = True
If TabSSS.Tab = 4 Then
   Screen.MousePointer = 11
   rptPayroll.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
   rptPayroll.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
   rptPayroll.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
   PrintSQLReport rptPayroll, HRMS_REPORT_PATH & "finalpay.rpt", "{ytddetails.empno} = " & N2Str2Null(IMPNO) & " and {ytddetails.yeer} = '" & Year(LOGDATE) & "'", HRMS_REPORT_Connection, 1
   Screen.MousePointer = 0
End If
End Sub

Private Sub cmdPrintPayroll_Click()
Dim Filter As String
Screen.MousePointer = 11
Dim CLID As String
grdPayroll.Row = grdPayroll.Row
grdPayroll.Col = 13
CLID = grdPayroll.Text
If CLID <> "" Then
   rptPayroll.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
   rptPayroll.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
   rptPayroll.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
   If wizVar.EncryptAccess(LOGNAME) = "77697A∂ùº∞û¡¡¢¡" Then
      If MsgQuestionBox("With PaySlip?", "Individual") = True Then
         PrintSQLReport rptPayroll, HRMS_REPORT_PATH & "indivpayslip.rpt", "{payroll.empno} = " & N2Str2Null(IMPNO), HRMS_REPORT_Connection, 1
      End If
   End If
   PrintSQLReport rptPayroll, HRMS_REPORT_PATH & "ledger.rpt", "{payroll.empno} = " & N2Str2Null(IMPNO), HRMS_REPORT_Connection, 1
End If
Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
If TabSSS.Tab = 0 Then cmdSavePayroll.Value = True
If TabSSS.Tab = 1 Then
   Dim LAWNType As String
   If cboLoanType.Text = "SSS Salary Loan" Then LAWNType = "'SSAL'"
   If cboLoanType.Text = "SSS Calamity Loan" Then LAWNType = "'CSAL'"
   If cboLoanType.Text = "Pag-Ibig Salary Loan" Then LAWNType = "'PSAL'"
   If cboLoanType.Text = "Pag-Ibig MP Loan" Then LAWNType = "'MPL'"
   If cboLoanType.Text = "Pag-Ibig Housing Loan" Then LAWNType = "'HLL'"
   If cboLoanType.Text = "BL Loan" Then LAWNType = "'BLL'"
   Dim VtxtAcctNo, VtxtDateGranted, VtxtDateStarted, VtxtMaturityDate As String
   Dim VtxtAmountLoaned, VtxtMonthlyDed, VtxtSMonthlyDed, VtxtLoanBalance As Double
   Dim VtxtOtherTypeDed, VtxtDeduction_Option As String
   Dim VtxtOtherTypeDedAmount As Double
   Dim LonID As Integer
   VtxtAcctNo = N2Str2Null(txtAcctNo.Text)
   VtxtDateGranted = N2Date2Null(txtDateGranted.Text)
   VtxtDateStarted = N2Date2Null(txtDateStarted.Text)
   VtxtMaturityDate = N2Date2Null(txtMaturityDate.Text)
   VtxtAmountLoaned = NumericVal(txtAmountLoaned.Text)
   VtxtMonthlyDed = NumericVal(txtMonthlyDed.Text)
   VtxtSMonthlyDed = NumericVal(txtSMonthlyDed.Text)
   VtxtLoanBalance = NumericVal(txtLoanBalance.Text)
   VtxtOtherTypeDed = N2Str2Null(txtOtherTypeDed.Text)
   If Opt1.Value = True Then VtxtDeduction_Option = "'OPT15thOnly'"
   If Opt2.Value = True Then VtxtDeduction_Option = "'OPT30thOnly'"
   If Opt3.Value = True Then VtxtDeduction_Option = "'OPT15th30th'"
   VtxtOtherTypeDedAmount = NumericVal(txtOtherTypeDedAmount.Text)
   If AddorEdit = "ADD" Then
      gconHRMS.Execute "insert into loanmas " & _
                       "(loantype,empno,acctno,dategranted,dateStarted,maturitydate,amountloaned,monthlyded,smonthlyded,loanbalance,Deduction_Option,OtherTypeDed,OtherTypeDedAmount)" & _
                       " values (" & LAWNType & ", '" & IMPNO & "', " & VtxtAcctNo & ", " & VtxtDateGranted & ", " & VtxtDateStarted & _
                       ", " & VtxtMaturityDate & ", " & VtxtAmountLoaned & ", " & VtxtMonthlyDed & ", " & VtxtSMonthlyDed & ", " & VtxtLoanBalance & ", " & VtxtDeduction_Option & ", " & VtxtOtherTypeDed & ", " & VtxtOtherTypeDedAmount & ")"
      ShowSuccessFullyAdded
   Else
      grdLoanMas.Col = 8
      LonID = grdLoanMas.Text
      gconHRMS.Execute "update loanmas set" & _
                       " loantype = " & LAWNType & "," & _
                       " acctno = " & VtxtAcctNo & "," & _
                       " dategranted = " & VtxtDateGranted & "," & _
                       " dateStarted = " & VtxtDateStarted & "," & _
                       " maturitydate = " & VtxtMaturityDate & "," & _
                       " amountloaned = " & VtxtAmountLoaned & "," & _
                       " monthlyded = " & VtxtMonthlyDed & "," & _
                       " smonthlyded = " & VtxtSMonthlyDed & "," & _
                       " loanbalance = " & VtxtLoanBalance & "," & _
                       " Deduction_Option = " & VtxtDeduction_Option & "," & _
                       " OtherTypeDed = " & VtxtOtherTypeDed & "," & _
                       " OtherTypeDedAmount = " & VtxtOtherTypeDedAmount & _
                       " where id = " & LonID
      ShowSuccessFullyUpdated
   End If
   grdStore
End If
If TabSSS.Tab = 2 Then
   If txtSSSNo.Text <> "" Then
      gconHRMS.Execute "update sss set" & _
                       " employeeshare = " & NumericVal(txtSSSMonthly.Text) & "," & _
                       " datestart = '" & txtSSSStarted.Text & "'," & _
                       " lastdatecont = '" & txtSSSLast.Text & "'" & _
                       " where sssno = " & N2Str2Null(rsEmpInfo!sssno)
      ShowSuccessFullyUpdated
   End If
   If txtPhilHealthNo.Text <> "" Then
      gconHRMS.Execute "update philhealth set" & _
                       " employeeshare = " & NumericVal(txtPHMonthly.Text) & "," & _
                       " datestart = '" & txtPHStarted.Text & "'," & _
                       " lastdatecont = '" & txtPHLast.Text & "'" & _
                       " where phno = " & N2Str2Null(rsEmpInfo!phno)
      ShowSuccessFullyUpdated
   End If
   grdStore
End If
If TabSSS.Tab = 3 Then
   If txtPagIbigNo.Text <> "" Then
      gconHRMS.Execute "update pagibig set" & _
                       " employeeshare = " & NumericVal(txtPagIbigMonthly.Text) & "," & _
                       " datestart = '" & txtPagIbigStarted.Text & "'," & _
                       " lastdatecont = '" & txtPagIbigLast.Text & "'" & _
                       " where pagibigno = " & N2Str2Null(rsEmpInfo!pagibigno)
      ShowSuccessFullyUpdated
   End If
   If txtTINNo.Text <> "" Then
      gconHRMS.Execute "update tin set" & _
                       " deduction = " & NumericVal(txtTINMonthly.Text) & "," & _
                       " datestart = '" & txtTINStarted.Text & "'," & _
                       " lastdatecont = '" & txtTINLast.Text & "'" & _
                       " where tinno = " & N2Str2Null(rsEmpInfo!tinno)
      ShowSuccessFullyUpdated
   End If
   grdStore
End If
If TabSSS.Tab = 4 Then
   gconHRMS.Execute "update ytddetails set" & _
                    " midyear = " & NumericVal(txtMidYear.Text) & "," & _
                    " remsal = " & NumericVal(txtYTDRemSal.Text) & "," & _
                    " remot = " & NumericVal(txtYTDRemOT.Text) & "," & _
                    " remwtax = " & NumericVal(txtYTDRemWTax.Text) & "," & _
                    " remDed = " & NumericVal(txtYTDRemDed.Text) & "," & _
                    " ytdincome = " & NumericVal(txtYTDGross.Text) & "," & _
                    " personalex = " & NumericVal(txtTaxExemp.Text) & "," & _
                    " t13thmonth = " & NumericVal(txt13thMonth.Text) & "," & _
                    " taxdue = " & NumericVal(txtYTDTaxDue.Text) & _
                    " where empno = " & N2Str2Null(rsEmpInfo!empno) & " AND yeer = '" & Year(LOGDATE) & "'"
   ShowSuccessFullyUpdated
End If
cmdCancel.Value = True
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Private Sub cmdSavePayroll_Click()
Dim net, gruss As Double
GENFROM = Format(txtFrom.Text, "Short Date")
GENTO = Format(txtTo.Text, "Short Date")
If IsDate(GENFROM) = False Then
   MsgSpeechBox "Error in From Date"
   Exit Sub
End If
If IsDate(GENTO) = False Then
   MsgSpeechBox "Error in To Date"
   Exit Sub
End If
IMPNO = IMPNO
DelEXIST
SEYV
End Sub

Sub DelEXIST()
gconHRMS.Execute "delete * from payroll where (paydatefrom >= #" & GENFROM & "#)" & _
                 " AND (paydateto <= #" & GENTO & "#) and empno = '" & IMPNO & "'"
gconHRMS.Execute "delete * from SSSdet where (deyt = #" & GENTO & "#) and empno = '" & IMPNO & "'"
gconHRMS.Execute "delete * from LoanMasDet where (deyt = #" & GENTO & "#) and empno = '" & IMPNO & "'"
gconHRMS.Execute "delete * from philhealthdet where (deyt = #" & GENTO & "#) and empno = '" & IMPNO & "'"
gconHRMS.Execute "delete * from pagibigdet where (deyt = #" & GENTO & "#) and empno = '" & IMPNO & "'"
gconHRMS.Execute "delete * from tindet where (deyt = #" & GENTO & "#) and empno = '" & IMPNO & "'"
gconHRMS.Execute "delete * from atmdet where (deyt = #" & GENTO & "#) and empno = '" & IMPNO & "'"
End Sub

Sub SEYV()
On Error GoTo ErrorCode
Dim rsPrevPayroll, rsAllPrevPayroll, rsCommission As ADODB.Recordset
Dim rsEmpInfoClone As ADODB.Recordset
Dim SUWELDO, SUWELDOKINSE As Double
Dim dedPAGIBIG, dedEmpPAGIBIG As Double
Dim dedTIN, dedSSS As Double
Dim dedEmpSSS, dedPhilHealth As Double
Dim dedEmpPhilhealth, dedSalLoan As Double
Dim dedCalLoan As Double
Dim dedMPLoan, dedHLLoan, dedBLLoan, SalGross As Double
Dim TotSalaryAdvance, TotUndertime, TotAbsent As Double
Dim TotTelBill, TotOthers As Double
Dim TotOvertime, TotTaxableAdj As Double
Dim TotNonTaxableAdj, TotHoliday As Double
Dim TotCommission, TotCommissionTax As Double
Dim i, Cnt As Integer
Dim amt, NUMDAYS, DEYLI, BULANAN, SUWELDOTRIENTA As Double
Dim VARPAYSTATUS As String
VARPAYSTATUS = "U"
Screen.MousePointer = 11
SUWELDOKINSE = NumericVal(txtRate.Text)
DEYLI = NumericVal(txtDailyRate.Text)
TotOvertime = NumericVal(txtOvertime.Text)
TotTaxableAdj = NumericVal(txtTaxableAdj.Text)
TotNonTaxableAdj = NumericVal(txtNonTaxableAdj.Text)
TotHoliday = NumericVal(txtHoliday.Text)
TotCommission = NumericVal(txtCommission.Text)
dedSSS = NumericVal(txtSSS.Text)
dedPhilHealth = NumericVal(txtMed.Text)
dedSalLoan = NumericVal(txtSalLoan.Text)
dedCalLoan = NumericVal(txtCalLoan.Text)
dedMPLoan = NumericVal(txtPagMPLoan.Text)
dedHLLoan = NumericVal(txtPagHLLoan.Text)
dedBLLoan = NumericVal(txtBLLoan.Text)
dedPAGIBIG = NumericVal(txtPagIbig.Text)
dedTIN = NumericVal(txtTAX.Text)
TotSalaryAdvance = NumericVal(txtSalaryAdvance.Text)
TotUndertime = NumericVal(txtUndertime.Text)
TotAbsent = NumericVal(txtAbsent.Text)
TotTelBill = NumericVal(txtTelBill.Text)
TotOthers = NumericVal(txtOthers.Text)
SalGross = NumericVal(txtGrossWage.Text)
dedEmpPhilhealth = 0
dedEmpPAGIBIG = 0
dedEmpSSS = 0
Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select id,empno,empstatus,salarycode,exstatus,accountno,sssno,tinno from empinfo where empno = " & N2Str2Null(IMPNO), gconHRMS
If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
   If Null2String(rsEmpInfo!empstatus) = "M" Then
      NUMDAYS = 0
      SUWELDOTRIENTA = SetSalary(Null2String(rsEmpInfo!salarycode))
   Else
      SUWELDOTRIENTA = SUWELDOKINSE
      NUMDAYS = SUWELDOKINSE / DEYLI
      If Day(GENFROM) > 15 Then
         Set rsPrevPayroll = New ADODB.Recordset
             rsPrevPayroll.Open "select empno,paydatefrom,rate from payroll where empno = " & N2Str2Null(rsEmpInfo!empno) & " and paydatefrom = #" & CDate(firstDay(GENFROM)) & "#", gconHRMS
         If Not rsPrevPayroll.EOF And Not rsPrevPayroll.BOF Then
            SUWELDOTRIENTA = SUWELDOKINSE + N2Str2Zero(rsPrevPayroll!Rate)
         End If
      End If
   End If
   Set rsCommission = New ADODB.Recordset
       rsCommission.Open "select * from Commission where empno = " & N2Str2Null(rsEmpInfo!empno) & " AND " & _
                         "(deyt >= #" & Format(GENFROM, "Short Date") & "#)" & _
                         " AND (deyt <= #" & Format(GENTO, "Short Date") & "#)", gconHRMS, adOpenForwardOnly, adLockReadOnly
   TotCommissionTax = 0
   If Not rsCommission.EOF And Not rsCommission.BOF Then
      rsCommission.MoveFirst
      Do While Not rsCommission.EOF
         TotCommissionTax = TotCommissionTax + N2Str2Zero(rsCommission!tax)
         rsCommission.MoveNext
      Loop
   End If
   Set rsLoanMas = New ADODB.Recordset
       rsLoanMas.Open "select * from loanmas where empno = " & N2Str2Null(rsEmpInfo!empno) & " AND (dategranted <= #" & Format(GENTO, "Short Date") & "#) AND (maturitydate >= #" & Format(GENTO, "Short Date") & "#) order by dategranted desc", gconHRMS
   If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
      rsLoanMas.MoveFirst
      Do While Not rsLoanMas.EOF
         If N2Str2Zero(rsLoanMas!LoanBalance) > 0 Then
            If Null2String(rsLoanMas!loantype) = "SSAL" Then
               gconHRMS.Execute "insert into LoanMasDet " & _
                                "(empno,acctno,amount,deyt,loantype)" & _
                                " values (" & N2Str2Null(rsLoanMas!empno) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedSalLoan) & ", '" & GENTO & "', 'SSAL')"
               gconHRMS.Execute "update loanmas set " & _
                                " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedSalLoan) & _
                                " where empno = " & N2Str2Null(rsLoanMas!empno)
            End If
            If Null2String(rsLoanMas!loantype) = "CSAL" Then
               gconHRMS.Execute "insert into LoanMasDet " & _
                                "(empno,acctno,amount,deyt,loantype)" & _
                                " values (" & N2Str2Null(rsLoanMas!empno) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedCalLoan) & ", '" & GENTO & "','CSAL')"
               gconHRMS.Execute "update loanmas set " & _
                                " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedCalLoan) & _
                                " where empno = " & N2Str2Null(rsLoanMas!empno)
            End If
            If Null2String(rsLoanMas!loantype) = "MPL" Then
               gconHRMS.Execute "insert into LoanMasDet " & _
                                "(empno,acctno,amount,deyt,loantype)" & _
                                " values (" & N2Str2Null(rsLoanMas!empno) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedMPLoan) & ", '" & GENTO & "','MPL')"
               gconHRMS.Execute "update loanmas set " & _
                                " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedMPLoan) & _
                                " where empno = " & N2Str2Null(rsLoanMas!empno)
            End If
            If Null2String(rsLoanMas!loantype) = "HLL" Then
               gconHRMS.Execute "insert into LoanMasDet " & _
                                "(empno,acctno,amount,deyt,loantype)" & _
                                " values (" & N2Str2Null(rsLoanMas!empno) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedHLLoan) & ", '" & GENTO & "','HLL')"
               gconHRMS.Execute "update loanmas set " & _
                                " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedHLLoan) & _
                                " where empno = " & N2Str2Null(rsLoanMas!empno)
            End If
            If Null2String(rsLoanMas!loantype) = "BLL" Then
               gconHRMS.Execute "insert into LoanMasDet " & _
                                "(empno,acctno,amount,deyt,loantype)" & _
                                " values (" & N2Str2Null(rsLoanMas!empno) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedBLLoan) & ", '" & GENTO & "','BLL')"
               gconHRMS.Execute "update loanmas set " & _
                                " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedBLLoan) & _
                                " where empno = " & N2Str2Null(rsLoanMas!empno)
            End If
         End If
         rsLoanMas.MoveNext
      Loop
   End If
   If Day(GENTO) > 15 Then
      dedEmpSSS = EmployerSSSshare(N2Str2Zero(SUWELDOTRIENTA))
      Set rsSSS = New ADODB.Recordset
          rsSSS.Open "select * from sss where empno = " & N2Str2Null(rsEmpInfo!empno), gconHRMS, adOpenForwardOnly, adLockReadOnly
      If rsSSS.EOF And rsSSS.BOF Then
         gconHRMS.Execute "insert into sss " & _
                          "(empno,sssno,datestart,employeeshare,employershare,lastdatecont)" & _
                          " values (" & N2Str2Null(rsEmpInfo!empno) & ", " & N2Str2Null(rsEmpInfo!sssno) & ", '" & GENTO & "'," & _
                          " " & dedSSS & ", " & dedEmpSSS & ", '" & GENTO & "')"
      Else
         gconHRMS.Execute "update sss set" & _
                          " employeeshare = " & dedSSS & "," & _
                          " employershare = " & dedEmpSSS & "," & _
                          " lastdatecont = '" & GENTO & "'" & _
                          " where empno = " & N2Str2Null(rsEmpInfo!empno)
      End If
      Set rsSSS = New ADODB.Recordset
          rsSSS.Open "select * from sss where empno = " & N2Str2Null(rsEmpInfo!empno), gconHRMS, adOpenForwardOnly, adLockReadOnly
      If Not rsSSS.EOF And Not rsSSS.BOF Then
         gconHRMS.Execute "insert into sssdet " & _
                          "(aydi,deyt,empno,employeeamount,employeramount)" & _
                          " values (" & rsSSS!aydi & ", '" & GENTO & "', " & N2Str2Null(rsEmpInfo!empno) & ", " & dedSSS & ", " & dedEmpSSS & ")"
      End If
      dedEmpPhilhealth = PhilHealthShare(N2Str2Zero(SUWELDOTRIENTA))
      Set rsPH = New ADODB.Recordset
          rsPH.Open "select * from philhealth where empno = " & N2Str2Null(rsEmpInfo!empno), gconHRMS, adOpenForwardOnly, adLockReadOnly
      If rsPH.EOF And rsPH.BOF Then
         gconHRMS.Execute "insert into philhealth " & _
                          "(empno,phno,datestart,employeeshare,employershare,lastdatecont)" & _
                          " values (" & N2Str2Null(rsEmpInfo!empno) & ", " & N2Str2Null(rsEmpInfo!sssno) & ", '" & GENTO & "'," & _
                          " " & dedPhilHealth & ", " & dedEmpPhilhealth & ", '" & GENTO & "')"
      Else
         gconHRMS.Execute "update philhealth set" & _
                          " employeeshare = " & dedPhilHealth & "," & _
                          " employershare = " & dedEmpPhilhealth & "," & _
                          " lastdatecont = '" & GENTO & "'" & _
                          " where empno = " & N2Str2Null(rsEmpInfo!empno)
      End If
      Set rsPH = New ADODB.Recordset
          rsPH.Open "select * from philhealth where empno = " & N2Str2Null(rsEmpInfo!empno), gconHRMS, adOpenForwardOnly, adLockReadOnly
      If Not rsPH.EOF And Not rsPH.BOF Then
         gconHRMS.Execute "insert into philhealthdet " & _
                          "(aydi,deyt,empno,employeeamount,employeramount)" & _
                          " values (" & rsPH!aydi & ", '" & GENTO & "', " & N2Str2Null(rsEmpInfo!empno) & ", " & dedPhilHealth & ", " & dedEmpPhilhealth & ")"
      End If
   Else
      Set rsPagIbig = New ADODB.Recordset
          rsPagIbig.Open "select * from Pagibig where empno = " & N2Str2Null(rsEmpInfo!empno), gconHRMS, adOpenForwardOnly, adLockReadOnly
      If rsPagIbig.EOF And rsPagIbig.BOF Then
         dedEmpPAGIBIG = (PagIbigShare(SetSalary(Null2String(rsEmpInfo!salarycode))))
         gconHRMS.Execute "insert into PagIbig " & _
                          "(empno,pagibigno,datestart,employeeshare,employershare,lastdatecont)" & _
                          " values (" & N2Str2Null(rsEmpInfo!empno) & ", " & N2Str2Null(rsEmpInfo!sssno) & ", '" & GENTO & "'," & _
                          " " & dedPAGIBIG & ", " & dedEmpPAGIBIG & ", '" & GENTO & "')"
      Else
         dedEmpPAGIBIG = (PagIbigShare(SetSalary(Null2String(rsEmpInfo!salarycode))))
         gconHRMS.Execute "update pagibig set" & _
                          " employeeshare = " & dedPAGIBIG & "," & _
                          " employershare = " & dedEmpPAGIBIG & "," & _
                          " lastdatecont = '" & GENTO & "'" & _
                          " where empno = " & N2Str2Null(rsEmpInfo!empno)
         gconHRMS.Execute "insert into pagibigdet " & _
                          "(aydi,deyt,empno,employeeamount,employeramount)" & _
                          " values (" & rsPagIbig!aydi & ", '" & GENTO & "', " & N2Str2Null(rsEmpInfo!empno) & ", " & dedPAGIBIG & ", " & dedEmpPAGIBIG & ")"
      End If
   End If
   Set rsTIN = New ADODB.Recordset
       rsTIN.Open "select * from TIN where empno = " & N2Str2Null(rsEmpInfo!empno), gconHRMS, adOpenForwardOnly, adLockReadOnly
   If rsTIN.EOF And rsTIN.BOF Then
      gconHRMS.Execute "insert into TIN " & _
                       "(empno,tinno,datestart,deduction,lastdatecont)" & _
                       " values (" & N2Str2Null(rsEmpInfo!empno) & ", " & N2Str2Null(rsEmpInfo!tinno) & ", '" & GENTO & "'," & _
                       " " & dedTIN & ", '" & GENTO & "')"
   Else
      gconHRMS.Execute "update TIN set" & _
                       " deduction = " & dedTIN & "," & _
                       " lastdatecont = '" & GENTO & "'" & _
                       " where empno = " & N2Str2Null(rsEmpInfo!empno)
      gconHRMS.Execute "insert into tindet " & _
                       "(aydi,empno,deyt,amount)" & _
                       " values (" & rsTIN!aydi & ", " & N2Str2Null(rsEmpInfo!empno) & ", '" & GENTO & "', " & dedTIN & ")"
   End If
   SUWELDO = (SUWELDOKINSE + TotOvertime + TotHoliday + TotTaxableAdj + TotNonTaxableAdj) - (dedPhilHealth + dedSSS + dedPAGIBIG + dedTIN + dedSalLoan + dedCalLoan + dedBLLoan + dedHLLoan + dedMPLoan + TotSalaryAdvance + TotUndertime + TotTelBill + TotAbsent + TotOthers)
   gconHRMS.Execute "insert into payroll " & _
                    "(empno,taxcode,rate,monthlyrate,DailyRate,ndays,overtime,holiday,commission,commissiontax,taxableadj,nontaxableadj,gross,undertime,sssE,sssR,philhealthE,philhealthR,pagibig,tax,ssssalloan,ssscalloan,pagMPloan,pagHLLoan,BLLoan,SalaryAdvance,absent,telbill,others,paydatefrom,paydateto,netpay,payrollstatus)" & _
                    " values (" & N2Str2Null(rsEmpInfo!empno) & ", " & N2Str2Null(rsEmpInfo!exstatus) & ", " & (SUWELDOKINSE) & _
                    ", " & N2Str2Zero(SUWELDOTRIENTA) & ", " & N2Str2Zero(DEYLI) & ", " & NUMDAYS & _
                    ", " & TotOvertime & ", " & TotHoliday & ", " & TotCommission & ", " & TotCommissionTax & ", " & TotTaxableAdj & ", " & TotNonTaxableAdj & _
                    ", " & (SUWELDOKINSE) + TotOvertime + TotHoliday + TotTaxableAdj + TotNonTaxableAdj & _
                    ", " & TotUndertime & ", " & dedSSS & ", " & dedEmpSSS & _
                    ", " & dedPhilHealth & ", " & dedEmpPhilhealth & ", " & dedPAGIBIG & ", " & dedTIN & ", " & dedSalLoan & ", " & dedCalLoan & ", " & dedMPLoan & ", " & dedHLLoan & ", " & dedBLLoan & _
                    ", " & TotSalaryAdvance & ", " & TotAbsent & ", " & TotTelBill & ", " & TotOthers & _
                    ", '" & GENFROM & "', '" & GENTO & "', " & SUWELDO & ", '" & VARPAYSTATUS & "')"
   gconHRMS.Execute "insert into atmdet " & _
                    "(acctno,empno,atmid,deyt,netamount) " & _
                    "values (" & N2Str2Null(rsEmpInfo!AccountNo) & ", " & N2Str2Null(rsEmpInfo!empno) & ", " & rsEmpInfo!ID & _
                    ", '" & GENTO & "', " & SUWELDO & ")"
End If
rsRefresh
On Error Resume Next
rsEmpInfo.Find "id = " & LabID.Caption
If EMPINFOSHOW = True Then
   frmHRMSEmpInfo.rsRefresh
   frmHRMSEmpInfo.StoreMemvars
End If
Screen.MousePointer = 0
Exit Sub

ErrorCode:
ShowVBError
Screen.MousePointer = 0
End Sub

Function SetSalary(SalCode As String) As Double
Set rsSalaryGrade = New ADODB.Recordset
    rsSalaryGrade.Open "select code,salary from salarygrade where code = '" & SalCode & "'", gconHRMS
If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
   SetSalary = N2Str2Zero(rsSalaryGrade!Salary)
End If
End Function

Function SetDailyRate(SalCode As String) As Double
Set rsSalaryGrade = New ADODB.Recordset
    rsSalaryGrade.Open "select code,dailyrate from salarygrade where code = '" & SalCode & "'", gconHRMS
If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
   SetDailyRate = N2Str2Zero(rsSalaryGrade!DailyRate)
End If
End Function

Private Sub cmdUpdatePrevYTD_Click()
picPrevYTD.Enabled = False
cmdEditPrevYTD.Enabled = True
cmdUpdatePrevYTD.Enabled = False
gconHRMS.Execute "update empinfo set " & _
                 " PreviousYTDGross = " & NumericVal(txtPYTDGross.Text) & "," & _
                 " PreviousYTDTax = " & NumericVal(txtPYTDTax.Text) & "," & _
                 " PreviousYTDSSS = " & NumericVal(txtPYTDSSS.Text) & "," & _
                 " PreviousYTDPHIC = " & NumericVal(txtPYTDPHIC.Text) & "," & _
                 " PreviousYTDPagIbig = " & NumericVal(txtPYTDPagIbig.Text) & "," & _
                 " PreviousYTDMidYear = " & NumericVal(txtPYTDMidYear.Text) & _
                 " where id = " & LabID.Caption
ShowSuccessFullyUpdated
rsRefresh
rsEmpInfo.Find "id = " & LabID.Caption
StoreMemvars
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyEscape
            grdLoanMasDet.ZOrder 1
       Case Else
            MoveKeyPress KeyCode
End Select
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
wizMacApp1.MacCaption = Me.Caption
wizMacApp1.Buttons = CloseMinimize
initGrid
rsRefresh
txtSearch.Text = ""
picAddEditPayroll.Visible = False
cmdPayroll.Visible = False
cboLoanType.Clear
cboLoanType.AddItem "SSS Salary Loan"
cboLoanType.AddItem "SSS Calamity Loan"
cboLoanType.AddItem "Pag-Ibig Salary Loan"
cboLoanType.AddItem "Pag-Ibig MP Loan"
cboLoanType.AddItem "Pag-Ibig Housing Loan"
cboLoanType.AddItem "BL Loan"
DisAbleFrames
StoreMemvars
DrawXPCtl Me
LEDGERSHOW = True
Screen.MousePointer = 0
End Sub

Sub DisAbleFrames()
fraLoanMas.Enabled = False
fraSSSMED.Enabled = False
fraPagIbigTIN.Enabled = False
picYTD.Enabled = False
End Sub

Sub EnAbleFrames()
fraLoanMas.Enabled = True
fraSSSMED.Enabled = True
fraPagIbigTIN.Enabled = True
End Sub

Sub rsRefresh()
If EMPINFOSHOW = True Then
   Set rsEmpInfo = New ADODB.Recordset
       rsEmpInfo.Open "select * from empinfo where empno = '" & EmpInfoEmpno.Caption & "'", gconHRMS, adOpenForwardOnly, adLockReadOnly
ElseIf HEADEMPINFOSHOW = True Then
   Set rsEmpInfo = New ADODB.Recordset
       rsEmpInfo.Open "select * from empinfo where empno = '" & frmHRMSEmpInfo.LabID.Caption & "'", gconHRMS, adOpenForwardOnly, adLockReadOnly
Else
   Set rsEmpInfo = New ADODB.Recordset
       rsEmpInfo.Open "select * from empinfo WHERE EMPLEVEL = '" & "E" & "' order by lastname,firstname,middlename asc", gconHRMS, adOpenForwardOnly, adLockReadOnly
End If
End Sub

Sub StoreMemvars()
On Error GoTo ErrorCode
If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
   Screen.MousePointer = 11
   DoEvents
   LabID.Caption = rsEmpInfo!ID
   txtName.Text = RTrim(rsEmpInfo!lastname) + ", " + RTrim(rsEmpInfo!firstname) + " " + RTrim(rsEmpInfo!middlename)
   IMPNO = Null2String(rsEmpInfo!empno)
   txtPosition.Text = Null2String(rsEmpInfo!Position)
   txtSSSNo.Text = Null2String(rsEmpInfo!sssno)
   txtPhilHealthNo.Text = Null2String(rsEmpInfo!phno)
   txtPagIbigNo.Text = Null2String(rsEmpInfo!pagibigno)
   txtTINNo.Text = Null2String(rsEmpInfo!tinno)
                 
   txtPYTDGross.Text = N2Str2Zero(rsEmpInfo!PreviousYTDGross)
   txtPYTDTax.Text = N2Str2Zero(rsEmpInfo!PreviousYTDTax)
   txtPYTDSSS.Text = N2Str2Zero(rsEmpInfo!PreviousYTDSSS)
   txtPYTDPHIC.Text = N2Str2Zero(rsEmpInfo!PreviousYTDPHIC)
   txtPYTDPagIbig.Text = N2Str2Zero(rsEmpInfo!PreviousYTDPagIbig)
   txtPYTDMidYear.Text = N2Str2Zero(rsEmpInfo!PreviousYTDMidYear)

   If Null2String(rsEmpInfo!picfilname) <> "" Then
      On Error Resume Next
      'LoadPic imgDispPic, HRMS_PICTURES_PATH & Null2String(rsEmpInfo!picfilname)
   Else
      LoadPic imgDispPic, ""
   End If
   Screen.MousePointer = 0
   grdStore
   StoreYTD
   Screen.MousePointer = 0
Else
   ShowNoRecord
   Unload Me
End If
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Sub StoreYTD()
On Error GoTo ErrorCode
Set rsYTDDetails = New ADODB.Recordset
    rsYTDDetails.Open "select * from ytddetails where empno = '" & rsEmpInfo!empno & "' and yeer = '" & Year(LOGDATE) & "'", gconHRMS, adOpenForwardOnly, adLockReadOnly
If Not rsYTDDetails.EOF And Not rsYTDDetails.BOF Then
   txtYTDBasicPay.Text = (N2Str2Zero(rsYTDDetails!ytdbasicpay) - N2Str2Zero(rsYTDDetails!overtime)) + N2Str2Zero(rsYTDDetails!nontaxable)
   txtYTDOvertime.Text = N2Str2Zero(rsYTDDetails!overtime)
   txtYTDCommission.Text = N2Str2Zero(rsYTDDetails!commission)
   txtMidYear.Text = N2Str2Zero(rsYTDDetails!midyear)
   txtYTDRemSal.Text = N2Str2Zero(rsYTDDetails!remsal)
   txtYTDRemOT.Text = N2Str2Zero(rsYTDDetails!remot)
   txtYTDRemWTax.Text = N2Str2Zero(rsYTDDetails!remwtax)
   txtYTDRemDed.Text = N2Str2Zero(rsYTDDetails!remded)
   txtYTDGross.Text = N2Str2Zero(rsYTDDetails!ytdincome)
   txtYTDSSSPAGIBIGPHIC.Text = N2Str2Zero(rsYTDDetails!nontaxable)
   txtTaxExemp.Text = N2Str2Zero(rsYTDDetails!PersonalEx)
   If N2Str2Zero(rsYTDDetails!Taxdue) <= 0 Then txtYTDTaxDue.Text = "Exempted" Else txtYTDTaxDue.Text = N2Str2Zero(rsYTDDetails!Taxdue)
   txtYTDTaxWithHeld.Text = N2Str2Zero(rsYTDDetails!ytdtax) + N2Str2Zero(rsYTDDetails!commissiontax) + N2Str2Zero(rsYTDDetails!decytdtax)
   txtYTDTaxRefund.Text = (N2Str2Zero(rsYTDDetails!ytdtax) + N2Str2Zero(rsYTDDetails!commissiontax) + N2Str2Zero(rsYTDDetails!decytdtax)) - N2Str2Zero(rsYTDDetails!Taxdue)
   txtTaxableIncome.Text = Null2String(rsYTDDetails!NetTaxable)
   txtAdjSalary.Text = Null2String(rsYTDDetails!remsal)
   txt13thMonth.Text = Null2String(rsYTDDetails!t13thmonth)
   txtTaxRefund.Text = (N2Str2Zero(rsYTDDetails!ytdtax) + N2Str2Zero(rsYTDDetails!commissiontax) + N2Str2Zero(rsYTDDetails!decytdtax)) - N2Str2Zero(rsYTDDetails!Taxdue)
   txtTotalPay.Text = (NumericVal(txtAdjSalary.Text) + NumericVal(txt13thMonth.Text) + NumericVal(txtTaxRefund.Text) + NumericVal(txtYTDRemOT.Text)) - NumericVal(txtYTDRemDed.Text)
Else
   InitYTDMemVars
End If
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Sub InitYTDMemVars()
txtYTDBasicPay.Text = 0
txtYTDOvertime.Text = 0
txtYTDCommission.Text = 0
txtMidYear.Text = 0
txtYTDRemSal.Text = 0
txtYTDRemOT.Text = 0
txtYTDRemWTax.Text = 0
txtYTDRemDed.Text = 0
txtYTDGross.Text = 0
txtYTDSSSPAGIBIGPHIC.Text = 0
txtTaxExemp.Text = 0
txtYTDTaxDue.Text = 0
txtYTDTaxWithHeld.Text = 0
txtYTDTaxRefund.Text = 0

txtAdjSalary.Text = 0
txt13thMonth.Text = 0
txtTaxRefund.Text = 0
txtTotalPay.Text = 0
End Sub

Sub StoreLoanMemVars()
On Error GoTo ErrorCode
Dim Cnt, crt As Integer
Dim LonType As String
grdLoanMasDet.ZOrder 1
Set rsLoanMas = New ADODB.Recordset
    rsLoanMas.Open "select * from loanmas where empno = '" & IMPNO & "' order by dategranted desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
   clearLoangrd
   rsLoanMas.MoveFirst
   crt = 0
   Do While Not rsLoanMas.EOF
      crt = crt + 1
      If Null2String(rsLoanMas!loantype) = "SSAL" Then LonType = "SSS Salary Loan"
      If Null2String(rsLoanMas!loantype) = "CSAL" Then LonType = "SSS Calamity Loan"
      If Null2String(rsLoanMas!loantype) = "PSAL" Then LonType = "Pag-Ibig Salary Loan"
      If Null2String(rsLoanMas!loantype) = "HDMF" Then LonType = "Pag-Ibig HDMF"
      If Null2String(rsLoanMas!loantype) = "MPL" Then LonType = "Pag-Ibig MP Loan"
      If Null2String(rsLoanMas!loantype) = "HLL" Then LonType = "Pag-Ibig Housing Loan"
      If Null2String(rsLoanMas!loantype) = "BLL" Then LonType = "BL Loan"
      grdLoanMas.AddItem LonType & Chr(9) & Null2String(rsLoanMas!acctno) & Chr(9) & Null2String(rsLoanMas!DateGranted) & _
                         Chr(9) & Null2String(rsLoanMas!MaturityDate) & Chr(9) & N2Str2Zero(rsLoanMas!AmountLoaned) & Chr(9) & N2Str2Zero(rsLoanMas!monthlyded) & _
                         Chr(9) & N2Str2Zero(rsLoanMas!smonthlyded) & Chr(9) & N2Str2Zero(rsLoanMas!LoanBalance) & Chr(9) & rsLoanMas!ID & Chr(9) & Null2String(rsLoanMas!empno)
      rsLoanMas.MoveNext
   Loop
   If crt > 0 Then grdLoanMas.RemoveItem 1
Else
   clearLoangrd
End If
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Sub InitLoans()
cboLoanType.Clear
cboLoanType.AddItem "SSS Salary Loan"
cboLoanType.AddItem "SSS Calamity Loan"
cboLoanType.AddItem "Pag-Ibig Salary Loan"
cboLoanType.AddItem "Pag-Ibig MP Loan"
cboLoanType.AddItem "Pag-Ibig Housing Loan"
cboLoanType.AddItem "BL Loan"
txtAcctNo.Text = ""
txtDateGranted.Text = LOGDATE
txtDateStarted.Text = LOGDATE
txtMaturityDate.Text = DateSerial(Year(LOGDATE) + 2, Month(LOGDATE), Day(LOGDATE))
txtAmountLoaned.Text = 0
txtMonthlyDed.Text = 0
txtSMonthlyDed.Text = 0
txtLoanBalance.Text = 0
cmdLoanMas.ZOrder 0
fraLoanMas.ZOrder 0
clearLoangrd
End Sub

Sub StorePHMemvars()
On Error GoTo ErrorCode
If Not rsPH.EOF And Not rsPH.BOF Then
   txtPHMonthly.Text = Null2String(rsPH!employeeshare)
   txtPHStarted.Text = Null2String(rsPH!datestart)
   txtPHLast.Text = Null2String(rsPH!lastdatecont)
   Set rsPHDet = New ADODB.Recordset
       rsPHDet.Open "select * from philhealthdet where aydi = " & rsPH!aydi & " order by deyt desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsPHDet.EOF And Not rsPHDet.BOF Then
      rsPHDet.MoveFirst
      Do While Not rsPHDet.EOF
         grdPhilHealth.AddItem rsPHDet!deyt & Chr(9) & Format(rsPHDet!employeramount, "###,##0.00") & Chr(9) & Format(rsPHDet!employeeamount, "###,##0.00") & Chr(9) & rsPHDet!ID
         rsPHDet.MoveNext
      Loop
      grdPhilHealth.RemoveItem 1
   Else
      clearPHgrd
   End If
Else
   txtPHMonthly.Text = ""
   txtPHStarted.Text = ""
   txtPHLast.Text = ""
End If
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Sub StorePagIbigMemvars()
On Error GoTo ErrorCode
If Not rsPagIbig.EOF And Not rsPagIbig.BOF Then
   txtPagIbigMonthly.Text = Null2String(rsPagIbig!employeeshare)
   txtPagIbigStarted.Text = Null2String(rsPagIbig!datestart)
   txtPagIbigLast.Text = Null2String(rsPagIbig!lastdatecont)
   Set rsPagibigdet = New ADODB.Recordset
       rsPagibigdet.Open "select * from pagibigdet where aydi = " & rsPagIbig!aydi & " order by deyt desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsPagibigdet.EOF And Not rsPagibigdet.BOF Then
      rsPagibigdet.MoveFirst
      Do While Not rsPagibigdet.EOF
         grdPagIbig.AddItem Null2String(rsPagibigdet!deyt) & Chr(9) & N2Str2Zero(rsPagibigdet!employeramount) & Chr(9) & N2Str2Zero(rsPagibigdet!employeeamount) & Chr(9) & rsPagibigdet!ID
         rsPagibigdet.MoveNext
      Loop
      grdPagIbig.RemoveItem 1
   Else
      clearPagIbiggrd
   End If
Else
   txtPagIbigMonthly.Text = ""
   txtPagIbigStarted.Text = ""
   txtPagIbigLast.Text = ""
End If
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Sub StoreTINMemvars()
On Error GoTo ErrorCode
If Not rsTIN.EOF And Not rsTIN.BOF Then
   clearTINgrd
   txtTINMonthly.Text = Null2String(rsTIN!deduction)
   txtTINStarted.Text = Null2String(rsTIN!datestart)
   txtTINLast.Text = Null2String(rsTIN!lastdatecont)
   Set rsTINdet = New ADODB.Recordset
       rsTINdet.Open "select * from TINdet where aydi = " & rsTIN!aydi & " order by deyt desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsTINdet.EOF And Not rsTINdet.BOF Then
      rsTINdet.MoveFirst
      Do While Not rsTINdet.EOF
         grdTIN.AddItem Null2String(rsTINdet!deyt) & Chr(9) & N2Str2Zero(rsTINdet!amount) & Chr(9) & rsTINdet!ID
         rsTINdet.MoveNext
      Loop
      grdTIN.RemoveItem 1
   Else
      clearTINgrd
   End If
Else
   txtTINMonthly.Text = ""
   txtTINStarted.Text = ""
   txtTINLast.Text = ""
End If
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Sub InitMemVars()
Dim MM, ddFROM, YY As String
MM = Trim(Str(Month(LOGDATE)))
YY = Trim(Str(Year(LOGDATE)))
If Day(LOGDATE) > 15 Then
   txtFrom.Text = DateSerial(YY, MM, 16)
   txtTo.Text = lastDay(txtFrom.Text)
Else
   txtFrom.Text = DateSerial(YY, MM, 1)
   txtTo.Text = DateSerial(YY, MM, 15)
End If
txtRate.Text = 0#
txtDailyRate.Text = 0#
txtTaxableAdj.Text = 0
txtNonTaxableAdj.Text = 0
txtOvertime.Text = 0#
txtHoliday.Text = 0#
txtCommission.Text = 0#

txtSSS.Text = 0#
txtMed.Text = 0#
txtPagIbig.Text = 0#
txtTAX.Text = 0#
txtSalLoan.Text = 0#
txtCalLoan.Text = 0#
txtPagHLLoan.Text = 0#
txtPagMPLoan.Text = 0#
txtBLLoan.Text = 0#

txtSalaryAdvance.Text = 0#
txtUndertime.Text = 0#
txtAbsent.Text = 0#
txtTelBill.Text = 0#
txtOthers.Text = 0#
showGrossWage
End Sub

Sub FillPayroll()
Set rsPayroll = New ADODB.Recordset
    rsPayroll.Open "select * from payroll where empno = '" & IMPNO & "' order by paydatefrom desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
If Not rsPayroll.EOF And Not rsPayroll.BOF Then
   Screen.MousePointer = 11
   rsPayroll.MoveFirst
   Do While Not rsPayroll.EOF
      grdPayroll.AddItem Null2Date(rsPayroll!paydatefrom) & _
      Chr(9) & Null2Date(rsPayroll!paydateto) & _
      Chr(9) & N2Str2Zero(rsPayroll!Rate) & _
      Chr(9) & N2Str2Zero(rsPayroll!DailyRate) & _
      Chr(9) & N2Str2Zero(rsPayroll!taxableadj) & _
      Chr(9) & N2Str2Zero(rsPayroll!nontaxableadj) & _
      Chr(9) & N2Str2Zero(rsPayroll!overtime) & _
      Chr(9) & N2Str2Zero(rsPayroll!holiday) & _
      Chr(9) & N2Str2Zero(rsPayroll!commission) & _
      Chr(9) & N2Str2Zero(rsPayroll!gross) & _
      Chr(9) & N2Str2Zero(rsPayroll!sssE) & _
      Chr(9) & N2Str2Zero(rsPayroll!philhealthE) & _
      Chr(9) & N2Str2Zero(rsPayroll!pagibig) & _
      Chr(9) & N2Str2Zero(rsPayroll!tax) & _
      Chr(9) & N2Str2Zero(rsPayroll!ssssalloan) & _
      Chr(9) & N2Str2Zero(rsPayroll!ssscalloan) & _
      Chr(9) & N2Str2Zero(rsPayroll!pagmploan) & _
      Chr(9) & N2Str2Zero(rsPayroll!paghlloan) & Chr(9) & N2Str2Zero(rsPayroll!blloan) & _
      Chr(9) & N2Str2Zero(rsPayroll!undertime) & _
      Chr(9) & N2Str2Zero(rsPayroll!SalaryAdvance) & _
      Chr(9) & N2Str2Zero(rsPayroll!absent) & _
      Chr(9) & N2Str2Zero(rsPayroll!telbill) & _
      Chr(9) & N2Str2Zero(rsPayroll!others) & _
      Chr(9) & N2Str2Zero(rsPayroll!netpay) & _
      Chr(9) & rsPayroll!ID
      rsPayroll.MoveNext
   Loop
   grdPayroll.RemoveItem 1
   Screen.MousePointer = 0
Else
   clearPayrollgrd
End If
End Sub

Sub OtherRefresh()
Set rsSSS = New ADODB.Recordset
    rsSSS.Open "select * from sss where sssno = " & N2Str2Null(rsEmpInfo!sssno) & " order by datestart desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
Set rsPH = New ADODB.Recordset
    rsPH.Open "select * from philhealth where phno = " & N2Str2Null(rsEmpInfo!sssno) & " order by datestart desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
Set rsPagIbig = New ADODB.Recordset
    rsPagIbig.Open "select * from pagibig where pagibigno = " & N2Str2Null(rsEmpInfo!sssno) & " order by datestart desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
Set rsTIN = New ADODB.Recordset
    rsTIN.Open "select * from tin where tinno = " & N2Str2Null(rsEmpInfo!tinno) & " order by datestart desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub clearPayrollgrd()
cleargrid grdPayroll
End Sub

Sub clearSSSgrd()
cleargrid grdSSS
End Sub

Sub clearLoangrd()
cleargrid grdLoanMas
End Sub

Sub clearLoanDetgrd()
cleargrid grdLoanMasDet
End Sub

Sub clearPHgrd()
cleargrid grdPhilHealth
End Sub

Sub clearPagIbiggrd()
cleargrid grdPagIbig
End Sub

Sub clearTINgrd()
cleargrid grdTIN
End Sub

Sub initGrid()
With grdPayroll
   .Rows = 2
   .ColWidth(0) = 900
   .ColWidth(1) = 900
   .ColWidth(2) = 900
   .ColWidth(3) = 900
   .ColWidth(4) = 900
   .ColWidth(5) = 900
   .ColWidth(6) = 900
   .ColWidth(7) = 900
   .ColWidth(8) = 900
   .ColWidth(9) = 900
   .ColWidth(10) = 900
   .ColWidth(11) = 900
   .ColWidth(12) = 900
   .ColWidth(13) = 900
   .ColWidth(14) = 900
   .ColWidth(15) = 900
   .ColWidth(16) = 950
   .ColWidth(17) = 950
   .ColWidth(18) = 950
   .ColWidth(19) = 900
   .ColWidth(20) = 900
   .ColWidth(21) = 900
   .ColWidth(22) = 900
   .ColWidth(23) = 900
   .ColWidth(24) = 900
   .ColWidth(25) = 1
   .Row = 0
   .Col = 0
   .Text = "From"
   .Col = 1
   .Text = "To"
   .Col = 2
   .Text = "Rate"
   .Col = 3
   .Text = "Daily Rate"
   .Col = 4
   .Text = "Taxable Adj."
   .Col = 5
   .Text = "Non-Tax Adj."
   .Col = 6
   .Text = "Overtime"
   .Col = 7
   .Text = "Holiday"
   .Col = 8
   .Text = "Commission"
   .Col = 9
   .Text = "Gross"
   .Col = 10
   .Text = "SSS"
   .Col = 11
   .Text = "Med"
   .Col = 12
   .Text = "Pag-Ibig"
   .Col = 13
   .Text = "TAX"
   .Col = 14
   .Text = "SSS Sal. Loan"
   .Col = 15
   .Text = "SSS Cal. Loan"
   .Col = 16
   .Text = "PI MP Loan"
   .Col = 17
   .Text = "PI HL Loan"
   .Col = 18
   .Text = "BL Loan"
   .Col = 19
   .Text = "UT/Late"
   .Col = 20
   .Text = "Cash Adv."
   .Col = 21
   .Text = "Absent"
   .Col = 22
   .Text = "Tel. Bill"
   .Col = 23
   .Text = "Others"
   .Col = 24
   .Text = "Net Pay"
   .Col = 25
   .Text = "ID"
End With
With grdSSS
   .Rows = 2
   .ColWidth(0) = 1100
   .ColWidth(1) = 1200
   .ColWidth(2) = 1250
   .ColWidth(3) = 1
   .Row = 0
   .Col = 0
   .Text = "Date"
   .Col = 1
   .Text = "Employer Share"
   .Col = 2
   .Text = "Employee Share"
   .Col = 3
   .Text = "ID"
End With
With grdLoanMas
   .Rows = 2
   .ColWidth(0) = 1600
   .ColWidth(1) = 1300
   .ColWidth(2) = 1300
   .ColWidth(3) = 1300
   .ColWidth(4) = 1300
   .ColWidth(5) = 1300
   .ColWidth(6) = 1300
   .ColWidth(7) = 1300
   .ColWidth(8) = 1
   .ColWidth(9) = 1
   .Row = 0
   .Col = 0
   .Text = "Loan Type"
   .Col = 1
   .Text = "Account No."
   .Col = 2
   .Text = "Date Granted"
   .Col = 3
   .Text = "Maturity Date"
   .Col = 4
   .Text = "Amount Loaned"
   .Col = 5
   .Text = "Monthly Ded"
   .Col = 6
   .Text = "S-Monthly Ded"
   .Col = 7
   .Text = "Loan Balance"
   .Col = 8
   .Text = "ID"
   .Col = 9
   .Text = "empno"
End With
With grdLoanMasDet
   .Rows = 2
   .ColWidth(0) = 1300
   .ColWidth(1) = 1300
   .ColWidth(2) = 1300
   .ColWidth(3) = 1
   .Row = 0
   .Col = 0
   .Text = "Acct. No"
   .Col = 1
   .Text = "Date"
   .Col = 2
   .Text = "Amount"
   .Col = 3
   .Text = "ID"
End With
With grdPhilHealth
   .Rows = 2
   .ColWidth(0) = 1100
   .ColWidth(1) = 1200
   .ColWidth(2) = 1250
   .ColWidth(3) = 1
   .Row = 0
   .Col = 0
   .Text = "Date"
   .Col = 1
   .Text = "Employer Share"
   .Col = 2
   .Text = "Employee Share"
   .Col = 3
   .Text = "ID"
End With
With grdPagIbig
   .Rows = 2
   .ColWidth(0) = 1100
   .ColWidth(1) = 1200
   .ColWidth(2) = 1250
   .ColWidth(3) = 1
   .Row = 0
   .Col = 0
   .Text = "Date"
   .Col = 1
   .Text = "Employer Share"
   .Col = 2
   .Text = "Employee Share"
   .Col = 3
   .Text = "ID"
End With
With grdTIN
   .Rows = 2
   .ColWidth(0) = 1200
   .ColWidth(1) = 1200
   .ColWidth(2) = 1
   .Row = 0
   .Col = 0
   .Text = "Date"
   .Col = 1
   .Text = "Deduction"
   .Col = 2
   .Text = "ID"
End With
End Sub

Sub grdStore()
On Error GoTo ErrorCode
Screen.MousePointer = 11
clearPayrollgrd
clearLoangrd
clearSSSgrd
clearPHgrd
clearPagIbiggrd
clearTINgrd
FillPayroll
OtherRefresh
If Not rsSSS.EOF And Not rsSSS.BOF Then
   txtSSSMonthly.Text = Null2String(rsSSS!employeeshare)
   txtSSSStarted.Text = Null2String(rsSSS!datestart)
   txtSSSLast.Text = Null2String(rsSSS!lastdatecont)
   Set rsSSSdet = New ADODB.Recordset
       rsSSSdet.Open "select * from sssdet where aydi = " & rsSSS!aydi & " order by deyt desc", gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsSSSdet.EOF And Not rsSSSdet.BOF Then
      rsSSSdet.MoveFirst
      Do While Not rsSSSdet.EOF
         grdSSS.AddItem Null2Date(rsSSSdet!deyt) & Chr(9) & N2Str2Zero(rsSSSdet!employeramount) & Chr(9) & N2Str2Zero(rsSSSdet!employeeamount) & Chr(9) & rsSSSdet!ID
         rsSSSdet.MoveNext
      Loop
      grdSSS.RemoveItem 1
   Else
      clearSSSgrd
   End If
Else
   txtSSSMonthly.Text = ""
   txtSSSStarted.Text = ""
   txtSSSLast.Text = ""
End If
StoreLoanMemVars
StorePHMemvars
StorePagIbigMemvars
StoreTINMemvars
Screen.MousePointer = 0
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
LEDGERSHOW = False
UnloadForm Me
End Sub

Private Sub grdLoanMas_Click()
Dim Cnt As Integer
Dim LoanAcctNo As String
Cnt = 0
grdLoanMas.Row = grdLoanMas.Row
grdLoanMas.Col = 1
LoanAcctNo = N2Str2Null(grdLoanMas.Text)
grdLoanMas.Col = 9
clearLoanDetgrd
grdLoanMasDet.ZOrder 0
If grdLoanMas.Text <> "" Then
   Set rsLoanmasDet = New ADODB.Recordset
       rsLoanmasDet.Open "select * from loanmasdet where empno = " & N2Str2Null(grdLoanMas.Text) & " AND acctno = " & LoanAcctNo & " order by deyt desc", gconHRMS
   If Not rsLoanmasDet.EOF And Not rsLoanmasDet.BOF Then
      rsLoanmasDet.MoveFirst
      Do While Not rsLoanmasDet.EOF
         Cnt = Cnt + 1
         grdLoanMasDet.AddItem Null2String(rsLoanmasDet!acctno) & Chr(9) & Null2String(rsLoanmasDet!deyt) & Chr(9) & N2Str2Zero(rsLoanmasDet!amount) & Chr(9) & rsLoanmasDet!ID
         rsLoanmasDet.MoveNext
      Loop
   End If
   If Cnt > 0 Then grdLoanMasDet.RemoveItem 1
End If
End Sub

Private Sub grdLoanMas_DblClick()
grdLoanMas.Row = grdLoanMas.Row
grdLoanMas.Col = 8
CLID = grdLoanMas.Text
If CLID <> "" Then
   Set rsLoanMas = New ADODB.Recordset
       rsLoanMas.Open "select * from LoanMas where id =" & CLID, gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
      cmdLoanMas.ZOrder 0
      fraLoanMas.ZOrder 0
      fraLoanMas.Enabled = True
      If Null2String(rsLoanMas!loantype) = "SSAL" Then
        cboLoanType.Text = "SSS Salary Loan"
      End If
      If Null2String(rsLoanMas!loantype) = "CSAL" Then
        cboLoanType.Text = "SSS Calamity Loan"
      End If
      If Null2String(rsLoanMas!loantype) = "PSAL" Then
        cboLoanType.Text = "Pag-Ibig Salary Loan"
      End If
      If Null2String(rsLoanMas!loantype) = "MPL" Then
        cboLoanType.Text = "Pag-Ibig MP Loan"
      End If
      If Null2String(rsLoanMas!loantype) = "HLL" Then
        cboLoanType.Text = "Pag-Ibig Housing Loan"
      End If
      If Null2String(rsLoanMas!loantype) = "BLL" Then
        cboLoanType.Text = "BL Loan"
      End If
      txtAcctNo.Text = Null2String(rsLoanMas!acctno)
      txtDateGranted.Text = Null2Date(rsLoanMas!DateGranted)
      txtDateStarted.Text = Null2Date(rsLoanMas!DateStarted)
      txtMaturityDate.Text = Null2Date(rsLoanMas!MaturityDate)
      txtAmountLoaned.Text = N2Str2Zero(rsLoanMas!AmountLoaned)
      txtMonthlyDed.Text = N2Str2Zero(rsLoanMas!monthlyded)
      txtSMonthlyDed.Text = N2Str2Zero(rsLoanMas!smonthlyded)
      txtLoanBalance.Text = N2Str2Zero(rsLoanMas!LoanBalance)
      
      If Null2String(rsLoanMas!Deduction_Option) = "OPT15thOnly" Then Opt1.Value = True
      If Null2String(rsLoanMas!Deduction_Option) = "OPT30thOnly" Then Opt2.Value = True
      If Null2String(rsLoanMas!Deduction_Option) = "OPT15th30th" Then Opt3.Value = True
      txtOtherTypeDed.Text = Null2String(rsLoanMas!OtherTypeDed)
      txtOtherTypeDedAmount.Text = N2Str2Zero(rsLoanMas!OtherTypeDedAmount)
      Picture1.Visible = False
      Picture2.Visible = True
   End If
End If
End Sub

Private Sub grdPagIbig_DblClick()
Dim piID As String
grdPagIbig.Row = grdPagIbig.Row
grdPagIbig.Col = 3
piID = grdPagIbig.Text
If piID <> "" Then
   Set rsPagibigdet = New ADODB.Recordset
       rsPagibigdet.Open "select * from PagIbigdet where id =" & piID, gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsPagibigdet.EOF And Not rsPagibigdet.BOF Then
      StoreDed (Null2Date(rsPagibigdet!deyt))
   End If
End If
End Sub

Private Sub grdPayroll_DblClick()
cmdEdit.Value = True
End Sub

Private Sub grdPhilHealth_DblClick()
Dim phID As String
grdPhilHealth.Row = grdPhilHealth.Row
grdPhilHealth.Col = 3
phID = grdPhilHealth.Text
If phID <> "" Then
   Set rsPHDet = New ADODB.Recordset
       rsPHDet.Open "select * from philhealthdet where id =" & phID, gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsPHDet.EOF And Not rsPHDet.BOF Then
      StoreDed (Null2Date(rsPHDet!deyt))
   End If
End If
End Sub

Private Sub grdLoanMasDet_DblClick()
Dim slID As String
grdLoanMasDet.Row = grdLoanMasDet.Row
grdLoanMasDet.Col = 3
slID = grdLoanMasDet.Text
If slID <> "" Then
   Set rsLoanmasDet = New ADODB.Recordset
       rsLoanmasDet.Open "select * from LoanMasdet where id =" & slID, gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsLoanmasDet.EOF And Not rsLoanmasDet.BOF Then
      StoreDed (Null2Date(rsLoanmasDet!deyt))
   End If
End If
End Sub

Function StoreDed(Diyt As String)
Set rsPayroll = New ADODB.Recordset
    rsPayroll.Open "select * from payroll where empno = '" & IMPNO & "' and paydateto = #" & Format(Diyt, "Short Date") & "#", gconHRMS, adOpenForwardOnly, adLockReadOnly
If Not rsPayroll.EOF And Not rsPayroll.BOF Then
   MAIDIT (rsPayroll!ID)
End If
End Function

Private Sub grdSSS_DblClick()
Dim ssID As String
grdSSS.Row = grdSSS.Row
grdSSS.Col = 3
ssID = grdSSS.Text
If ssID <> "" Then
   Set rsSSSdet = New ADODB.Recordset
       rsSSSdet.Open "select * from sssdet where id =" & ssID, gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsSSSdet.EOF And Not rsSSSdet.BOF Then
      StoreDed (Null2Date(rsSSSdet!deyt))
   End If
End If
End Sub

Private Sub grdTIN_DblClick()
Dim tiID As String
grdTIN.Row = grdTIN.Row
grdTIN.Col = 2
tiID = grdTIN.Text
If tiID <> "" Then
   Set rsTINdet = New ADODB.Recordset
       rsTINdet.Open "select * from tindet where id =" & tiID, gconHRMS, adOpenForwardOnly, adLockReadOnly
   If Not rsTINdet.EOF And Not rsTINdet.BOF Then
      StoreDed (Null2Date(rsTINdet!deyt))
   End If
End If
End Sub

Private Sub TabSSS_Click(PreviousTab As Integer)
If TabSSS.Tab = 0 Then
   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
ElseIf TabSSS.Tab = 1 Then
   cmdAdd.Enabled = True
   cmdDelete.Enabled = False
Else
   cmdAdd.Enabled = False
   cmdDelete.Enabled = False
End If
cmdEdit.Enabled = True
cmdPrint.Enabled = True
End Sub

Private Sub txtAbsent_Change()
ShowTotDed
End Sub

Private Sub txtAbsent_LostFocus()
ShowTotDed
End Sub

Private Sub txtAmountLoaned_Change()
txtMonthlyDed.Text = Round(((NumericVal(txtAmountLoaned.Text) * 0.12) + NumericVal(txtAmountLoaned.Text)) / 24, 2)
txtSMonthlyDed.Text = Round(NumericVal(txtMonthlyDed.Text) / 2, 2)
txtLoanBalance.Text = Round((NumericVal(txtAmountLoaned.Text) * 0.12) + NumericVal(txtAmountLoaned.Text), 2)
End Sub

Private Sub txtCalLoan_Change()
ShowTotDed
End Sub

Private Sub txtCalLoan_LostFocus()
ShowTotDed
End Sub

Private Sub txtOtherTypeDedAmount_Change()
txtLoanBalance.Text = Round(NumericVal(txtLoanBalance.Text) - NumericVal(txtOtherTypeDedAmount.Text), 2)
End Sub

Private Sub txtSalaryAdvance_Change()
ShowTotDed
End Sub

Private Sub txtCommission_Change()
showGrossWage
End Sub

Private Sub txtCommission_LostFocus()
showGrossWage
End Sub

Private Sub txtHoliday_Change()
showGrossWage
End Sub

Private Sub txtHoliday_LostFocus()
showGrossWage
End Sub

Private Sub txtMed_Change()
ShowTotDed
End Sub

Private Sub txtMed_LostFocus()
ShowTotDed
End Sub

Private Sub txtMidYear_LostFocus()
YTDComputation
End Sub

Private Sub txtMonthlyDed_LostFocus()
txtSMonthlyDed.Text = NumericVal(txtMonthlyDed.Text) / 2
End Sub

Private Sub txtNonTaxableAdj_Change()
showGrossWage
End Sub

Private Sub txtOthers_Change()
ShowTotDed
End Sub

Private Sub txtOthers_LostFocus()
ShowTotDed
End Sub

Private Sub txtOvertime_Change()
showGrossWage
End Sub

Private Sub txtOvertime_LostFocus()
showGrossWage
End Sub

Private Sub txtPagIbig_Change()
ShowTotDed
End Sub

Private Sub txtPagIbig_LostFocus()
ShowTotDed
End Sub

Private Sub txtPHMonthly_LostFocus()
If AddorEdit = "ADD" Then
   If NumericVal(txtPHMonthly.Text) = 0# Then
      txtPHMonthly.Text = Round(PhilHealthShare(SetSalary(Null2String(rsEmpInfo!salarycode))), 2)
   End If
End If
End Sub

Private Sub txtRate_Change()
showGrossWage
If AddorEdit = "ADD" Then
   txtDailyRate.Text = Round((NumericVal(txtRate.Text) * 12) / 314, 2)
End If
End Sub

Private Sub txtRate_LostFocus()
showGrossWage
End Sub

Private Sub txtSalLoan_Change()
ShowTotDed
End Sub

Private Sub txtSalLoan_LostFocus()
ShowTotDed
End Sub

Private Sub txtSMonthlyDed_Change()
txtMonthlyDed.Text = NumericVal(txtSMonthlyDed.Text) * 2
End Sub

Private Sub txtSSS_Change()
ShowTotDed
End Sub

Private Sub txtSSS_LostFocus()
ShowTotDed
End Sub

Private Sub txtSSSMonthly_LostFocus()
If NumericVal(txtSSSMonthly.Text) = 0# Then
   txtSSSMonthly.Text = Round(EmployeeSSSshare(SetSalary(Null2String(rsEmpInfo!salarycode))), 2)
End If
End Sub

Private Sub txtTax_Change()
ShowTotDed
End Sub

Private Sub txtTAX_LostFocus()
ShowTotDed
End Sub

Private Sub txtTaxableAdj_Change()
showGrossWage
End Sub

Private Sub txtTelBill_Change()
ShowTotDed
End Sub

Private Sub txtTelBill_LostFocus()
ShowTotDed
End Sub

Private Sub txtUndertime_Change()
ShowTotDed
End Sub

Private Sub txtUndertime_LostFocus()
ShowTotDed
End Sub

Private Sub txtYTDRemDed_Change()
YTDComputation
End Sub

Private Sub txtYTDRemOT_Change()
YTDComputation
End Sub

Private Sub txtYTDRemSal_Change()
YTDComputation
End Sub

Private Sub txtYTDRemWTax_Change()
YTDComputation
End Sub

Sub YTDComputation()
On Error Resume Next
   txtYTDGross.Text = Round(NumericVal(txtYTDRemSal.Text) + NumericVal(txtYTDBasicPay.Text) + NumericVal(txtYTDCommission.Text) + NumericVal(txtYTDOvertime.Text) + NumericVal(txtYTDRemOT.Text), 2)
   txtYTDSSSPAGIBIGPHIC.Text = Round(N2Str2Zero(rsYTDDetails!nontaxable), 2)
   txtTaxableIncome.Text = Round(NumericVal(txtYTDGross.Text) - (NumericVal(txtYTDOvertime.Text) + NumericVal(txtTaxExemp.Text) + NumericVal(txtYTDSSSPAGIBIGPHIC.Text)), 2)
   If NumericVal(Tax_Due(NumericVal(txtTaxableIncome.Text))) <= 0 Then
      txtYTDTaxDue.Text = "Exempted"
   Else
      txtYTDTaxDue.Text = Round(Tax_Due(NumericVal(txtTaxableIncome.Text)), 2)
   End If
   txtYTDTaxRefund.Text = Round(NumericVal(txtYTDTaxDue.Text) - (NumericVal(txtYTDTaxWithHeld.Text) + NumericVal(txtYTDRemWTax.Text)), 2)
   
   txtAdjSalary.Text = Round(NumericVal(txtYTDRemSal.Text), 2)
   txtMidYear.Text = Round(NumericVal(txtMidYear.Text), 2)
   txt13thMonth.Text = Round(((NumericVal(txtYTDBasicPay.Text) + NumericVal(txtYTDRemSal.Text)) / 12) - txtMidYear.Text, 2)
   If NumericVal(txtYTDTaxRefund.Text) < 0 Then
      txtTaxRefund.Text = Round(Abs(txtYTDTaxRefund.Text), 2)
   Else
      txtTaxRefund.Text = Round(NumericVal(txtYTDTaxRefund.Text), 2)
   End If
   txtTotalPay.Text = Round((NumericVal(txtAdjSalary.Text) + NumericVal(txt13thMonth.Text) + NumericVal(txtTaxRefund.Text) + NumericVal(txtYTDRemOT.Text)) - NumericVal(txtYTDRemDed.Text), 2)
End Sub

Private Sub lsAdjustment_ItemClick(ByVal Item As MSComctlLib.ListItem)
rsEmpInfo.Bookmark = rsFind(rsEmpInfo.Clone, "empno", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lsAdjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lsAdjustment
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

Private Sub lsAdjustment_DblClick()
cmdEdit.Value = True
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then FillGrid Else FillSearchGrid (txtSearch.Text)
End Sub

Sub FillGrid()
Dim rsEmpInfo2 As ADODB.Recordset
lsAdjustment.Sorted = False: lsAdjustment.ListItems.Clear
Set rsEmpInfo2 = New ADODB.Recordset
Set rsEmpInfo2 = gconHRMS.Execute("select lastname+', '+firstname,empno from empinfo WHERE EMPLEVEL = 'E' AND RESIGNED IS NULL order by lastname+', '+firstname asc")
If Not (rsEmpInfo2.EOF And rsEmpInfo2.BOF) Then
   Listview_Loadval Me.lsAdjustment.ListItems, rsEmpInfo2
   lsAdjustment.Refresh
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsEmpInfo2 As ADODB.Recordset
lsAdjustment.Sorted = False: lsAdjustment.ListItems.Clear
Set rsEmpInfo2 = New ADODB.Recordset
Set rsEmpInfo2 = gconHRMS.Execute("select lastname+', '+firstname,empno from empinfo  where EMPLEVEL = 'E' AND RESIGNED IS NULL and lastname+', '+firstname like'" & XXX & "%' order by lastname+', '+firstname asc")
If Not (rsEmpInfo2.EOF And rsEmpInfo2.BOF) Then
   Listview_Loadval Me.lsAdjustment.ListItems, rsEmpInfo2
   lsAdjustment.Refresh
End If
End Sub

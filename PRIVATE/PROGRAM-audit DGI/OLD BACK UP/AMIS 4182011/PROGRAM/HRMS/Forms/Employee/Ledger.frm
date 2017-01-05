VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmHRMSLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Ledger"
   ClientHeight    =   7425
   ClientLeft      =   345
   ClientTop       =   780
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Haettenschweiler"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "Ledger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   11685
   Begin TabDlg.SSTab TabSSS 
      Height          =   4905
      Left            =   2640
      TabIndex        =   29
      Top             =   1620
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   8652
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   1058
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
      TabPicture(0)   =   "Ledger.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPayroll"
      Tab(0).Control(1)=   "fraDetails"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Loans"
      TabPicture(1)   =   "Ledger.frx":075C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "SSS/PHIC."
      TabPicture(2)   =   "Ledger.frx":0BAE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pag-Ibig / TAX"
      TabPicture(3)   =   "Ledger.frx":0EC8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "YTD Details"
      TabPicture(4)   =   "Ledger.frx":11E2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture6"
      Tab(4).ControlCount=   1
      Begin VB.PictureBox Picture3 
         Height          =   4125
         Left            =   60
         ScaleHeight     =   4065
         ScaleWidth      =   8775
         TabIndex        =   30
         Top             =   690
         Width           =   8835
         Begin MSFlexGridLib.MSFlexGrid grdLoanMasDet 
            Height          =   2115
            Left            =   60
            TabIndex        =   48
            Top             =   1650
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   3731
            _Version        =   393216
            Cols            =   4
            ForeColorFixed  =   0
            BackColorSel    =   -2147483633
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
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
         Begin MSFlexGridLib.MSFlexGrid grdLoanMas 
            Height          =   4005
            Left            =   30
            TabIndex        =   50
            Top             =   30
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   7064
            _Version        =   393216
            Cols            =   10
            ForeColorFixed  =   0
            BackColorSel    =   -2147483633
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            TextStyleFixed  =   3
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
         Begin VB.CommandButton cmdLoanMas 
            Caption         =   "Command1"
            Height          =   3315
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   360
            Width           =   4125
         End
         Begin VB.Frame fraLoanMas 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   3195
            Left            =   2370
            TabIndex        =   31
            Top             =   330
            Width           =   3945
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
               Left            =   1470
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   210
               Width           =   2355
            End
            Begin MSMask.MaskEdBox txtAmountLoaned 
               Height          =   315
               Left            =   2550
               TabIndex        =   33
               Top             =   1650
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
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtMonthlyDed 
               Height          =   315
               Left            =   2550
               TabIndex        =   34
               Top             =   2010
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
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtSMonthlyDed 
               Height          =   315
               Left            =   2550
               TabIndex        =   35
               Top             =   2370
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
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtLoanBalance 
               Height          =   315
               Left            =   2550
               TabIndex        =   36
               Top             =   2730
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
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtAcctNo 
               Height          =   315
               Left            =   1470
               TabIndex        =   37
               Top             =   570
               Width           =   2355
               _ExtentX        =   4154
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
               Left            =   2550
               TabIndex        =   38
               Top             =   930
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtMaturityDate 
               Height          =   315
               Left            =   2550
               TabIndex        =   39
               Top             =   1290
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
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
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
               Left            =   30
               TabIndex        =   47
               Top             =   1680
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
               Left            =   30
               TabIndex        =   46
               Top             =   2040
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
               Left            =   30
               TabIndex        =   45
               Top             =   2400
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
               Left            =   30
               TabIndex        =   44
               Top             =   2760
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
               Left            =   30
               TabIndex        =   43
               Top             =   1320
               Width           =   2445
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
               TabIndex        =   42
               Top             =   600
               Width           =   1845
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
               TabIndex        =   41
               Top             =   240
               Width           =   1815
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
               Left            =   30
               TabIndex        =   40
               Top             =   960
               Width           =   2445
            End
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   4185
         Left            =   -74940
         ScaleHeight     =   4125
         ScaleWidth      =   8805
         TabIndex        =   51
         Top             =   660
         Width           =   8865
         Begin VB.CommandButton cmdSelectYear 
            Caption         =   "Show"
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
            Left            =   2760
            TabIndex        =   54
            Top             =   30
            Width           =   1125
         End
         Begin VB.ComboBox cboYear 
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
            Left            =   1530
            TabIndex        =   53
            Text            =   "Combo1"
            Top             =   30
            Width           =   1185
         End
         Begin VB.Frame fraYTD 
            Appearance      =   0  'Flat
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
            Height          =   3825
            Left            =   30
            TabIndex        =   52
            Top             =   300
            Width           =   8745
            Begin TabDlg.SSTab SSTab1 
               Height          =   3585
               Left            =   4080
               TabIndex        =   192
               Top             =   180
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   6324
               _Version        =   393216
               Tab             =   2
               TabHeight       =   520
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "FINAL PAY"
               TabPicture(0)   =   "Ledger.frx":14FC
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "Label60"
               Tab(0).Control(1)=   "Label69"
               Tab(0).Control(2)=   "Label70"
               Tab(0).Control(3)=   "Label71"
               Tab(0).Control(4)=   "Label76"
               Tab(0).Control(5)=   "Label77"
               Tab(0).Control(6)=   "Label78"
               Tab(0).Control(7)=   "Label79"
               Tab(0).Control(8)=   "Label80"
               Tab(0).Control(9)=   "Label81"
               Tab(0).Control(10)=   "txtTOTAL_PAY"
               Tab(0).Control(11)=   "txtFP_TOTALPAY"
               Tab(0).Control(12)=   "txtFP_REMAININGCOLA"
               Tab(0).Control(13)=   "txtFP_REMAININGSALARY"
               Tab(0).Control(14)=   "txtFP_SICKLEAVE"
               Tab(0).Control(15)=   "txtFP_VACATIONLEAVE"
               Tab(0).Control(16)=   "txtFP_REMAININGDEDUCTION"
               Tab(0).Control(17)=   "txtFP_REMAININGTAXWITHHELD"
               Tab(0).Control(18)=   "txtFP_REMAININGOVERTIME"
               Tab(0).Control(19)=   "txtFP_REMAININGCOMMISSION"
               Tab(0).ControlCount=   20
               TabCaption(1)   =   "DEDUCTIONS"
               TabPicture(1)   =   "Ledger.frx":1518
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Label82"
               Tab(1).Control(1)=   "Label83"
               Tab(1).Control(2)=   "Label84"
               Tab(1).Control(3)=   "Label85"
               Tab(1).Control(4)=   "Label86"
               Tab(1).Control(5)=   "Label87"
               Tab(1).Control(6)=   "Label66"
               Tab(1).Control(7)=   "Label75"
               Tab(1).Control(8)=   "Label90"
               Tab(1).Control(9)=   "txtDED_HDMFPREMIUM"
               Tab(1).Control(10)=   "txtTOTAL_REMDED"
               Tab(1).Control(11)=   "txtDED_TAXPAYABLE"
               Tab(1).Control(12)=   "txtDED_OTHERS"
               Tab(1).Control(13)=   "txtDED_UNDERTIME"
               Tab(1).Control(14)=   "txtDED_ABSENT"
               Tab(1).Control(15)=   "txtDED_PHICPREMIUM"
               Tab(1).Control(16)=   "txtDED_SSSPREMIUM"
               Tab(1).ControlCount=   17
               TabCaption(2)   =   "SUMMARY"
               TabPicture(2)   =   "Ledger.frx":1534
               Tab(2).ControlEnabled=   -1  'True
               Tab(2).Control(0)=   "Label43"
               Tab(2).Control(0).Enabled=   0   'False
               Tab(2).Control(1)=   "Label56"
               Tab(2).Control(1).Enabled=   0   'False
               Tab(2).Control(2)=   "Label58"
               Tab(2).Control(2).Enabled=   0   'False
               Tab(2).Control(3)=   "Label64"
               Tab(2).Control(3).Enabled=   0   'False
               Tab(2).Control(4)=   "Label65"
               Tab(2).Control(4).Enabled=   0   'False
               Tab(2).Control(5)=   "Label68"
               Tab(2).Control(5).Enabled=   0   'False
               Tab(2).Control(6)=   "Label88"
               Tab(2).Control(6).Enabled=   0   'False
               Tab(2).Control(7)=   "Label89"
               Tab(2).Control(7).Enabled=   0   'False
               Tab(2).Control(8)=   "Label91"
               Tab(2).Control(8).Enabled=   0   'False
               Tab(2).Control(9)=   "txtTaxableAdjustment"
               Tab(2).Control(9).Enabled=   0   'False
               Tab(2).Control(10)=   "txtSUMMARY_BASICPAY"
               Tab(2).Control(10).Enabled=   0   'False
               Tab(2).Control(11)=   "txtSUMMARY_COMMISSION"
               Tab(2).Control(11).Enabled=   0   'False
               Tab(2).Control(12)=   "txtSUMMARY_BONUS"
               Tab(2).Control(12).Enabled=   0   'False
               Tab(2).Control(13)=   "txtSUMMARY_MIDYEAR"
               Tab(2).Control(13).Enabled=   0   'False
               Tab(2).Control(14)=   "txtSUMMARY_ADJUSTEDSALARY"
               Tab(2).Control(14).Enabled=   0   'False
               Tab(2).Control(15)=   "txtSUMMARY_TAXREFUND"
               Tab(2).Control(15).Enabled=   0   'False
               Tab(2).Control(16)=   "txtSUMMARY_13THMONTHPAY"
               Tab(2).Control(16).Enabled=   0   'False
               Tab(2).Control(17)=   "txtSUMMARY_OVERTIME"
               Tab(2).Control(17).Enabled=   0   'False
               Tab(2).ControlCount=   18
               Begin MSMask.MaskEdBox txtFP_REMAININGCOMMISSION 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   94
                  Top             =   1920
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
               Begin MSMask.MaskEdBox txtFP_REMAININGOVERTIME 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   93
                  Top             =   1620
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
               Begin MSMask.MaskEdBox txtFP_REMAININGTAXWITHHELD 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   95
                  Top             =   2220
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
               Begin MSMask.MaskEdBox txtFP_REMAININGDEDUCTION 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   96
                  Top             =   2880
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
               Begin MSMask.MaskEdBox txtFP_VACATIONLEAVE 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   89
                  Top             =   420
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
               Begin MSMask.MaskEdBox txtFP_SICKLEAVE 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   90
                  Top             =   720
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
               Begin MSMask.MaskEdBox txtFP_REMAININGSALARY 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   91
                  Top             =   1020
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
               Begin MSMask.MaskEdBox txtFP_REMAININGCOLA 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   92
                  Top             =   1320
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
               Begin MSMask.MaskEdBox txtFP_TOTALPAY 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   97
                  Top             =   3180
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
               Begin MSMask.MaskEdBox txtDED_SSSPREMIUM 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   98
                  Top             =   810
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
               Begin MSMask.MaskEdBox txtDED_PHICPREMIUM 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   99
                  Top             =   1110
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
               Begin MSMask.MaskEdBox txtDED_ABSENT 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   101
                  Top             =   1710
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
               Begin MSMask.MaskEdBox txtDED_UNDERTIME 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   102
                  Top             =   2010
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
               Begin MSMask.MaskEdBox txtDED_OTHERS 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   103
                  Top             =   2310
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
               Begin MSMask.MaskEdBox txtDED_TAXPAYABLE 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   104
                  Top             =   2610
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
               Begin MSMask.MaskEdBox txtSUMMARY_OVERTIME 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   107
                  Top             =   1050
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
               Begin MSMask.MaskEdBox txtSUMMARY_13THMONTHPAY 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   111
                  Top             =   2250
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
               Begin MSMask.MaskEdBox txtSUMMARY_TAXREFUND 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   113
                  Top             =   2850
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
               Begin MSMask.MaskEdBox txtSUMMARY_ADJUSTEDSALARY 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   109
                  Top             =   1650
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
               Begin MSMask.MaskEdBox txtSUMMARY_MIDYEAR 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   110
                  Top             =   1950
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
               Begin MSMask.MaskEdBox txtSUMMARY_BONUS 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   112
                  Top             =   2550
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
               Begin MSMask.MaskEdBox txtSUMMARY_COMMISSION 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   106
                  Top             =   750
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
               Begin MSMask.MaskEdBox txtSUMMARY_BASICPAY 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   105
                  Top             =   450
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
               Begin MSMask.MaskEdBox txtTOTAL_REMDED 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   228
                  Top             =   3150
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
               Begin MSMask.MaskEdBox txtTOTAL_PAY 
                  Height          =   255
                  Left            =   -71880
                  TabIndex        =   230
                  Top             =   2520
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
               Begin MSMask.MaskEdBox txtDED_HDMFPREMIUM 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   100
                  Top             =   1410
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
               Begin MSMask.MaskEdBox txtTaxableAdjustment 
                  Height          =   255
                  Left            =   2970
                  TabIndex        =   108
                  Top             =   1350
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
               Begin VB.Label Label91 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAXABLE ADJUSTMENT"
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
                  Left            =   420
                  TabIndex        =   233
                  Top             =   1350
                  Width           =   2235
               End
               Begin VB.Label Label90 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "HDMF PREMIUM"
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
                  Left            =   -74130
                  TabIndex        =   232
                  Top             =   1410
                  Width           =   2445
               End
               Begin VB.Label Label81 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL PAY"
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
                  Left            =   -74730
                  TabIndex        =   231
                  Top             =   2520
                  Width           =   2445
               End
               Begin VB.Label Label75 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL REMAINING DEDUCTIONS"
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
                  Height          =   495
                  Left            =   -74190
                  TabIndex        =   229
                  Top             =   2910
                  Width           =   3345
               End
               Begin VB.Label Label66 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "REMAINING DEDUCTIONS"
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
                  Left            =   -74130
                  TabIndex        =   227
                  Top             =   480
                  Width           =   2955
               End
               Begin VB.Label Label89 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "BASIC PAY"
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
                  Left            =   420
                  TabIndex        =   226
                  Top             =   450
                  Width           =   2505
               End
               Begin VB.Label Label88 
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
                  Left            =   420
                  TabIndex        =   225
                  Top             =   750
                  Width           =   2505
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
                  Left            =   420
                  TabIndex        =   224
                  Top             =   1950
                  Width           =   2175
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
                  Left            =   420
                  TabIndex        =   223
                  Top             =   1650
                  Width           =   2235
               End
               Begin VB.Label Label64 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX PAYABLE/(TAX REFUND)"
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
                  Left            =   420
                  TabIndex        =   222
                  Top             =   2850
                  Width           =   2565
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
                  Left            =   420
                  TabIndex        =   221
                  Top             =   2250
                  Width           =   1845
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
                  Left            =   420
                  TabIndex        =   220
                  Top             =   1050
                  Width           =   2175
               End
               Begin VB.Label Label43 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "BONUS"
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
                  Left            =   420
                  TabIndex        =   219
                  Top             =   2550
                  Width           =   1845
               End
               Begin VB.Label Label87 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX PAYABLE"
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
                  Left            =   -74130
                  TabIndex        =   218
                  Top             =   2610
                  Width           =   2445
               End
               Begin VB.Label Label86 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "OTHERS"
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
                  Left            =   -74130
                  TabIndex        =   217
                  Top             =   2310
                  Width           =   2445
               End
               Begin VB.Label Label85 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "UNDERTIME"
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
                  Left            =   -74130
                  TabIndex        =   216
                  Top             =   2010
                  Width           =   2445
               End
               Begin VB.Label Label84 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ABSENT"
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
                  Left            =   -74130
                  TabIndex        =   215
                  Top             =   1710
                  Width           =   2445
               End
               Begin VB.Label Label83 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "PHIC PREMIUM"
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
                  Left            =   -74130
                  TabIndex        =   214
                  Top             =   1110
                  Width           =   2445
               End
               Begin VB.Label Label82 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "SSS PREMIUM"
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
                  Left            =   -74130
                  TabIndex        =   213
                  Top             =   810
                  Width           =   2445
               End
               Begin VB.Label Label80 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "FINAL NET PAY"
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
                  Left            =   -74730
                  TabIndex        =   212
                  Top             =   3180
                  Width           =   2445
               End
               Begin VB.Label Label79 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "REMAINING COLA"
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
                  Left            =   -74730
                  TabIndex        =   211
                  Top             =   1320
                  Width           =   2175
               End
               Begin VB.Label Label78 
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
                  Left            =   -74730
                  TabIndex        =   210
                  Top             =   1020
                  Width           =   2175
               End
               Begin VB.Label Label77 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "SICK LEAVE"
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
                  Left            =   -74730
                  TabIndex        =   209
                  Top             =   720
                  Width           =   2445
               End
               Begin VB.Label Label76 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "VACATION LEAVE"
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
                  Left            =   -74730
                  TabIndex        =   208
                  Top             =   420
                  Width           =   2445
               End
               Begin VB.Label Label71 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "LESS: REMAINING DEDUCTION"
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
                  Left            =   -74730
                  TabIndex        =   207
                  Top             =   2880
                  Width           =   2925
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
                  Left            =   -74730
                  TabIndex        =   206
                  Top             =   2220
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
                  Left            =   -74730
                  TabIndex        =   205
                  Top             =   1620
                  Width           =   2445
               End
               Begin VB.Label Label60 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "REMAINING COMMISSION"
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
                  Left            =   -74730
                  TabIndex        =   204
                  Top             =   1920
                  Width           =   2505
               End
            End
            Begin MSMask.MaskEdBox txtYTDTaxRefund 
               Height          =   255
               Left            =   2850
               TabIndex        =   88
               Top             =   3480
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
               TabIndex        =   78
               Top             =   780
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
               TabIndex        =   81
               Top             =   1680
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
               TabIndex        =   76
               Top             =   180
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
               TabIndex        =   82
               Top             =   1980
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
               TabIndex        =   83
               Top             =   2280
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
               Left            =   2850
               TabIndex        =   85
               Top             =   2880
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
               Left            =   2850
               TabIndex        =   86
               Top             =   3180
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
               Left            =   2850
               TabIndex        =   84
               Top             =   2580
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
            Begin MSMask.MaskEdBox txtYTDRemCOLA 
               Height          =   255
               Left            =   2850
               TabIndex        =   79
               Top             =   1080
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
            Begin MSMask.MaskEdBox txtYTDCola 
               Height          =   255
               Left            =   2850
               TabIndex        =   77
               Top             =   480
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
            Begin MSMask.MaskEdBox txtYTDSICKLEAVECONVERSION 
               Height          =   255
               Left            =   2850
               TabIndex        =   80
               Top             =   1380
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
            Begin VB.Label Label74 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "SICK LEAVE CONVERSION"
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
               TabIndex        =   203
               Top             =   1380
               Width           =   2685
            End
            Begin VB.Label Label73 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "COLA"
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
               TabIndex        =   202
               Top             =   480
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
               Left            =   90
               TabIndex        =   201
               Top             =   2580
               Width           =   1845
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
               Left            =   90
               TabIndex        =   200
               Top             =   3180
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
               Left            =   90
               TabIndex        =   199
               Top             =   2880
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
               TabIndex        =   198
               Top             =   2280
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
               TabIndex        =   197
               Top             =   1980
               Width           =   2685
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
               TabIndex        =   196
               Top             =   780
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
               TabIndex        =   195
               Top             =   1680
               Width           =   2445
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
               TabIndex        =   194
               Top             =   180
               Width           =   2175
            End
            Begin VB.Label Label72 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "REMAINING COLA"
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
               TabIndex        =   193
               Top             =   1080
               Width           =   2175
            End
            Begin VB.Label Label63 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "TAX PAYABLE/(TAX REFUND)"
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
               TabIndex        =   55
               Top             =   3510
               Width           =   2685
            End
         End
         Begin VB.Label Label8 
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
            Left            =   90
            TabIndex        =   191
            Top             =   60
            Width           =   1485
         End
      End
      Begin VB.Frame fraDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4185
         Left            =   -74940
         TabIndex        =   133
         Top             =   630
         Width           =   8835
         Begin MSFlexGridLib.MSFlexGrid grdPayroll 
            Height          =   3975
            Left            =   60
            TabIndex        =   134
            Top             =   150
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   7011
            _Version        =   393216
            Cols            =   24
            FixedCols       =   2
            ForeColor       =   0
            ForeColorFixed  =   0
            BackColorSel    =   -2147483633
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            TextStyleFixed  =   3
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
      Begin VB.PictureBox picPayroll 
         Height          =   405
         Left            =   -74760
         ScaleHeight     =   345
         ScaleWidth      =   2880
         TabIndex        =   135
         Top             =   3120
         Width           =   2940
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
            TabIndex        =   139
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
            TabIndex        =   138
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
            TabIndex        =   137
            Top             =   0
            Width           =   735
         End
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
            TabIndex        =   136
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   4125
         Left            =   -74910
         ScaleHeight     =   4065
         ScaleWidth      =   8775
         TabIndex        =   87
         Top             =   690
         Width           =   8835
         Begin VB.Frame fraSSSMED 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1665
            Left            =   30
            TabIndex        =   114
            Top             =   -60
            Width           =   8685
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
               TabIndex        =   116
               Top             =   180
               Width           =   2265
            End
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
               TabIndex        =   115
               Top             =   180
               Width           =   2145
            End
            Begin MSMask.MaskEdBox txtSSSMonthly 
               Height          =   315
               Left            =   2070
               TabIndex        =   117
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
               TabIndex        =   118
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
               TabIndex        =   119
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
               TabIndex        =   120
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
               TabIndex        =   121
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
               TabIndex        =   122
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
               TabIndex        =   130
               Top             =   210
               Width           =   1455
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
               TabIndex        =   129
               Top             =   210
               Width           =   1395
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
               TabIndex        =   128
               Top             =   570
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
               TabIndex        =   127
               Top             =   570
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
               TabIndex        =   126
               Top             =   930
               Width           =   1395
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
               TabIndex        =   125
               Top             =   930
               Width           =   1455
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
               TabIndex        =   124
               Top             =   1290
               Width           =   1965
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
               TabIndex        =   123
               Top             =   1290
               Width           =   1875
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdPhilHealth 
            Height          =   2385
            Left            =   4380
            TabIndex        =   131
            Top             =   1650
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4207
            _Version        =   393216
            Cols            =   4
            ForeColorFixed  =   0
            BackColorSel    =   -2147483633
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            TextStyleFixed  =   3
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
            Height          =   2385
            Left            =   30
            TabIndex        =   132
            Top             =   1650
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   4207
            _Version        =   393216
            Cols            =   4
            ForeColorFixed  =   0
            BackColorSel    =   -2147483633
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            TextStyleFixed  =   3
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
      Begin VB.PictureBox Picture5 
         Height          =   4125
         Left            =   -74910
         ScaleHeight     =   4065
         ScaleWidth      =   8775
         TabIndex        =   56
         Top             =   690
         Width           =   8835
         Begin VB.Frame fraPagIbigTIN 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1665
            Left            =   30
            TabIndex        =   57
            Top             =   -60
            Width           =   8685
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
               TabIndex        =   59
               Top             =   180
               Width           =   2265
            End
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
               TabIndex        =   58
               Top             =   180
               Width           =   2145
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
               TabIndex        =   61
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
               TabIndex        =   62
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
               TabIndex        =   63
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
               TabIndex        =   64
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
               TabIndex        =   65
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
               TabIndex        =   73
               Top             =   210
               Width           =   1395
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
               TabIndex        =   72
               Top             =   210
               Width           =   1365
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
               TabIndex        =   71
               Top             =   570
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
               TabIndex        =   70
               Top             =   570
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
               TabIndex        =   69
               Top             =   930
               Width           =   1365
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
               TabIndex        =   68
               Top             =   930
               Width           =   1395
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
               TabIndex        =   67
               Top             =   1290
               Width           =   1875
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
               TabIndex        =   66
               Top             =   1290
               Width           =   1965
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdPagIbig 
            Height          =   2385
            Left            =   30
            TabIndex        =   74
            Top             =   1650
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   4207
            _Version        =   393216
            Cols            =   4
            ForeColorFixed  =   0
            BackColorSel    =   -2147483633
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            TextStyleFixed  =   3
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
            Height          =   2385
            Left            =   4380
            TabIndex        =   75
            Top             =   1650
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4207
            _Version        =   393216
            Cols            =   3
            ForeColorFixed  =   0
            BackColorSel    =   -2147483633
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            TextStyleFixed  =   3
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
      Height          =   3495
      Left            =   2700
      Picture         =   "Ledger.frx":1550
      ScaleHeight     =   3435
      ScaleWidth      =   8655
      TabIndex        =   140
      Top             =   2460
      Width           =   8715
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
         TabIndex        =   141
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
         TabIndex        =   142
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
         TabIndex        =   143
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
         TabIndex        =   144
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
         TabIndex        =   145
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
         TabIndex        =   146
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
         TabIndex        =   147
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
         Left            =   7140
         TabIndex        =   151
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
      Begin MSMask.MaskEdBox txtTAX 
         Height          =   315
         Left            =   4800
         TabIndex        =   156
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
         TabIndex        =   157
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
         TabIndex        =   158
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
         TabIndex        =   159
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
         TabIndex        =   160
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
         Left            =   7140
         TabIndex        =   150
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
      Begin MSMask.MaskEdBox txtTelBill 
         Height          =   315
         Left            =   7140
         TabIndex        =   152
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
      Begin MSMask.MaskEdBox txtOthers 
         Height          =   315
         Left            =   7140
         TabIndex        =   153
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
      Begin MSMask.MaskEdBox txtNetPay 
         Height          =   315
         Left            =   7140
         TabIndex        =   155
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
         TabIndex        =   161
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
         Left            =   7140
         TabIndex        =   154
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
         TabIndex        =   162
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
         TabIndex        =   163
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
      Begin MSMask.MaskEdBox txtNonTaxableAdj 
         Height          =   315
         Left            =   1590
         TabIndex        =   164
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
      Begin MSMask.MaskEdBox txtPagSal 
         Height          =   315
         Left            =   4800
         TabIndex        =   148
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
      Begin MSMask.MaskEdBox txtHDMF 
         Height          =   315
         Left            =   4800
         TabIndex        =   149
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
         TabIndex        =   188
         Top             =   1590
         Width           =   2265
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
         TabIndex        =   187
         Top             =   1230
         Width           =   1875
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
         TabIndex        =   186
         Top             =   870
         Width           =   1875
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
         Left            =   6180
         TabIndex        =   185
         Top             =   2640
         Width           =   2025
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
         TabIndex        =   184
         Top             =   3000
         Width           =   1995
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
         Left            =   6330
         TabIndex        =   183
         Top             =   3000
         Width           =   1815
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
         TabIndex        =   182
         Top             =   510
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
         TabIndex        =   181
         Top             =   1920
         Width           =   1995
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
         TabIndex        =   180
         Top             =   300
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
         TabIndex        =   179
         Top             =   690
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
         TabIndex        =   178
         Top             =   1080
         Width           =   1455
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
         TabIndex        =   177
         Top             =   1860
         Width           =   1455
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
         TabIndex        =   176
         Top             =   2250
         Width           =   1605
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
         Left            =   6420
         TabIndex        =   175
         Top             =   660
         Width           =   1785
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
         TabIndex        =   174
         Top             =   1470
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
         TabIndex        =   173
         Top             =   120
         Width           =   1365
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
         TabIndex        =   172
         Top             =   90
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   3120
         Y1              =   435
         Y2              =   435
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
         TabIndex        =   171
         Top             =   2280
         Width           =   1875
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
         TabIndex        =   170
         Top             =   2670
         Width           =   1995
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   3120
         X2              =   3120
         Y1              =   -30
         Y2              =   3420
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
         Left            =   6360
         TabIndex        =   169
         Top             =   300
         Width           =   1845
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
         Left            =   6390
         TabIndex        =   168
         Top             =   1080
         Width           =   1815
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
         Left            =   6450
         TabIndex        =   167
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label Label51 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Pag-Ibig HDMF"
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
         TabIndex        =   166
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label50 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PAG Sal. Loan"
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
         TabIndex        =   165
         Top             =   2640
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   2640
      Picture         =   "Ledger.frx":428C
      ScaleHeight     =   1125
      ScaleWidth      =   7185
      TabIndex        =   17
      Top             =   360
      Width           =   7215
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
         TabIndex        =   19
         Top             =   600
         Width           =   6015
      End
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
         TabIndex        =   18
         Top             =   150
         Width           =   6015
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
         TabIndex        =   22
         Top             =   270
         Width           =   615
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
         TabIndex        =   21
         Top             =   600
         Width           =   885
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
         TabIndex        =   20
         Top             =   150
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture7 
      Height          =   1155
      Left            =   10050
      ScaleHeight     =   1095
      ScaleWidth      =   1245
      TabIndex        =   16
      Top             =   360
      Width           =   1305
      Begin VB.Image imgDispPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1065
         Left            =   30
         Stretch         =   -1  'True
         Top             =   30
         Width           =   1185
      End
   End
   Begin VB.PictureBox picPayroll2 
      Height          =   405
      Left            =   6180
      ScaleHeight     =   345
      ScaleWidth      =   1440
      TabIndex        =   12
      Top             =   4260
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
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   0
         Width           =   735
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
      Left            =   2880
      Picture         =   "Ledger.frx":6545
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4350
      Width           =   1185
   End
   Begin VB.PictureBox wizMacApp1 
      Height          =   320
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   11625
      TabIndex        =   190
      Top             =   0
      Width           =   11685
   End
   Begin wizButton.cmd cmdPayroll 
      Height          =   3615
      Left            =   2640
      TabIndex        =   189
      Top             =   2400
      Width           =   8835
      _ExtentX        =   15584
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
      MICON           =   "Ledger.frx":6987
   End
   Begin VB.PictureBox Picture11 
      Height          =   7005
      Left            =   60
      Picture         =   "Ledger.frx":69A3
      ScaleHeight     =   6945
      ScaleWidth      =   2445
      TabIndex        =   23
      Top             =   360
      Width           =   2505
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   60
      Picture         =   "Ledger.frx":1A700
      ScaleHeight     =   6945
      ScaleWidth      =   2475
      TabIndex        =   24
      Top             =   390
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
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   60
         Width           =   2415
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   6465
         Left            =   0
         TabIndex        =   26
         Top             =   450
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   11404
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
         MouseIcon       =   "Ledger.frx":1D43C
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
         Picture         =   "Ledger.frx":1D59E
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2640
      ScaleHeight     =   855
      ScaleWidth      =   8895
      TabIndex        =   27
      Top             =   6510
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
         MouseIcon       =   "Ledger.frx":3130B
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":3145D
         Style           =   1  'Graphical
         TabIndex        =   7
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
         MouseIcon       =   "Ledger.frx":31767
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":318B9
         Style           =   1  'Graphical
         TabIndex        =   6
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
         MouseIcon       =   "Ledger.frx":32183
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":322D5
         Style           =   1  'Graphical
         TabIndex        =   5
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
         MouseIcon       =   "Ledger.frx":32B9F
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":32CF1
         Style           =   1  'Graphical
         TabIndex        =   4
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
         MouseIcon       =   "Ledger.frx":335BB
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":3370D
         Style           =   1  'Graphical
         TabIndex        =   3
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
         MouseIcon       =   "Ledger.frx":33FD7
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":34129
         Style           =   1  'Graphical
         TabIndex        =   2
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
         MouseIcon       =   "Ledger.frx":349F3
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":34B45
         Style           =   1  'Graphical
         TabIndex        =   1
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
         MouseIcon       =   "Ledger.frx":34F87
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":350D9
         Style           =   1  'Graphical
         TabIndex        =   0
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
      ScaleHeight     =   855
      ScaleWidth      =   8895
      TabIndex        =   28
      Top             =   6510
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
         MouseIcon       =   "Ledger.frx":3551B
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":3566D
         Style           =   1  'Graphical
         TabIndex        =   9
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
         MouseIcon       =   "Ledger.frx":366AF
         MousePointer    =   99  'Custom
         Picture         =   "Ledger.frx":36801
         Style           =   1  'Graphical
         TabIndex        =   8
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
      Left            =   2760
      TabIndex        =   11
      Top             =   570
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
      Left            =   3810
      TabIndex        =   10
      Top             =   1050
      Width           =   255
   End
End
Attribute VB_Name = "frmHRMSLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo, rsPAYROLL, rsDeductions                                As ADODB.Recordset
Attribute rsPAYROLL.VB_VarUserMemId = 1073938432
Attribute rsDeductions.VB_VarUserMemId = 1073938432
Dim rsLoanMas, rsCalamityLoan, rsLoanmasDet                           As ADODB.Recordset
Attribute rsLoanMas.VB_VarUserMemId = 1073938435
Attribute rsCalamityLoan.VB_VarUserMemId = 1073938435
Attribute rsLoanmasDet.VB_VarUserMemId = 1073938435
Dim rsSSS, rsSSSdet, rsPH, rsPHDet                                    As ADODB.Recordset
Attribute rsSSS.VB_VarUserMemId = 1073938438
Attribute rsSSSdet.VB_VarUserMemId = 1073938438
Attribute rsPH.VB_VarUserMemId = 1073938438
Attribute rsPHDet.VB_VarUserMemId = 1073938438
Dim rsPagIbig, rsTIN, rsPagibigdet                                    As ADODB.Recordset
Attribute rsPagIbig.VB_VarUserMemId = 1073938442
Attribute rsTIN.VB_VarUserMemId = 1073938442
Attribute rsPagibigdet.VB_VarUserMemId = 1073938442
Dim rsTINdet, rsYTDDETAILS                                            As ADODB.Recordset
Attribute rsTINdet.VB_VarUserMemId = 1073938445
Attribute rsYTDDETAILS.VB_VarUserMemId = 1073938445
Dim AddorEdit, CLID                                                   As String
Attribute AddorEdit.VB_VarUserMemId = 1073938447
Attribute CLID.VB_VarUserMemId = 1073938447
Dim ToBeVat                                                           As Double
Attribute ToBeVat.VB_VarUserMemId = 1073938449
Dim EMPLIVIL                                                          As String
Attribute EMPLIVIL.VB_VarUserMemId = 1073938450
Dim YEAR_TO_DATE_SSS, YEAR_TO_DATE_PHIC, YEAR_TO_DATE_PAGIBIG         As Double
Attribute YEAR_TO_DATE_SSS.VB_VarUserMemId = 1073938451
Attribute YEAR_TO_DATE_PHIC.VB_VarUserMemId = 1073938451
Attribute YEAR_TO_DATE_PAGIBIG.VB_VarUserMemId = 1073938451

Function MAIDIT(XXX As String)
    AddorEdit = "EDIT"
    cmdPayroll.ZOrder 0
    picAddEditPayroll.ZOrder 0
    txtFrom.Enabled = True
    txtTo.Enabled = True
    MAHIRO
    StoreEntry XXX
End Function

Function StoreEntry(ByVal ID As Variant)
    Set rsPAYROLL = New ADODB.Recordset
    rsPAYROLL.Open "select * from HRMS_Payroll where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        labPayrollID.Caption = rsPAYROLL!ID
        txtFrom.Text = Null2Date(rsPAYROLL!paydatefrom)
        txtTo.Text = Null2Date(rsPAYROLL!paydateto)
        txtRate.Text = N2Str2Zero(rsPAYROLL!Rate)
        txtDailyRate.Text = N2Str2Zero(rsPAYROLL!DailyRate)
        txtTaxableAdj.Text = N2Str2IntZero(rsPAYROLL!TAXABLEADJ)
        txtNonTaxableAdj.Text = N2Str2IntZero(rsPAYROLL!NONTAXABLEADJ)
        txtOvertime.Text = N2Str2Zero(rsPAYROLL!OVERTIME)
        txtHoliday.Text = N2Str2Zero(rsPAYROLL!HOLIDAY)
        txtCommission.Text = N2Str2Zero(rsPAYROLL!commission)
        txtSSS.Text = N2Str2Zero(rsPAYROLL!SSSE)
        txtMed.Text = N2Str2Zero(rsPAYROLL!PHILHEALTHE)
        txtPagIbig.Text = N2Str2Zero(rsPAYROLL!PAGIBIG)
        txtTAX.Text = N2Str2Zero(rsPAYROLL!TAX)
        txtSalLoan.Text = N2Str2Zero(rsPAYROLL!SSSSALLOAN)
        txtCalLoan.Text = N2Str2Zero(rsPAYROLL!SSSCALLOAN)
        txtPagSal.Text = N2Str2Zero(rsPAYROLL!PAGSALLOAN)
        txtHDMF.Text = N2Str2Zero(rsPAYROLL!PAGHDMFLOAN)
        txtUndertime.Text = N2Str2Zero(rsPAYROLL!UNDERTIME)
        txtAbsent.Text = N2Str2Zero(rsPAYROLL!ABSENT)
        txtTelBill.Text = N2Str2Zero(rsPAYROLL!telbill)
        txtOthers.Text = N2Str2Zero(rsPAYROLL!Others)
        showGrossWage
    End If
End Function

Function SetSalary() As Double
    Dim rsEmpInformation                                              As ADODB.Recordset
    Dim rsSalaryGrade                                                 As ADODB.Recordset

    Set rsEmpInformation = New ADODB.Recordset
    Set rsEmpInformation = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & rsEmpInfo!EMPNO & "'")
    If Not rsEmpInformation.EOF And Not rsEmpInformation.BOF Then
        Set rsSalaryGrade = New ADODB.Recordset
        rsSalaryGrade.Open "select code,salary from HRMS_salarygrade where code = '" & rsEmpInformation!SalaryCode & "'", gconDMIS
        If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
            SetSalary = Round(N2Str2Zero(rsSalaryGrade!SALARY), 2)
        End If
    End If
End Function

Function FindLoanType(LOAN_ID As String) As String
    Dim RSTMP                                                         As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Description From HRMS_LoanCode Where Code = '" & LOAN_ID & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindLoanType = RSTMP!Description
    End If

    Set RSTMP = Nothing
End Function

Function StoreDed(Diyt As String)
    Set rsPAYROLL = New ADODB.Recordset
    rsPAYROLL.Open "select * from HRMS_Payroll where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & IMPNO & "' and paydateto = '" & Format(Diyt, "Short Date") & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        MAIDIT rsPAYROLL!ID
    End If
End Function

Sub MAHIRO()
    EnAbleFrames
    TabSSS.TabEnabled(0) = False
    Picture1.Visible = False
    Picture2.Visible = True
    cmdPayroll.Visible = True
    picPayroll.Visible = True
    picAddEditPayroll.Visible = True
    picAddEditPayroll.Enabled = True
    picPayroll2.Visible = True
End Sub

Sub showGrossWage()
    Dim GrossWeyg                                                     As Double
    GrossWeyg = NumericVal(txtRate.Text) + NumericVal(txtTaxableAdj.Text) + NumericVal(txtOvertime.Text) + NumericVal(txtHoliday.Text)
    txtGrossWage.Text = NumericVal(txtRate.Text) + NumericVal(txtTaxableAdj.Text) + NumericVal(txtNonTaxableAdj.Text) + NumericVal(txtOvertime.Text) + NumericVal(txtHoliday.Text)
    If WIZVAR.EncryptAccess(LOGNAME) = "77697A∂ùº∞û¡¡¢¡" Then
        txtTAX.Text = TaxDedSemiMonthly(GrossWeyg, Null2String(rsEmpInfo!ExStatus))
        If Day(txtFrom.Text) < 15 Then
            txtPagIbig.Text = PagIbigShare(NumericVal(txtRate.Text) * 2)
        Else
            txtSSS.Text = EmployeeSSSshare(NumericVal(txtRate.Text) * 2)
            txtMed.Text = PhilHealthShare(NumericVal(txtRate.Text) * 2)
        End If
    End If
    If AddorEdit = "ADD" Then
        txtTAX.Text = TaxDedSemiMonthly(GrossWeyg, Null2String(rsEmpInfo!ExStatus))
        If Day(txtFrom.Text) < 15 Then
            txtPagIbig.Text = PagIbigShare(NumericVal(txtRate.Text) * 2)
        Else
            txtSSS.Text = EmployeeSSSshare(NumericVal(txtRate.Text) * 2)
            txtMed.Text = PhilHealthShare(NumericVal(txtRate.Text) * 2)
        End If
    End If
    ShowTotDed
    ShowNetPay
End Sub

Sub ShowTotDed()
    txtTotDed.Text = NumericVal(txtSSS.Text) + NumericVal(txtMed.Text) + NumericVal(txtPagIbig.Text) + NumericVal(txtTAX.Text) + NumericVal(txtSalLoan.Text) + NumericVal(txtCalLoan.Text) + NumericVal(txtPagSal.Text) + NumericVal(txtHDMF.Text) + NumericVal(txtUndertime.Text) + NumericVal(txtAbsent.Text) + NumericVal(txtTelBill.Text) + NumericVal(txtOthers.Text)
    ShowNetPay
End Sub

Sub ShowNetPay()
    txtNetPay.Text = NumericVal(txtGrossWage.Text) - NumericVal(txtTotDed.Text)
End Sub

Sub DelEXIST()
    gconDMIS.Execute "delete from HRMS_Payroll where (paydatefrom >= '" & GENFROM & "')" & _
                   " AND (paydateto <= '" & GENTO & "') and empno = '" & IMPNO & "' and EMPLEVEL = " & EMPLIVIL
    gconDMIS.Execute "delete from HRMS_SSSdet where (deyt = '" & GENTO & "') and empno = '" & IMPNO & "' and EMPLEVEL = " & EMPLIVIL
    gconDMIS.Execute "delete from HRMS_loanmasDet where (deyt = '" & GENTO & "') and empno = '" & IMPNO & "' and EMPLEVEL = " & EMPLIVIL
    gconDMIS.Execute "delete from HRMS_philhealthdet where (deyt = '" & GENTO & "') and empno = '" & IMPNO & "' and EMPLEVEL = " & EMPLIVIL
    gconDMIS.Execute "delete from HRMS_pagibigdet where (deyt = '" & GENTO & "') and empno = '" & IMPNO & "' and EMPLEVEL = " & EMPLIVIL
    gconDMIS.Execute "delete from HRMS_tindet where (deyt = '" & GENTO & "') and empno = '" & IMPNO & "' and EMPLEVEL = " & EMPLIVIL
    gconDMIS.Execute "delete from HRMS_atmdet where (deyt = '" & GENTO & "') and empno = '" & IMPNO & "' and EMPLEVEL = " & EMPLIVIL

    LogAudit "X", "DELETE PAYROLL", IMPNO & "-" & GENTO
End Sub

Sub SEYV()
    On Error GoTo Errorcode
    Dim rsPrevPayroll, rsAllPrevPayroll, rsCommission                 As ADODB.Recordset
    Dim rsEmpInfoClone                                                As ADODB.Recordset
    Dim SUWELDO, SUWELDOKINSE                                         As Double
    Dim dedPAGIBIG, dedEmpPAGIBIG                                     As Double
    Dim dedTIN, dedSSS                                                As Double
    Dim dedEmpSSS, dedPhilHealth                                      As Double
    Dim dedEmpPhilhealth, dedSalLoan                                  As Double
    Dim dedCalLoan, dedPagSalLoan                                     As Double
    Dim dedHDMFLoan, SalGross                                         As Double
    Dim TotUndertime, TotAbsent                                       As Double
    Dim TotTelBill, TotOthers                                         As Double
    Dim TotOvertime, TotTaxableAdj                                    As Double
    Dim TotNonTaxableAdj, TotHoliday                                  As Double
    Dim TotCommission, TotCommissionTax                               As Double
    Dim I, CNT                                                        As Integer
    Dim amt, NUMDAYS, DEYLI, BULANAN, SUWELDOTRIENTA                  As Double
    Dim VARPAYSTATUS                                                  As String
    VARPAYSTATUS = "P"
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
    dedPagSalLoan = NumericVal(txtPagSal.Text)
    dedHDMFLoan = NumericVal(txtHDMF.Text)
    dedPAGIBIG = NumericVal(txtPagIbig.Text)
    dedTIN = NumericVal(txtTAX.Text)
    TotUndertime = NumericVal(txtUndertime.Text)
    TotAbsent = NumericVal(txtAbsent.Text)
    TotTelBill = NumericVal(txtTelBill.Text)
    TotOthers = NumericVal(txtOthers.Text)
    SalGross = NumericVal(txtGrossWage.Text)
    dedEmpPhilhealth = 0
    dedEmpPAGIBIG = 0
    dedEmpSSS = 0
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(IMPNO), gconDMIS
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        If Null2String(rsEmpInfo!EMPSTATUS) = "M" Then
            NUMDAYS = 0
            SUWELDOTRIENTA = N2Str2Zero(rsEmpInfo!SALARY)
        Else
            SUWELDOTRIENTA = SUWELDOKINSE
            NUMDAYS = SUWELDOKINSE / DEYLI
            If Day(GENFROM) > 15 Then
                Set rsPrevPayroll = New ADODB.Recordset
                rsPrevPayroll.Open "select empno,paydatefrom,rate from HRMS_Payroll where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " and paydatefrom = '" & CDate(firstDay(GENFROM)) & "'", gconDMIS
                If Not rsPrevPayroll.EOF And Not rsPrevPayroll.BOF Then
                    SUWELDOTRIENTA = SUWELDOKINSE + N2Str2Zero(rsPrevPayroll!Rate)
                End If
            End If
        End If
        Set rsCommission = New ADODB.Recordset
        rsCommission.Open "select * from Commission where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND " & _
                          "(deyt >= '" & Format(GENFROM, "Short Date") & "')" & _
                        " AND (deyt <= '" & Format(GENTO, "Short Date") & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        TotCommissionTax = 0
        If Not rsCommission.EOF And Not rsCommission.BOF Then
            rsCommission.MoveFirst
            Do While Not rsCommission.EOF
                TotCommissionTax = TotCommissionTax + N2Str2Zero(rsCommission!TAX)
                rsCommission.MoveNext
            Loop
        End If
        Set rsLoanMas = New ADODB.Recordset
        rsLoanMas.Open "select * from HRMS_loanmas where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND (dategranted <= '" & Format(GENTO, "Short Date") & "') AND (maturitydate >= '" & Format(GENTO, "Short Date") & "') order by dategranted desc", gconDMIS
        If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
            rsLoanMas.MoveFirst
            Do While Not rsLoanMas.EOF
                If N2Str2Zero(rsLoanMas!LoanBalance) > 0 Then
                    If Null2String(rsLoanMas!LOANTYPE) = "SSAL" Then
                        gconDMIS.Execute "insert into LoanMasDet " & _
                                         "(EMPLEVEL,empno,acctno,amount,deyt,loantype)" & _
                                       " values (" & EMPLIVIL & "," & N2Str2Null(rsLoanMas!EMPNO) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedSalLoan) & ", '" & GENTO & "', 'SSAL')"
                        gconDMIS.Execute "update loanmas set " & _
                                       " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedSalLoan) & _
                                       " where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsLoanMas!EMPNO)
                    End If
                    If Null2String(rsLoanMas!LOANTYPE) = "CSAL" Then
                        gconDMIS.Execute "insert into LoanMasDet " & _
                                         "(EMPLEVEL,empno,acctno,amount,deyt,loantype)" & _
                                       " values (" & EMPLIVIL & "," & N2Str2Null(rsLoanMas!EMPNO) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedCalLoan) & ", '" & GENTO & "','CSAL')"
                        gconDMIS.Execute "update loanmas set " & _
                                       " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedPagSalLoan) & _
                                       " where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsLoanMas!EMPNO)
                    End If
                    If Null2String(rsLoanMas!LOANTYPE) = "PSAL" Then
                        gconDMIS.Execute "insert into LoanMasDet " & _
                                         "(EMPLEVEL,empno,acctno,amount,deyt,loantype)" & _
                                       " values (" & EMPLIVIL & "," & N2Str2Null(rsLoanMas!EMPNO) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedPagSalLoan) & ", '" & GENTO & "','PSAL')"
                        gconDMIS.Execute "update loanmas set " & _
                                       " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedPagSalLoan) & _
                                       " where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsLoanMas!EMPNO)
                    End If
                    If Null2String(rsLoanMas!LOANTYPE) = "HDMF" Then
                        gconDMIS.Execute "insert into LoanMasDet " & _
                                         "(EMPLEVEL,empno,acctno,amount,deyt,loantype)" & _
                                       " values (" & EMPLIVIL & "," & N2Str2Null(rsLoanMas!EMPNO) & ", " & N2Str2Null(rsLoanMas!acctno) & ", " & N2Str2Zero(dedHDMFLoan) & ", '" & GENTO & "','HDMF')"
                        gconDMIS.Execute "update loanmas set " & _
                                       " loanbalance = " & N2Str2Zero(rsLoanMas!LoanBalance) - N2Str2Zero(dedHDMFLoan) & _
                                       " where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsLoanMas!EMPNO)
                    End If
                End If
                rsLoanMas.MoveNext
            Loop
        End If
        If Day(GENTO) > 15 Then
            dedEmpSSS = EmployerSSSshare(N2Str2Zero(SUWELDOTRIENTA))
            Set rsSSS = New ADODB.Recordset
            rsSSS.Open "select * from HRMS_SSS where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If rsSSS.EOF And rsSSS.BOF Then
                gconDMIS.Execute "insert into sss " & _
                                 "(EMPLEVEL,empno,sssno,datestart,employeeshare,employershare,lastdatecont)" & _
                               " values (" & EMPLIVIL & "," & N2Str2Null(rsEmpInfo!EMPNO) & ", " & N2Str2Null(rsEmpInfo!SSSNO) & ", '" & GENTO & "'," & _
                               " " & dedSSS & ", " & dedEmpSSS & ", '" & GENTO & "')"
            Else
                gconDMIS.Execute "update sss set" & _
                               " employeeshare = " & dedSSS & "," & _
                               " employershare = " & dedEmpSSS & "," & _
                               " lastdatecont = '" & GENTO & "'" & _
                               " where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO)
            End If
            Set rsSSS = New ADODB.Recordset
            rsSSS.Open "select * from HRMS_SSS where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsSSS.EOF And Not rsSSS.BOF Then
                gconDMIS.Execute "insert into sssdet " & _
                                 "(EMPLEVEL,aydi,deyt,empno,employeeamount,employeramount)" & _
                               " values (" & EMPLIVIL & "," & rsSSS!aydi & ", '" & GENTO & "', " & N2Str2Null(rsEmpInfo!EMPNO) & ", " & dedSSS & ", " & dedEmpSSS & ")"
            End If
            dedEmpPhilhealth = PhilHealthShare(N2Str2Zero(SUWELDOTRIENTA))
            Set rsPH = New ADODB.Recordset
            rsPH.Open "select * from HRMS_philhealth where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If rsPH.EOF And rsPH.BOF Then
                gconDMIS.Execute "insert into philhealth " & _
                                 "(EMPLEVEL,empno,phno,datestart,employeeshare,employershare,lastdatecont)" & _
                               " values (" & EMPLIVIL & "," & N2Str2Null(rsEmpInfo!EMPNO) & ", " & N2Str2Null(rsEmpInfo!SSSNO) & ", '" & GENTO & "'," & _
                               " " & dedPhilHealth & ", " & dedEmpPhilhealth & ", '" & GENTO & "')"
            Else
                gconDMIS.Execute "update philhealth set" & _
                               " employeeshare = " & dedPhilHealth & "," & _
                               " employershare = " & dedEmpPhilhealth & "," & _
                               " lastdatecont = '" & GENTO & "'" & _
                               " where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO)
            End If
            Set rsPH = New ADODB.Recordset
            rsPH.Open "select * from HRMS_philhealth where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPH.EOF And Not rsPH.BOF Then
                gconDMIS.Execute "insert into philhealthdet " & _
                                 "(EMPLEVEL,aydi,deyt,empno,employeeamount,employeramount)" & _
                               " values (" & EMPLIVIL & "," & rsPH!aydi & ", '" & GENTO & "', " & N2Str2Null(rsEmpInfo!EMPNO) & ", " & dedPhilHealth & ", " & dedEmpPhilhealth & ")"
            End If
        Else
            Set rsPagIbig = New ADODB.Recordset
            rsPagIbig.Open "select * from HRMS_pagibig where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If rsPagIbig.EOF And rsPagIbig.BOF Then
                dedEmpPAGIBIG = (PagIbigShare(N2Str2Zero(rsEmpInfo!SALARY)))
                gconDMIS.Execute "insert into PagIbig " & _
                                 "(EMPLEVEL,empno,pagibigno,datestart,employeeshare,employershare,lastdatecont)" & _
                               " values (" & EMPLIVIL & "," & N2Str2Null(rsEmpInfo!EMPNO) & ", " & N2Str2Null(rsEmpInfo!SSSNO) & ", '" & GENTO & "'," & _
                               " " & dedPAGIBIG & ", " & dedEmpPAGIBIG & ", '" & GENTO & "')"
            Else
                dedEmpPAGIBIG = (PagIbigShare(N2Str2Zero(rsEmpInfo!SALARY)))
                gconDMIS.Execute "update pagibig set" & _
                               " employeeshare = " & dedPAGIBIG & "," & _
                               " employershare = " & dedEmpPAGIBIG & "," & _
                               " lastdatecont = '" & GENTO & "'" & _
                               " where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO)
                gconDMIS.Execute "insert into pagibigdet " & _
                                 "(EMPLEVEL,aydi,deyt,empno,employeeamount,employeramount)" & _
                               " values (" & EMPLIVIL & "," & rsPagIbig!aydi & ", '" & GENTO & "', " & N2Str2Null(rsEmpInfo!EMPNO) & ", " & dedPAGIBIG & ", " & dedEmpPAGIBIG & ")"
            End If
        End If
        Set rsTIN = New ADODB.Recordset
        rsTIN.Open "select * from HRMS_tin where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If rsTIN.EOF And rsTIN.BOF Then
            gconDMIS.Execute "insert into TIN " & _
                             "(EMPLEVEL,empno,tinno,datestart,deduction,lastdatecont)" & _
                           " values (" & EMPLIVIL & "," & N2Str2Null(rsEmpInfo!EMPNO) & ", " & N2Str2Null(rsEmpInfo!tinno) & ", '" & GENTO & "'," & _
                           " " & dedTIN & ", '" & GENTO & "')"
        Else
            gconDMIS.Execute "update TIN set" & _
                           " deduction = " & dedTIN & "," & _
                           " lastdatecont = '" & GENTO & "'" & _
                           " where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO)
            gconDMIS.Execute "insert into tindet " & _
                             "(EMPLEVEL,aydi,empno,deyt,amount)" & _
                           " values (" & EMPLIVIL & "," & rsTIN!aydi & ", " & N2Str2Null(rsEmpInfo!EMPNO) & ", '" & GENTO & "', " & dedTIN & ")"
        End If
        SUWELDO = (SUWELDOKINSE + TotOvertime + TotHoliday + TotTaxableAdj + TotNonTaxableAdj) - (dedPhilHealth + dedSSS + dedPAGIBIG + dedTIN + dedSalLoan + dedCalLoan + dedPagSalLoan + dedHDMFLoan + TotUndertime + TotTelBill + TotAbsent + TotOthers)
        gconDMIS.Execute "insert into payroll " & _
                         "(EMPLEVEL,empno,taxcode,rate,monthlyrate,DailyRate,ndays,overtime,holiday,commission,commissiontax,taxableadj,nontaxableadj,gross,undertime,sssE,sssR,philhealthE,philhealthR,pagibig,tax,ssssalloan,ssscalloan,pagsalloan,paghdmfloan,absent,telbill,others,paydatefrom,paydateto,netpay,payrollstatus)" & _
                       " values (" & EMPLIVIL & "," & N2Str2Null(rsEmpInfo!EMPNO) & ", " & N2Str2Null(rsEmpInfo!ExStatus) & ", " & (SUWELDOKINSE) & _
                         ", " & N2Str2Zero(SUWELDOTRIENTA) & ", " & N2Str2Zero(DEYLI) & ", " & NUMDAYS & _
                         ", " & TotOvertime & ", " & TotHoliday & ", " & TotCommission & ", " & TotCommissionTax & ", " & TotTaxableAdj & ", " & TotNonTaxableAdj & _
                         ", " & (SUWELDOKINSE) + TotOvertime + TotHoliday + TotTaxableAdj + TotNonTaxableAdj & _
                         ", " & TotUndertime & ", " & dedSSS & ", " & dedEmpSSS & _
                         ", " & dedPhilHealth & ", " & dedEmpPhilhealth & ", " & dedPAGIBIG & ", " & dedTIN & ", " & dedSalLoan & ", " & dedCalLoan & ", " & dedPagSalLoan & ", " & dedHDMFLoan & _
                         ", " & TotAbsent & ", " & TotTelBill & ", " & TotOthers & _
                         ", '" & GENFROM & "', '" & GENTO & "', " & SUWELDO & ", '" & VARPAYSTATUS & "')"
        gconDMIS.Execute "insert into atmdet " & _
                         "(EMPLEVEL,acctno,empno,atmid,deyt,netamount) " & _
                         "values (" & EMPLIVIL & "," & N2Str2Null(rsEmpInfo!ACCOUNTNO) & ", " & N2Str2Null(rsEmpInfo!EMPNO) & ", " & rsEmpInfo!ID & _
                         ", '" & GENTO & "', " & SUWELDO & ")"
    End If
    rsrefresh
    'on error resume next
    rsEmpInfo.Find "id = " & LabID.Caption
    If EMPINFOSHOW = True Then
        frmHRMSEmpInfo.rsrefresh
        frmHRMSEmpInfo.StoreMemVars
    End If
    Screen.MousePointer = 0

    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Sub FillCboLoanType()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim SPACES                                                        As String

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_LoanCode Order By Description ASC")
    cboLoanType.Clear
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        Do While Not RSTMP.EOF
            If Len(RSTMP!CODE) = 3 Then SPACES = " "
            If Len(RSTMP!CODE) = 2 Then SPACES = "  "
            If Len(RSTMP!CODE) = 1 Then SPACES = "   "
            cboLoanType.AddItem Null2String(RSTMP!Description) & " - " & SPACES & RSTMP!CODE

            RSTMP.MoveNext
        Loop
        cboLoanType.ListIndex = 0
    End If
    Set RSTMP = Nothing
End Sub

Sub DisAbleFrames()
    fraLoanMas.Enabled = False
    fraSSSMED.Enabled = False
    fraPagIbigTIN.Enabled = False
    'fraYTD.Enabled = False
End Sub

Sub EnAbleFrames()
    fraLoanMas.Enabled = True
    fraSSSMED.Enabled = True
    fraPagIbigTIN.Enabled = True
End Sub

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & EMPINFOEMPNO.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & frmHRMSEmpInfo.LabID.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " order by lastname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
End Sub

Sub StoreMemVars()
    On Error GoTo Errorcode
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Screen.MousePointer = 11
        DoEvents
        LabID.Caption = rsEmpInfo!ID
        txtName.Text = RTrim(rsEmpInfo!lastname) + ", " + RTrim(rsEmpInfo!FIRSTNAME) + " " + RTrim(rsEmpInfo!MIDDLENAME)
        IMPNO = Null2String(rsEmpInfo!EMPNO)
        txtPosition.Text = Null2String(rsEmpInfo!Position)
        'txtSalary.Text = Null2String(rsEmpInfo!Salary)
        txtSSSNo.Text = Null2String(rsEmpInfo!SSSNO)
        txtPhilHealthNo.Text = Null2String(rsEmpInfo!phno)
        txtPagIbigNo.Text = Null2String(rsEmpInfo!pagibigno)
        txtTINNo.Text = Null2String(rsEmpInfo!tinno)
        If Null2String(rsEmpInfo!PICFILNAME) <> "" Then
            'on error resume next
            'LoadPic imgDispPic, HRMS_PICTURES_PATH & Null2String(rsEMPINFO!PICFILNAME)
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

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub StoreYTD()
    InitYTDMemVars
    Set rsYTDDETAILS = New ADODB.Recordset
    rsYTDDETAILS.Open "select * from HRMS_ytddetails where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & rsEmpInfo!EMPNO & "' and yeer = '" & cboYear.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
        txtYTDCola.Text = N2Str2Zero(rsYTDDETAILS!ytdcola)
        txtYTDRemSal.Text = N2Str2Zero(rsYTDDETAILS!remsal)
        txtYTDRemCOLA.Text = N2Str2Zero(rsYTDDETAILS!RemCOLA)
        txtYTDSICKLEAVECONVERSION.Text = N2Str2Zero(rsYTDDETAILS!YTDSICKLEAVECONVERSION)

        YEAR_TO_DATE_SSS = N2Str2Zero(rsYTDDETAILS!YTDSSS)
        YEAR_TO_DATE_PHIC = N2Str2Zero(rsYTDDETAILS!YTDPHIC)
        YEAR_TO_DATE_PAGIBIG = N2Str2Zero(rsYTDDETAILS!YTDPAGIBIG)
        txtYTDSSSPAGIBIGPHIC.Text = YEAR_TO_DATE_SSS + YEAR_TO_DATE_PHIC + YEAR_TO_DATE_PAGIBIG + N2Str2Zero(rsYTDDETAILS!DED_SSSPREMIUM) + N2Str2Zero(rsYTDDETAILS!DED_PHICPREMIUM)
        txtTaxExemp.Text = N2Str2Zero(rsYTDDETAILS!PersonalEx)

        txtYTDBasicPay.Text = N2Str2Zero(rsYTDDETAILS!ytdbasicpay) + NumericVal(txtYTDSSSPAGIBIGPHIC.Text)
        txtYTDGross.Text = NumericVal(txtYTDBasicPay.Text) + NumericVal(txtYTDCola.Text) + NumericVal(txtYTDRemCOLA.Text) + NumericVal(txtYTDRemSal.Text)

        txtYTDTaxWithHeld.Text = N2Str2Zero(rsYTDDETAILS!ytdtax) + N2Str2Zero(rsYTDDETAILS!decytdtax) + N2Str2Zero(rsYTDDETAILS!commissiontax) + N2Str2Zero(rsYTDDETAILS!FP_REMAININGTAXWITHHELD)

        txtFP_VACATIONLEAVE.Text = N2Str2Zero(rsYTDDETAILS!FP_VACATIONLEAVE)
        txtFP_SICKLEAVE.Text = N2Str2Zero(rsYTDDETAILS!FP_SICKLEAVE)
        txtFP_REMAININGSALARY.Text = N2Str2Zero(rsYTDDETAILS!remsal)
        txtFP_REMAININGOVERTIME.Text = N2Str2Zero(rsYTDDETAILS!FP_REMAININGOVERTIME)
        txtFP_REMAININGCOLA.Text = N2Str2Zero(rsYTDDETAILS!RemCOLA)
        txtFP_REMAININGCOMMISSION.Text = N2Str2Zero(rsYTDDETAILS!FP_REMAININGCOMMISSION)
        txtFP_REMAININGTAXWITHHELD.Text = N2Str2Zero(rsYTDDETAILS!FP_REMAININGTAXWITHHELD)

        txtDED_SSSPREMIUM.Text = N2Str2Zero(rsYTDDETAILS!DED_SSSPREMIUM)
        txtDED_PHICPREMIUM.Text = N2Str2Zero(rsYTDDETAILS!DED_PHICPREMIUM)
        txtDED_ABSENT.Text = N2Str2Zero(rsYTDDETAILS!DED_ABSENT)
        txtDED_UNDERTIME.Text = N2Str2Zero(rsYTDDETAILS!DED_UNDERTIME)
        txtDED_OTHERS.Text = N2Str2Zero(rsYTDDETAILS!DED_OTHERS)

        txtSUMMARY_BASICPAY.Text = N2Str2Zero(rsYTDDETAILS!ytdbasicpay) + YEAR_TO_DATE_SSS + YEAR_TO_DATE_PHIC + YEAR_TO_DATE_PAGIBIG
        txtSUMMARY_COMMISSION.Text = N2Str2Zero(rsYTDDETAILS!commission)
        txtSUMMARY_OVERTIME.Text = N2Str2Zero(rsYTDDETAILS!OVERTIME)

        txtTaxableAdjustment.Text = N2Str2Zero(rsYTDDETAILS!TAXABLEADJ)
        txtSUMMARY_ADJUSTEDSALARY.Text = N2Str2Zero(rsYTDDETAILS!remsal)
        txtSUMMARY_MIDYEAR.Text = N2Str2Zero(rsYTDDETAILS!midyear)
        txtSUMMARY_13THMONTHPAY.Text = N2Str2Zero(rsYTDDETAILS!t13thmonth)
        txtSUMMARY_BONUS.Text = N2Str2Zero(rsYTDDETAILS!bonus)
        ShowYTDDetails
    Else
        InitYTDMemVars
    End If
End Sub

Sub InitYTDMemVars()
    txtYTDBasicPay.Text = 0
    txtYTDCola.Text = 0
    txtYTDRemSal.Text = 0
    txtYTDRemCOLA.Text = 0
    txtYTDSICKLEAVECONVERSION.Text = 0
    txtYTDGross.Text = 0
    txtYTDSSSPAGIBIGPHIC.Text = 0
    txtTaxExemp.Text = 0
    txtTaxableIncome.Text = 0
    txtYTDTaxDue.Text = 0
    txtYTDTaxWithHeld.Text = 0
    txtYTDTaxRefund.Text = 0

    txtFP_VACATIONLEAVE.Text = 0
    txtFP_SICKLEAVE.Text = 0
    txtFP_REMAININGSALARY.Text = 0
    txtFP_REMAININGOVERTIME.Text = 0
    txtFP_REMAININGCOLA.Text = 0
    txtFP_REMAININGCOMMISSION.Text = 0
    txtFP_REMAININGTAXWITHHELD.Text = 0
    txtFP_REMAININGDEDUCTION.Text = 0
    txtFP_TOTALPAY.Text = 0

    txtDED_SSSPREMIUM.Text = 0
    txtDED_PHICPREMIUM.Text = 0
    txtDED_ABSENT.Text = 0
    txtDED_UNDERTIME.Text = 0
    txtDED_OTHERS.Text = 0
    txtDED_TAXPAYABLE.Text = 0

    txtSUMMARY_BASICPAY.Text = 0
    txtSUMMARY_COMMISSION.Text = 0
    txtSUMMARY_OVERTIME.Text = 0
    txtTaxableAdjustment.Text = 0
    txtSUMMARY_ADJUSTEDSALARY.Text = 0
    txtSUMMARY_MIDYEAR.Text = 0
    txtSUMMARY_13THMONTHPAY.Text = 0
    txtSUMMARY_BONUS.Text = 0
    txtSUMMARY_TAXREFUND.Text = 0
End Sub

Sub StoreLoanMemVars()
    On Error GoTo Errorcode

    Dim CNT, crt                                                      As Integer
    Dim LonType                                                       As String

    grdLoanMasDet.ZOrder 1
    Set rsLoanMas = New ADODB.Recordset
    rsLoanMas.Open "select * from HRMS_loanmas where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & IMPNO & "' order by dategranted desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
        'Call clearLoangrd
        rsLoanMas.MoveFirst
        crt = 0
        Do While Not rsLoanMas.EOF
            crt = crt + 1

            'COMMENT BY : MJP 10-08-07 10:57 PM -------------------------------------------------------
            'If Null2String(rsLoanMas!loantype) = "SSAL" Then LonType = "SSS Salary Loan"
            'If Null2String(rsLoanMas!loantype) = "CSAL" Then LonType = "SSS Calamity Loan"
            'If Null2String(rsLoanMas!loantype) = "PSAL" Then LonType = "Pag-Ibig Salary Loan"
            'If Null2String(rsLoanMas!loantype) = "HDMF" Then LonType = "Pag-Ibig HDMF"
            'COMMENT BY : MJP 10-08-07 10:57 PM -------------------------------------------------------

            'UPDATE BY : MJP 10-08-07 10:57 PM -------------------------------------------------------
            'DESCRIPTION : TO GET THE NEW TYPE OF LOAN
            LonType = FindLoanType(rsLoanMas!LOANTYPE)
            'UPDATE BY : MJP 10-08-07 10:57 PM -------------------------------------------------------

            grdLoanMas.AddItem LonType & Chr(9) & Null2String(rsLoanMas!acctno) & Chr(9) & Null2String(rsLoanMas!DateGranted) & _
                               Chr(9) & Null2String(rsLoanMas!MATURITYDATE) & Chr(9) & N2Str2Zero(rsLoanMas!AMOUNTLoaned) & Chr(9) & N2Str2Zero(rsLoanMas!MONTHLYDED) & _
                               Chr(9) & N2Str2Zero(rsLoanMas!SMONTHLYDED) & Chr(9) & N2Str2Zero(rsLoanMas!LoanBalance) & Chr(9) & rsLoanMas!ID & Chr(9) & Null2String(rsLoanMas!EMPNO)
            rsLoanMas.MoveNext
        Loop
        If crt > 0 Then grdLoanMas.RemoveItem 1
    Else
        clearLoangrd
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub InitLoans()
    'cboLoanType.Clear
    'cboLoanType.AddItem "SSS Salary Loan"
    'cboLoanType.AddItem "SSS Calamity Loan"
    'cboLoanType.AddItem "Pag-Ibig Salary Loan"
    'cboLoanType.AddItem "Pag-Ibig HDMF"

    txtAcctNo.Text = ""
    txtDateGranted.Text = Date
    txtMaturityDate.Text = DateSerial(YEAR(LOGDATE) + 2, MONTH(LOGDATE), Day(Date))
    txtAmountLoaned.Text = 0
    txtMonthlyDed.Text = 0
    txtSMonthlyDed.Text = 0
    txtLoanBalance.Text = 0
    cmdLoanMas.ZOrder 0
    fraLoanMas.ZOrder 0
    clearLoangrd
End Sub

Sub StorePHMemvars()
    On Error GoTo Errorcode

    If Not rsPH.EOF And Not rsPH.BOF Then
        txtPHMonthly.Text = Null2String(rsPH!EmployeeShare)
        txtPHStarted.Text = Null2String(rsPH!DateStart)
        txtPHLast.Text = Null2String(rsPH!LastDateCont)
        Set rsPHDet = New ADODB.Recordset
        rsPHDet.Open "select * from HRMS_philhealthdet where aydi = " & rsPH!aydi & " order by deyt desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPHDet.EOF And Not rsPHDet.BOF Then
            rsPHDet.MoveFirst
            Do While Not rsPHDet.EOF
                grdPhilHealth.AddItem rsPHDet!DEYT & Chr(9) & Format(rsPHDet!employeramount, "###,##0.00") & Chr(9) & Format(rsPHDet!EmployeeAMOUNT, "###,##0.00") & Chr(9) & rsPHDet!ID
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

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub StorePagIbigMemvars()
    On Error GoTo Errorcode

    If Not rsPagIbig.EOF And Not rsPagIbig.BOF Then
        txtPagIbigMonthly.Text = Null2String(rsPagIbig!EmployeeShare)
        txtPagIbigStarted.Text = Null2String(rsPagIbig!DateStart)
        txtPagIbigLast.Text = Null2String(rsPagIbig!LastDateCont)
        Set rsPagibigdet = New ADODB.Recordset
        rsPagibigdet.Open "select * from HRMS_pagibigdet where aydi = " & rsPagIbig!aydi & " order by deyt desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPagibigdet.EOF And Not rsPagibigdet.BOF Then
            rsPagibigdet.MoveFirst
            Do While Not rsPagibigdet.EOF
                grdPagIbig.AddItem Null2String(rsPagibigdet!DEYT) & Chr(9) & N2Str2Zero(rsPagibigdet!employeramount) & Chr(9) & N2Str2Zero(rsPagibigdet!EmployeeAMOUNT) & Chr(9) & rsPagibigdet!ID
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

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub StoreTINMemvars()
    On Error GoTo Errorcode

    If Not rsTIN.EOF And Not rsTIN.BOF Then
        clearTINgrd
        txtTINMonthly.Text = Null2String(rsTIN!Deduction)
        txtTINStarted.Text = Null2String(rsTIN!DateStart)
        txtTINLast.Text = Null2String(rsTIN!LastDateCont)
        Set rsTINdet = New ADODB.Recordset
        rsTINdet.Open "select * from HRMS_tindet where aydi = " & rsTIN!aydi & " order by deyt desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTINdet.EOF And Not rsTINdet.BOF Then
            rsTINdet.MoveFirst
            Do While Not rsTINdet.EOF
                grdTIN.AddItem Null2String(rsTINdet!DEYT) & Chr(9) & N2Str2Zero(rsTINdet!AMOUNT) & Chr(9) & rsTINdet!ID
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

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub InitMemvars()
    Dim MM, ddFROM, YY                                                As String
    MM = Trim(STR(MONTH(LOGDATE)))
    YY = Trim(STR(YEAR(LOGDATE)))
    If Day(Date) > 15 Then
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
    txtPagSal.Text = 0#
    txtHDMF.Text = 0#

    txtUndertime.Text = 0#
    txtAbsent.Text = 0#
    txtTelBill.Text = 0#
    txtOthers.Text = 0#
    showGrossWage
End Sub

Sub FillPayroll()
    Set rsPAYROLL = New ADODB.Recordset
    rsPAYROLL.Open "select * from HRMS_Payroll where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & IMPNO & "' order by paydatefrom desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        Screen.MousePointer = 11
        rsPAYROLL.MoveFirst
        Do While Not rsPAYROLL.EOF
            grdPayroll.AddItem Null2Date(rsPAYROLL!paydatefrom) & _
                               Chr(9) & Null2Date(rsPAYROLL!paydateto) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!Rate) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!DailyRate) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!TAXABLEADJ) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!NONTAXABLEADJ) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!OVERTIME) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!HOLIDAY) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!commission) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!GROSS) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!SSSE) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!PHILHEALTHE) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!PAGIBIG) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!TAX) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!SSSSALLOAN) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!SSSCALLOAN) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!PAGSALLOAN) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!PAGHDMFLOAN) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!UNDERTIME) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!ABSENT) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!telbill) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!Others) & _
                               Chr(9) & N2Str2Zero(rsPAYROLL!NETPAY) & _
                               Chr(9) & rsPAYROLL!ID

            rsPAYROLL.MoveNext
        Loop
        grdPayroll.RemoveItem 1
        Screen.MousePointer = 0
    Else
        clearPayrollgrd
    End If
End Sub

Sub OtherRefresh()
    Set rsSSS = New ADODB.Recordset
    rsSSS.Open "select * from HRMS_SSS where sssno = " & N2Str2Null(rsEmpInfo!SSSNO) & " order by datestart desc", gconDMIS, adOpenForwardOnly, adLockReadOnly

    Set rsPH = New ADODB.Recordset
    rsPH.Open "select * from HRMS_philhealth where phno = " & N2Str2Null(rsEmpInfo!SSSNO) & " order by datestart desc", gconDMIS, adOpenForwardOnly, adLockReadOnly

    Set rsPagIbig = New ADODB.Recordset
    rsPagIbig.Open "select * from HRMS_pagibig where pagibigno = " & N2Str2Null(rsEmpInfo!SSSNO) & " order by datestart desc", gconDMIS, adOpenForwardOnly, adLockReadOnly

    Set rsTIN = New ADODB.Recordset
    rsTIN.Open "select * from HRMS_tin where tinno = " & N2Str2Null(rsEmpInfo!tinno) & " order by datestart desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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

Sub InitGrid()
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
        .ColWidth(16) = 900
        .ColWidth(17) = 900
        .ColWidth(18) = 900
        .ColWidth(19) = 900
        .ColWidth(20) = 900
        .ColWidth(21) = 900
        .ColWidth(22) = 900
        .ColWidth(23) = 1
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
        .Text = "Pag Sal. Loan"
        .Col = 17
        .Text = "Pag HDMF"
        .Col = 18
        .Text = "UT/Late"
        .Col = 19
        .Text = "Absent"
        .Col = 20
        .Text = "Tel. Bill"
        .Col = 21
        .Text = "Others"
        .Col = 22
        .Text = "Net Pay"
        .Col = 23
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
    On Error GoTo Errorcode
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
        txtSSSMonthly.Text = Null2String(rsSSS!EmployeeShare)
        txtSSSStarted.Text = Null2String(rsSSS!DateStart)
        txtSSSLast.Text = Null2String(rsSSS!LastDateCont)
        Set rsSSSdet = New ADODB.Recordset
        rsSSSdet.Open "select * from HRMS_SSSdet where aydi = " & rsSSS!aydi & " order by deyt desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsSSSdet.EOF And Not rsSSSdet.BOF Then
            rsSSSdet.MoveFirst
            Do While Not rsSSSdet.EOF
                grdSSS.AddItem Null2Date(rsSSSdet!DEYT) & Chr(9) & N2Str2Zero(rsSSSdet!employeramount) & Chr(9) & N2Str2Zero(rsSSSdet!EmployeeAMOUNT) & Chr(9) & rsSSSdet!ID
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

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub ShowYTDDetails()
    txtYTDSSSPAGIBIGPHIC.Text = ToDoubleNumber(YEAR_TO_DATE_SSS + YEAR_TO_DATE_PHIC + YEAR_TO_DATE_PAGIBIG + NumericVal(txtDED_SSSPREMIUM.Text) + NumericVal(txtDED_PHICPREMIUM.Text))
    txtYTDGross.Text = Round(NumericVal(txtYTDRemSal.Text) + Val(txtTaxableAdjustment.Text) + NumericVal(txtYTDBasicPay.Text) + NumericVal(txtYTDCola.Text) + NumericVal(txtYTDRemCOLA.Text) + NumericVal(txtSUMMARY_COMMISSION.Text) + NumericVal(txtFP_REMAININGCOMMISSION.Text) + NumericVal(txtYTDSICKLEAVECONVERSION.Text), 2)
    'txtYTDGross.Text = Round(NumericVal(txtYTDRemSal.Text) + NumericVal(txtYTDBasicPay.Text) + NumericVal(txtFP_REMAININGCOMMISSION.Text) + NumericVal(txtSUMMARY_COMMISSION.Text) + NumericVal(txtYTDSICKLEAVECONVERSION.Text), 2)
    Dim CHAT_KIM                                                      As Double
    If SetSalary() > 30000 Then
        CHAT_KIM = SetSalary() - 30000
        txtTaxableIncome.Text = Round(((NumericVal(txtYTDGross.Text) + CHAT_KIM) - NumericVal(txtTaxExemp.Text)) - NumericVal(txtYTDSSSPAGIBIGPHIC.Text), 2)
    Else
        txtTaxableIncome.Text = Round((NumericVal(txtYTDGross.Text) - NumericVal(txtTaxExemp.Text)) - NumericVal(txtYTDSSSPAGIBIGPHIC.Text), 2)
    End If
    If NumericVal(txtYTDGross.Text) >= 60000 Then
        If Round(NumericVal(Tax_Due(NumericVal(txtTaxableIncome.Text))), 2) <= 0 Then
            txtYTDTaxDue.Text = "Exempted"
        Else
            txtYTDTaxDue.Text = Round(Tax_Due(NumericVal(txtTaxableIncome.Text)), 2)
        End If
    Else
        txtYTDTaxDue.Text = "Exempted"
    End If
    txtYTDTaxRefund.Text = Round(NumericVal(txtYTDTaxDue.Text) - (NumericVal(txtYTDTaxWithHeld.Text) + NumericVal(txtFP_REMAININGTAXWITHHELD.Text)), 2)
    txtDED_TAXPAYABLE.Text = Round(NumericVal(txtYTDTaxDue.Text) - (NumericVal(txtYTDTaxWithHeld.Text) + NumericVal(txtFP_REMAININGTAXWITHHELD.Text)), 2)
    txtSUMMARY_TAXREFUND.Text = Round(NumericVal(txtYTDTaxDue.Text) - (NumericVal(txtYTDTaxWithHeld.Text) + NumericVal(txtFP_REMAININGTAXWITHHELD.Text)), 2)
    txtSUMMARY_ADJUSTEDSALARY.Text = Round(NumericVal(txtYTDRemSal.Text), 2)
    txtSUMMARY_13THMONTHPAY.Text = Round(((NumericVal(txtYTDBasicPay.Text) + NumericVal(txtYTDRemSal.Text)) / 12) - NumericVal(txtSUMMARY_MIDYEAR), 2)
    txtTOTAL_PAY.Text = ToDoubleNumber((NumericVal(txtFP_VACATIONLEAVE.Text) + NumericVal(txtFP_SICKLEAVE.Text) + NumericVal(txtFP_REMAININGSALARY.Text) + NumericVal(txtFP_REMAININGOVERTIME.Text) + NumericVal(txtFP_REMAININGCOLA.Text) + NumericVal(txtFP_REMAININGCOMMISSION.Text) + NumericVal(txtFP_REMAININGTAXWITHHELD.Text) + NumericVal(Round(NumericVal(txtYTDTaxDue.Text) - (NumericVal(txtYTDTaxWithHeld.Text) + NumericVal(txtFP_REMAININGTAXWITHHELD.Text)), 2))))
    'If NumericVal(txtYTDTaxRefund.Text) < 0 Then
    '   txtSUMMARY_TAXREFUND.Text = Round(Abs(NumericVal(txtYTDTaxRefund.Text)), 2)
    'Else
    '   txtSUMMARY_TAXREFUND.Text = Round(NumericVal(txtYTDTaxRefund.Text), 2)
    'End If
    'txtFP_TOTALPAY.Text = Round((NumericVal(txtSUMMARY_ADJUSTEDSALARY.Text) + NumericVal(txtSUMMARY_13THMONTHPAY.Text) + NumericVal(txtSUMMARY_TAXREFUND.Text) + NumericVal(txtSUMMARY_OVERTIME.Text)) - NumericVal(txtFP_REMAININGDEDUCTION.Text), 2)
    txtFP_TOTALPAY.Text = ToDoubleNumber((NumericVal(txtFP_VACATIONLEAVE.Text) + NumericVal(txtFP_SICKLEAVE.Text) + NumericVal(txtFP_REMAININGSALARY.Text) + NumericVal(txtFP_REMAININGOVERTIME.Text) + NumericVal(txtFP_REMAININGCOLA.Text) + NumericVal(txtFP_REMAININGCOMMISSION.Text) + NumericVal(txtFP_REMAININGTAXWITHHELD.Text) + NumericVal(Round(NumericVal(txtYTDTaxDue.Text) - (NumericVal(txtYTDTaxWithHeld.Text) + NumericVal(txtFP_REMAININGTAXWITHHELD.Text)), 2))) - NumericVal(txtFP_REMAININGDEDUCTION.Text))
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False: lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL order by lastname+', '+firstname asc")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False: lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo  where EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL and lastname+', '+firstname like'" & XXX & "%' order by lastname+', '+firstname asc")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "ACESS_ADD", "EMPLOYEE LEDGER") = False Then Exit Sub

    AddorEdit = "ADD"
    EnAbleFrames
    If TabSSS.Tab = 0 Then
        cmdAddPayroll.Value = True
    End If

    If TabSSS.Tab = 1 Then
        EnAbleTab TabSSS, 1, 3
        InitLoans
    End If

    If TabSSS.Tab = 2 Then
        EnAbleTab TabSSS, 2, 3
        'on error resume next
        txtSSSNo.SetFocus
    End If
    If TabSSS.Tab = 3 Then
        EnAbleTab TabSSS, 3, 3
        'on error resume next
        txtPagIbigNo.SetFocus
    End If
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdAddPayroll_Click()
    AddorEdit = "ADD"
    picPayroll.Visible = True: cmdPayroll.Visible = True
    cmdPayroll.ZOrder 0: picPayroll.ZOrder 0: picAddEditPayroll.ZOrder 0
    txtFrom.Enabled = True: txtTo.Enabled = True
    MAHIRO
    InitMemvars
End Sub

Private Sub cmdCancel_Click()
    cmdLoanMas.ZOrder 1: fraLoanMas.ZOrder 1: cmdPayroll.ZOrder 1
    picPayroll.Visible = False: cmdPayroll.Visible = False
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
    TabSSS.Enabled = True
    picPayroll2.Visible = False
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "ACESS_DELETE", "EMPLOYEE LEDGER") = False Then Exit Sub
    cmdDeletePayroll.Value = True
End Sub

Private Sub cmdDeletePayroll_Click()
    grdPayroll.Col = 23

    If grdPayroll.Text <> "" Then
        'MsgSpeechBox "Delete selected Record? Are you Sure?"
        If MsgBoxXP("Are you Sure?", "Delete selected Record", XP_YesNo, msg_Question) = True Then
            Set rsPAYROLL = New ADODB.Recordset
            rsPAYROLL.Open "select * from HRMS_Payroll where id = " & grdPayroll.Text, gconDMIS
            If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
                GENTO = rsPAYROLL!paydateto
                GENFROM = rsPAYROLL!paydatefrom
                IMPNO = IMPNO

                LogAudit "X", "DELETE EMPLOYEE LEDGER PAYROLL", EMPLOYEE_NO & "-" & grdPayroll.Text

                DelEXIST
                ShowDeletedMsg
            End If
        End If
    Else
        MsgSpeechBox "Nothing to Delete"
    End If

    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "ACESS_EDIT", "EMPLOYEE LEDGER") = False Then Exit Sub
    AddorEdit = "EDIT"
    EnAbleFrames
    Picture1.Visible = False
    Picture2.Visible = True
    If TabSSS.Tab = 0 Then
        cmdEditPayroll.Value = True
    End If
    If TabSSS.Tab = 4 Then
        cmdEditYTD.Value = True
    End If
End Sub

Private Sub cmdEditPayroll_Click()
    Dim fild                                                          As String
    TabSSS.TabEnabled(0) = False
    TabSSS.TabEnabled(1) = False
    TabSSS.TabEnabled(2) = False
    TabSSS.TabEnabled(3) = False
    TabSSS.TabEnabled(4) = False
    grdPayroll.Row = grdPayroll.Row
    grdPayroll.Col = 23
    fild = grdPayroll.Text
    If fild <> "" Then
        MAIDIT (fild)
    End If
End Sub

Private Sub cmdEditYTD_Click()
    TabSSS.Tab = 4
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    picSearch.ZOrder 0
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsEmpInfo.MoveNext
    If rsEmpInfo.EOF Then
        rsEmpInfo.MoveLast
        MsgSpeechBox "Last of Record!"
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsEmpInfo.MovePrevious
    If rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        MsgSpeechBox "First of Record!"
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "ACESS_PRINT", "EMPLOYEE LEDGER") = False Then Exit Sub
    Dim FILTER                                                        As String
    If TabSSS.Tab = 0 Then
        cmdPrintPayroll.Value = True
    End If
    If TabSSS.Tab = 4 Then
        Screen.MousePointer = 11
        PrintSQLReport rptPayroll, HRMS_REPORT_PATH & "finalpay.rpt", "{ytddetails.empno} = " & N2Str2Null(IMPNO) & " and {ytddetails.yeer} = '" & cboYear.Text & "'", HRMS_REPORT_Connection, 1

        LogAudit "V", "PRINT EMPLOYEE LEDGER", EMPLOYEE_NO
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdPrintPayroll_Click()
    Dim FILTER                                                        As String
    Screen.MousePointer = 11
    Dim CLID                                                          As String
    grdPayroll.Row = grdPayroll.Row
    grdPayroll.Col = 13
    CLID = grdPayroll.Text

    If CLID <> "" Then
        LogAudit "V", "PRINT EMPLOYEE LEDGER", IMPNO
        PrintSQLReport rptPayroll, HRMS_REPORT_PATH & "ledger.rpt", "{payroll.empno} = " & N2Str2Null(IMPNO), HRMS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    Dim LonID                                                         As Integer

    If TabSSS.Tab = 0 Then
        cmdSavePayroll.Value = True
    End If

    If TabSSS.Tab = 1 Then
        Dim LAWNType, VtxtAcctNo, VtxtDateGranted, VtxtMaturityDate   As String
        Dim VtxtAmountLoaned, VtxtMonthlyDed, VtxtSMonthlyDed, VtxtLoanBalance As Double

        If txtAcctNo.Text = "" Then
            MsgBox "Account No. Cannot be Blank", vbInformation, "Loan Required"
            txtAcctNo.SetFocus
            Exit Sub
        End If

        'COMMENT BY : MJP 10-24-07 05:40 PM -----------------------------------------
        'If cboLoanType.Text = "SSS Salary Loan" Then
        ' LAWNType = "'SSAL'"
        'ElseIf cboLoanType.Text = "SSS Calamity Loan" Then
        ' LAWNType = "'CSAL'"
        'ElseIf cboLoanType.Text = "Pag-Ibig Salary Loan" Then
        ' LAWNType = "'PSAL'"
        'ElseIf cboLoanType.Text = "Pag-Ibig HDMF" Then
        ' LAWNType = "'HDMF'"
        'Else
        ' LAWNType = "'OTHER'"
        'End If
        'COMMENT BY : MJP 10-24-07 05:40 PM -----------------------------------------

        'UPDATE BY : MJP 10-24-07 05:40 PM ---------------------------------------------
        'DESCRIPTION : TO GET THE CODE ONLY FROM THE SELECTED ITEM
        LAWNType = Right(cboLoanType, 4)
        'UPDATE BY : MJP 10-24-07 05:40 PM ---------------------------------------------

        VtxtAcctNo = N2Str2Null(txtAcctNo.Text)
        VtxtDateGranted = N2Date2Null(txtDateGranted.Text)
        VtxtMaturityDate = N2Date2Null(txtMaturityDate.Text)
        VtxtAmountLoaned = N2Str2Zero(txtAmountLoaned.Text)
        VtxtMonthlyDed = N2Str2Zero(txtMonthlyDed.Text)
        VtxtSMonthlyDed = N2Str2Zero(txtSMonthlyDed.Text)
        VtxtLoanBalance = N2Str2Zero(txtLoanBalance.Text)
        If AddorEdit = "ADD" Then
            gconDMIS.Execute "insert into HRMS_loanmas " & _
                             "(EMPLEVEL,loantype,empno,acctno,dategranted,maturitydate,amountloaned,monthlyded,smonthlyded,loanbalance)" & _
                           " values (" & EMPLIVIL & ",'" & LAWNType & "', '" & IMPNO & "', " & VtxtAcctNo & ", " & VtxtDateGranted & _
                             ", " & VtxtMaturityDate & ", " & VtxtAmountLoaned & ", " & VtxtMonthlyDed & ", " & VtxtSMonthlyDed & ", " & VtxtLoanBalance & ")"

            LogAudit "A", "ADD EMPLOYEE LOAN", IMPNO & "-" & LAWNType
        Else
            grdLoanMas.Col = 8
            LonID = grdLoanMas.Text
            gconDMIS.Execute "update hrms_loanmas set" & _
                           " EMPLEVEL = " & EMPLIVIL & "," & _
                           " loantype = '" & LAWNType & "'," & _
                           " acctno = " & VtxtAcctNo & "," & _
                           " dategranted = " & VtxtDateGranted & "," & _
                           " maturitydate = " & VtxtMaturityDate & "," & _
                           " amountloaned = " & VtxtAmountLoaned & "," & _
                           " monthlyded = " & VtxtMonthlyDed & "," & _
                           " smonthlyded = " & VtxtSMonthlyDed & "," & _
                           " loanbalance = " & VtxtLoanBalance & _
                           " where id = " & LonID

            LogAudit "E", "UPDATE EMPLOYEE LOAN", IMPNO & "-" & LAWNType
        End If
        grdStore
    End If

    If TabSSS.Tab = 2 Then
        If txtSSSNo.Text <> "" Then
            gconDMIS.Execute "update sss set" & _
                           " employeeshare = " & N2Str2Zero(txtSSSMonthly.Text) & "," & _
                           " datestart = '" & txtSSSStarted.Text & "'," & _
                           " lastdatecont = '" & txtSSSLast.Text & "'" & _
                           " where sssno = " & N2Str2Null(rsEmpInfo!SSSNO)
        End If
        If txtPhilHealthNo.Text <> "" Then
            gconDMIS.Execute "update philhealth set" & _
                           " employeeshare = " & N2Str2Zero(txtPHMonthly.Text) & "," & _
                           " datestart = '" & txtPHStarted.Text & "'," & _
                           " lastdatecont = '" & txtPHLast.Text & "'" & _
                           " where phno = " & N2Str2Null(rsEmpInfo!phno)
        End If
        grdStore
    End If

    If TabSSS.Tab = 3 Then
        If txtPagIbigNo.Text <> "" Then
            gconDMIS.Execute "update pagibig set" & _
                           " employeeshare = " & N2Str2Zero(txtPagIbigMonthly.Text) & "," & _
                           " datestart = '" & txtPagIbigStarted.Text & "'," & _
                           " lastdatecont = '" & txtPagIbigLast.Text & "'" & _
                           " where pagibigno = " & N2Str2Null(rsEmpInfo!pagibigno)
        End If
        If txtTINNo.Text <> "" Then
            gconDMIS.Execute "update tin set" & _
                           " deduction = " & N2Str2Zero(txtTINMonthly.Text) & "," & _
                           " datestart = '" & txtTINStarted.Text & "'," & _
                           " lastdatecont = '" & txtTINLast.Text & "'" & _
                           " where tinno = " & N2Str2Null(rsEmpInfo!tinno)
        End If
        grdStore
    End If

    If TabSSS.Tab = 4 Then
        gconDMIS.Execute "update HRMS_ytddetails set" & _
                       " midyear = " & NumericVal(txtSUMMARY_MIDYEAR.Text) & ", remcola = " & NumericVal(txtYTDRemCOLA.Text) & "," & _
                       " remsal = " & NumericVal(txtYTDRemSal.Text) & ", remot = " & NumericVal(txtSUMMARY_OVERTIME.Text) & "," & _
                       " remwtax = " & NumericVal(txtFP_REMAININGTAXWITHHELD.Text) & ", remDed = " & NumericVal(txtFP_REMAININGDEDUCTION.Text) & "," & _
                       " ytdincome = " & NumericVal(txtYTDGross.Text) & ", personalex = " & NumericVal(txtTaxExemp.Text) & "," & _
                       " t13thmonth = " & NumericVal(txtSUMMARY_13THMONTHPAY.Text) & ", bonus = " & NumericVal(txtSUMMARY_BONUS.Text) & "," & _
                       " taxdue = " & NumericVal(txtYTDTaxDue.Text) & ", FP_VACATIONLEAVE = " & NumericVal(txtFP_VACATIONLEAVE.Text) & "," & _
                       " FP_SICKLEAVE = " & NumericVal(txtFP_SICKLEAVE.Text) & ", FP_REMAININGSALARY = " & NumericVal(txtFP_REMAININGSALARY.Text) & "," & _
                       " FP_REMAININGOVERTIME = " & NumericVal(txtFP_REMAININGOVERTIME.Text) & "," & _
                       " FP_REMAININGCOLA = " & NumericVal(txtFP_REMAININGCOLA.Text) & "," & _
                       " FP_REMAININGCOMMISSION = " & NumericVal(txtFP_REMAININGCOMMISSION.Text) & "," & _
                       " FP_REMAININGTAXWITHHELD = " & NumericVal(txtFP_REMAININGTAXWITHHELD.Text) & "," & _
                       " FP_REMAININGDEDUCTION = " & NumericVal(txtFP_REMAININGDEDUCTION.Text) & "," & _
                       " FP_TOTALPAY = " & NumericVal(txtFP_TOTALPAY.Text) & "," & _
                       " DED_SSSPREMIUM = " & NumericVal(txtDED_SSSPREMIUM.Text) & "," & _
                       " DED_PHICPREMIUM = " & NumericVal(txtDED_PHICPREMIUM.Text) & "," & _
                       " DED_HDMFPREMIUM = " & NumericVal(txtDED_HDMFPREMIUM.Text) & "," & _
                       " DED_ABSENT = " & NumericVal(txtDED_ABSENT.Text) & "," & _
                       " DED_UNDERTIME = " & NumericVal(txtDED_UNDERTIME.Text) & "," & _
                       " DED_OTHERS = " & NumericVal(txtDED_OTHERS.Text) & "," & _
                       " DED_TAXPAYABLE = " & Val(txtDED_TAXPAYABLE.Text) & _
                       " where empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND yeer = '" & cboYear.Text & "'"
    End If
    cmdCancel.Value = True

    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdSavePayroll_Click()
    Dim net, gruss                                                    As Double
    GENFROM = Format(txtFrom.Text, "Short Date")
    GENTO = Format(txtTo.Text, "Short Date")
    If IsDate(GENFROM) = False Then
        MsgBoxXP "Error in From Date", "Error!", XP_OKOnly, msg_Critical
        Exit Sub
    End If
    If IsDate(GENTO) = False Then
        MsgBoxXP "Error in To Date", "Error!", XP_OKOnly, msg_Critical
        Exit Sub
    End If
    IMPNO = IMPNO
    DelEXIST
    SEYV
End Sub

Private Sub cmdSelectYear_Click()
    StoreYTD
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            grdLoanMasDet.ZOrder 1
        Case vbKeyF5
            If TabSSS.Tab = 4 Then
                frmForm2316.Show
                frmForm2316.cboYear.Text = cboYear.Text
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    txtSearch.Text = ""
    InitGrid
    rsrefresh
    picPayroll.Visible = False: cmdPayroll.Visible = False

    'COMMENT BY : MJP 10-08-07 --------------------------------------------------------
    'cboLoanType.Clear
    'cboLoanType.AddItem "SSS Salary Loan"
    'cboLoanType.AddItem "SSS Calamity Loan"
    'cboLoanType.AddItem "Pag-Ibig Salary Loan"
    'cboLoanType.AddItem "Pag-Ibig HDMF"
    'COMMENT BY : MJP 10-08-07 --------------------------------------------------------

    'UPDATE BY : MJP 10-24-07 05:38 PM -----------------------------------------------
    'DESCRIPTION : TO FILL THE LIST OF LOAN FROM THE MASTER FILE
    FillCboLoanType
    'UPDATE BY : MJP 10-24-07 05:38 PM -----------------------------------------------

    FillcboYear cboYear
    cboYear.Text = YEAR(LOGDATE)
    DisAbleFrames
    StoreMemVars
    DrawXPCtl Me
    LEDGERSHOW = True
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LEDGERSHOW = False
    Set frmHRMSLedger = Nothing
End Sub

Private Sub grdLoanMas_Click()
    Dim CNT                                                           As Integer
    CNT = 0
    grdLoanMas.Col = 9
    If grdLoanMas.Text <> "" Then
        Set rsLoanmasDet = New ADODB.Recordset
        rsLoanmasDet.Open "select * from HRMS_loanmasdet where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(grdLoanMas.Text) & " order by deyt desc", gconDMIS
        If Not rsLoanmasDet.EOF And Not rsLoanmasDet.BOF Then
            clearLoanDetgrd
            grdLoanMasDet.ZOrder 0
            rsLoanmasDet.MoveFirst
            Do While Not rsLoanmasDet.EOF
                CNT = CNT + 1
                grdLoanMasDet.AddItem Null2String(rsLoanmasDet!acctno) & Chr(9) & Null2String(rsLoanmasDet!DEYT) & Chr(9) & N2Str2Zero(rsLoanmasDet!AMOUNT) & Chr(9) & rsLoanmasDet!ID
                rsLoanmasDet.MoveNext
            Loop
        End If
        If CNT > 0 Then grdLoanMasDet.RemoveItem 1
    End If
End Sub

Private Sub grdLoanMas_DblClick()
    grdLoanMas.Row = grdLoanMas.Row
    grdLoanMas.Col = 8
    CLID = grdLoanMas.Text
    If CLID <> "" Then
        Set rsLoanMas = New ADODB.Recordset
        rsLoanMas.Open "select * from HRMS_loanmas where id =" & CLID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
            cmdLoanMas.ZOrder 0
            fraLoanMas.ZOrder 0
            fraLoanMas.Enabled = True
            If Null2String(rsLoanMas!LOANTYPE) = "SSAL" Then
                cboLoanType.Text = "SSS Salary Loan"
            End If
            If Null2String(rsLoanMas!LOANTYPE) = "CSAL" Then
                cboLoanType.Text = "SSS Calamity Loan"
            End If
            If Null2String(rsLoanMas!LOANTYPE) = "PSAL" Then
                cboLoanType.Text = "Pag-Ibig Salary Loan"
            End If
            If Null2String(rsLoanMas!LOANTYPE) = "HDMF" Then
                cboLoanType.Text = "Pag-Ibig HDMF"
            End If
            txtAcctNo.Text = Null2String(rsLoanMas!acctno)
            txtDateGranted.Text = Null2Date(rsLoanMas!DateGranted)
            txtMaturityDate.Text = Null2Date(rsLoanMas!MATURITYDATE)
            txtAmountLoaned.Text = N2Str2Zero(rsLoanMas!AMOUNTLoaned)
            txtMonthlyDed.Text = N2Str2Zero(rsLoanMas!MONTHLYDED)
            txtSMonthlyDed.Text = N2Str2Zero(rsLoanMas!SMONTHLYDED)
            txtLoanBalance.Text = N2Str2Zero(rsLoanMas!LoanBalance)
            Picture1.Visible = False
            Picture2.Visible = True
        End If
    End If
End Sub

Private Sub grdPagIbig_DblClick()
    Dim piID                                                          As String
    grdPagIbig.Row = grdPagIbig.Row
    grdPagIbig.Col = 3
    piID = grdPagIbig.Text

    If piID <> "" Then
        Set rsPagibigdet = New ADODB.Recordset
        rsPagibigdet.Open "select * from HRMS_pagibigdet where id =" & piID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPagibigdet.EOF And Not rsPagibigdet.BOF Then
            StoreDed Null2Date(rsPagibigdet!DEYT)
        End If
    End If
End Sub

Private Sub grdPayroll_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub grdPhilHealth_DblClick()
    Dim phID                                                          As String
    grdPhilHealth.Row = grdPhilHealth.Row
    grdPhilHealth.Col = 3
    phID = grdPhilHealth.Text

    If phID <> "" Then
        Set rsPHDet = New ADODB.Recordset
        rsPHDet.Open "select * from HRMS_philhealthdet where id =" & phID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPHDet.EOF And Not rsPHDet.BOF Then
            StoreDed (Null2Date(rsPHDet!DEYT))
        End If
    End If
End Sub

Private Sub grdLoanMasDet_DblClick()
    Dim slID                                                          As String
    grdLoanMasDet.Row = grdLoanMasDet.Row
    grdLoanMasDet.Col = 3
    slID = grdLoanMasDet.Text
    If slID <> "" Then
        Set rsLoanmasDet = New ADODB.Recordset
        rsLoanmasDet.Open "select * from HRMS_loanmasdet where id =" & slID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsLoanmasDet.EOF And Not rsLoanmasDet.BOF Then
            StoreDed (Null2Date(rsLoanmasDet!DEYT))
        End If
    End If
End Sub

Private Sub grdSSS_DblClick()
    Dim ssID                                                          As String
    grdSSS.Row = grdSSS.Row
    grdSSS.Col = 3
    ssID = grdSSS.Text
    If ssID <> "" Then
        Set rsSSSdet = New ADODB.Recordset
        rsSSSdet.Open "select * from HRMS_SSSdet where id =" & ssID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsSSSdet.EOF And Not rsSSSdet.BOF Then
            StoreDed (Null2Date(rsSSSdet!DEYT))
        End If
    End If
End Sub

Private Sub grdTIN_DblClick()
    Dim tiID                                                          As String
    grdTIN.Row = grdTIN.Row
    grdTIN.Col = 2
    tiID = grdTIN.Text
    If tiID <> "" Then
        Set rsTINdet = New ADODB.Recordset
        rsTINdet.Open "select * from HRMS_tindet where id =" & tiID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTINdet.EOF And Not rsTINdet.BOF Then
            StoreDed (Null2Date(rsTINdet!DEYT))
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
    txtMonthlyDed.Text = ((NumericVal(txtAmountLoaned.Text) * 0.12) + NumericVal(txtAmountLoaned.Text)) / 24
    txtSMonthlyDed.Text = NumericVal(txtMonthlyDed.Text) / 2
    txtLoanBalance.Text = (NumericVal(txtAmountLoaned.Text) * 0.12) + NumericVal(txtAmountLoaned.Text)
End Sub

Private Sub txtCalLoan_Change()
    ShowTotDed
End Sub

Private Sub txtCalLoan_LostFocus()
    ShowTotDed
End Sub

Private Sub txtCommission_Change()
    showGrossWage
End Sub

Private Sub txtCommission_LostFocus()
    showGrossWage
End Sub

Private Sub txtDED_ABSENT_LostFocus()
    txtFP_REMAININGDEDUCTION.Text = NumericVal(txtDED_SSSPREMIUM.Text) + NumericVal(txtDED_PHICPREMIUM.Text) + NumericVal(txtDED_ABSENT.Text) + NumericVal(txtDED_UNDERTIME.Text) + NumericVal(txtDED_OTHERS.Text) + Val(txtDED_TAXPAYABLE.Text)
    ShowYTDDetails
End Sub

Private Sub txtDED_OTHERS_LostFocus()
    txtFP_REMAININGDEDUCTION.Text = NumericVal(txtDED_SSSPREMIUM.Text) + NumericVal(txtDED_PHICPREMIUM.Text) + NumericVal(txtDED_ABSENT.Text) + NumericVal(txtDED_UNDERTIME.Text) + NumericVal(txtDED_OTHERS.Text) + Val(txtDED_TAXPAYABLE.Text)
    ShowYTDDetails
End Sub

Private Sub txtDED_PHICPREMIUM_LostFocus()
    txtFP_REMAININGDEDUCTION.Text = NumericVal(txtDED_SSSPREMIUM.Text) + NumericVal(txtDED_PHICPREMIUM.Text) + NumericVal(txtDED_ABSENT.Text) + NumericVal(txtDED_UNDERTIME.Text) + NumericVal(txtDED_OTHERS.Text) + Val(txtDED_TAXPAYABLE.Text)
    ShowYTDDetails
End Sub

Private Sub txtDED_SSSPREMIUM_LostFocus()
    txtFP_REMAININGDEDUCTION.Text = NumericVal(txtDED_SSSPREMIUM.Text) + NumericVal(txtDED_PHICPREMIUM.Text) + NumericVal(txtDED_ABSENT.Text) + NumericVal(txtDED_UNDERTIME.Text) + NumericVal(txtDED_OTHERS.Text) + Val(txtDED_TAXPAYABLE.Text)
    ShowYTDDetails
End Sub

Private Sub txtDED_UNDERTIME_LostFocus()
    txtFP_REMAININGDEDUCTION.Text = NumericVal(txtDED_SSSPREMIUM.Text) + NumericVal(txtDED_PHICPREMIUM.Text) + NumericVal(txtDED_ABSENT.Text) + NumericVal(txtDED_UNDERTIME.Text) + NumericVal(txtDED_OTHERS.Text) + Val(txtDED_TAXPAYABLE.Text)
    ShowYTDDetails
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
            txtPHMonthly.Text = PhilHealthShare(rsEmpInfo!SALARY)
        End If
    End If
    With WIZVAR
        If .EncryptAccess(LOGNAME) = "77697A∂ùº∞û¡¡¢¡" Then
            If NumericVal(txtPHMonthly.Text) = 0# Then
                txtPHMonthly.Text = PhilHealthShare(rsEmpInfo!SALARY)
            End If
        End If
    End With
End Sub

Private Sub txtRate_Change()
    showGrossWage
    If AddorEdit = "ADD" Then
        txtDailyRate.Text = ((NumericVal(txtRate.Text) * 2) * 12) / 314
    End If
    With WIZVAR
        If .EncryptAccess(LOGNAME) = "77697A∂ùº∞û¡¡¢¡" Then
            txtDailyRate.Text = ((NumericVal(txtRate.Text) * 2) * 12) / 314
        End If
    End With
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
        txtSSSMonthly.Text = EmployeeSSSshare(rsEmpInfo!SALARY)
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

Private Sub txtFP_REMAININGDEDUCTION_LostFocus()
    txtTOTAL_REMDED.Text = ToDoubleNumber(NumericVal(txtFP_REMAININGDEDUCTION.Text))
    ShowYTDDetails
End Sub

Private Sub txtSUMMARY_OVERTIME_LostFocus()
    ShowYTDDetails
End Sub

Private Sub txtYTDRemCOLA_LostFocus()
    txtFP_REMAININGCOLA.Text = ToDoubleNumber(txtYTDRemCOLA.Text)
    ShowYTDDetails
End Sub

Private Sub txtYTDRemSal_LostFocus()
    txtFP_REMAININGSALARY.Text = ToDoubleNumber(txtYTDRemSal.Text)
    ShowYTDDetails
End Sub

Private Sub txtFP_REMAININGTAXWITHHELD_LostFocus()
    ShowYTDDetails
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsEmpInfo.Bookmark = rsFind(rsEmpInfo.Clone, "empno", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lsAdjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsAdjustment
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub lsAdjustment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSearch.Text) = "" Then FillGrid Else FillSearchGrid (txtSearch.Text)
End Sub

Private Sub txtYTDTaxRefund_LostFocus()
    txtDED_TAXPAYABLE.Text = ToDoubleNumber(txtYTDTaxRefund.Text)
End Sub


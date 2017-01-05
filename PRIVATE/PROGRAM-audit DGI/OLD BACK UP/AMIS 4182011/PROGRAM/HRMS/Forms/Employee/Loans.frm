VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMSLoans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loans Entry"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Loans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   13035
   Begin VB.PictureBox picLoan_Details 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   5520
      ScaleHeight     =   5115
      ScaleWidth      =   4425
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      Begin MSComCtl2.DTPicker txtTran_Date 
         Height          =   375
         Left            =   150
         TabIndex        =   33
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   39583
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pay Period Information"
         Height          =   1005
         Left            =   90
         TabIndex        =   37
         Top             =   2970
         Width           =   4155
         Begin VB.TextBox txtLoan_Year 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2670
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   450
            Width           =   1395
         End
         Begin VB.TextBox txtLoan_Month 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   450
            Width           =   1155
         End
         Begin VB.TextBox txtLoan_Cut_Off 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   450
            Width           =   1365
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Year"
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
            Height          =   210
            Index           =   3
            Left            =   2700
            TabIndex        =   40
            Top             =   210
            Width           =   915
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Month"
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
            Height          =   210
            Index           =   2
            Left            =   1470
            TabIndex        =   39
            Top             =   210
            Width           =   1035
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Cut-off Code"
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
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   38
            Top             =   210
            Width           =   1725
         End
      End
      Begin VB.TextBox txtLoanBal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1890
         Width           =   4185
      End
      Begin VB.OptionButton Option2 
         Caption         =   "(-)"
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
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2520
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OptionButton Option1 
         Caption         =   "(+)"
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2520
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtLoanDescription 
         Height          =   375
         Left            =   90
         TabIndex        =   28
         Top             =   1230
         Width           =   4155
      End
      Begin VB.TextBox txtTran_DetTranno 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "000000"
         Top             =   630
         Width           =   1005
      End
      Begin VB.TextBox txtTranLoanCode 
         Alignment       =   1  'Right Justify
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   630
         Width           =   1695
      End
      Begin VB.TextBox txtTran_AccountNo 
         Alignment       =   1  'Right Justify
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
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0000000000"
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox txtTran_Amount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2460
         TabIndex        =   36
         Top             =   2520
         Width           =   1785
      End
      Begin VB.CommandButton cmdTranClose 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   3540
         MouseIcon       =   "Loans.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   705
      End
      Begin VB.CommandButton cmdTranSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   2850
         MouseIcon       =   "Loans.frx":079A
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Save Entry"
         Top             =   4080
         Width           =   705
      End
      Begin VB.CommandButton cmdTranDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   2160
         MouseIcon       =   "Loans.frx":0C3C
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":0D8E
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Delete Selected Record"
         Top             =   4080
         Width           =   705
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balance"
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
         Height          =   210
         Left            =   90
         TabIndex        =   29
         Top             =   1650
         Width           =   1335
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   4425
         _Version        =   655364
         _ExtentX        =   7805
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Loan Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref #"
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
         Height          =   210
         Left            =   1530
         TabIndex        =   22
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   210
         Left            =   90
         TabIndex        =   27
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Code"
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
         Height          =   210
         Left            =   2565
         TabIndex        =   23
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   " Date"
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
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   2310
         Width           =   525
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Height          =   210
         Left            =   2430
         TabIndex        =   32
         Top             =   2310
         Width           =   750
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
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
         Height          =   210
         Left            =   90
         TabIndex        =   21
         Top             =   390
         Width           =   1170
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000D&
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   -210
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin VB.PictureBox picLoanOption 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   6120
      ScaleHeight     =   2955
      ScaleWidth      =   3285
      TabIndex        =   96
      Top             =   1920
      Visible         =   0   'False
      Width           =   3315
      Begin VB.PictureBox Picture6 
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
         Left            =   1680
         ScaleHeight     =   885
         ScaleWidth      =   1440
         TabIndex        =   101
         Top             =   1980
         Width           =   1440
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   720
            MouseIcon       =   "Loans.frx":10B9
            MousePointer    =   99  'Custom
            Picture         =   "Loans.frx":120B
            Style           =   1  'Graphical
            TabIndex        =   102
            ToolTipText     =   "Cancel"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Save"
            Height          =   795
            Left            =   30
            MouseIcon       =   "Loans.frx":1549
            MousePointer    =   99  'Custom
            Picture         =   "Loans.frx":169B
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Save Entry"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.TextBox txtLoanRemarks 
         Height          =   1155
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   100
         Top             =   810
         Width           =   2955
      End
      Begin VB.OptionButton optLoanOption_Disable 
         Caption         =   "Disable Loan"
         Height          =   375
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   420
         Width           =   1485
      End
      Begin VB.OptionButton optLoanOption_Enable 
         Caption         =   "Enable Loan"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label labselitem 
         Caption         =   "Label1"
         Height          =   225
         Left            =   510
         TabIndex        =   104
         Top             =   2460
         Visible         =   0   'False
         Width           =   405
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   345
         Left            =   0
         TabIndex        =   97
         Top             =   0
         Width           =   5235
         _Version        =   655364
         _ExtentX        =   9234
         _ExtentY        =   609
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox LOANBAL 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2550
      TabIndex        =   94
      Top             =   7170
      Width           =   1485
   End
   Begin VB.TextBox txtSTORE_BEG_BAL 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   93
      Top             =   7170
      Width           =   1455
   End
   Begin VB.PictureBox picSearch 
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
      Height          =   7785
      Left            =   30
      ScaleHeight     =   7785
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   0
      Width           =   2475
      Begin VB.OptionButton optSearch_LoanCode 
         Caption         =   "Loan Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   92
         Top             =   540
         Width           =   1815
      End
      Begin VB.OptionButton optSearch_Empno 
         Caption         =   "Employee No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   91
         Top             =   285
         Width           =   1815
      End
      Begin VB.OptionButton optSearch_EmployeeName 
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   90
         Top             =   30
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   810
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstEmployees 
         Height          =   6525
         Left            =   -30
         TabIndex        =   2
         Top             =   1200
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   11509
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
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Loans.frx":19EB
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
         Picture         =   "Loans.frx":1B4D
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
      Height          =   855
      Left            =   7380
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   78
      Top             =   6915
      Width           =   5580
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   4860
         MouseIcon       =   "Loans.frx":158BA
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":15A0C
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   4170
         MouseIcon       =   "Loans.frx":15D72
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":15EC4
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   3480
         MouseIcon       =   "Loans.frx":1622A
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":1637C
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   2790
         MouseIcon       =   "Loans.frx":166A7
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":167F9
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2100
         MouseIcon       =   "Loans.frx":16B55
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":16CA7
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   1410
         MouseIcon       =   "Loans.frx":16FBA
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":1710C
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   720
         MouseIcon       =   "Loans.frx":17406
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":17558
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   30
         MouseIcon       =   "Loans.frx":178B0
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":17A02
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture3 
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
      Height          =   675
      Left            =   2520
      ScaleHeight     =   675
      ScaleWidth      =   11175
      TabIndex        =   3
      Top             =   0
      Width           =   11175
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   6
         Top             =   270
         Width           =   5445
      End
      Begin VB.TextBox txtEmployeeNumber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1965
      End
      Begin VB.TextBox txtEmpLevel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7500
         TabIndex        =   8
         Top             =   270
         Width           =   465
      End
      Begin Crystal.CrystalReport rptDeductions 
         Left            =   2190
         Top             =   -120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Height          =   315
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
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
         Left            =   5505
         TabIndex        =   5
         Top             =   30
         Width           =   1710
      End
   End
   Begin VB.PictureBox picDeductions 
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
      Height          =   4305
      Left            =   4485
      ScaleHeight     =   4275
      ScaleWidth      =   6495
      TabIndex        =   49
      Top             =   1260
      Visible         =   0   'False
      Width           =   6525
      Begin VB.TextBox txtBEG_Bal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1890
         TabIndex        =   70
         Top             =   3630
         Width           =   1965
      End
      Begin VB.ComboBox cboDeductionOption 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1260
         Width           =   4335
      End
      Begin VB.TextBox txtLoan_SMonthlyDed 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1890
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   3240
         Width           =   1965
      End
      Begin VB.TextBox txtLoan_AmountLoaned 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1890
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   2850
         Width           =   1965
      End
      Begin VB.TextBox txtLoan_Acctno 
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
         Left            =   1890
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   1650
         Width           =   1935
      End
      Begin VB.TextBox txtTranno 
         Alignment       =   1  'Right Justify
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
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   510
         Width           =   1965
      End
      Begin VB.ComboBox cboLOAN_LOANTYPE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1890
         TabIndex        =   56
         Text            =   "cboLOAN_LOANTYPE"
         Top             =   885
         Width           =   4365
      End
      Begin MSComCtl2.DTPicker txtLoan_MaturityDate 
         Height          =   345
         Left            =   1890
         TabIndex        =   64
         Top             =   2430
         Width           =   1965
         _ExtentX        =   3466
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
         CalendarTitleForeColor=   -2147483635
         CalendarTrailingForeColor=   16777215
         Format          =   20578307
         CurrentDate     =   39525
      End
      Begin MSComCtl2.DTPicker txtLoan_DateGranted 
         Height          =   345
         Left            =   1890
         TabIndex        =   62
         Top             =   2040
         Width           =   1965
         _ExtentX        =   3466
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
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20578307
         CurrentDate     =   39525
      End
      Begin MSComCtl2.DTPicker txtLoan_DateStarted 
         Height          =   345
         Left            =   1905
         TabIndex        =   52
         Top             =   480
         Width           =   1665
         _ExtentX        =   2937
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
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20578307
         CurrentDate     =   39525
      End
      Begin VB.Frame Frame1 
         Height          =   1485
         Left            =   3990
         TabIndex        =   71
         Top             =   2550
         Width           =   2205
         Begin VB.TextBox txtLoan_LoanBalance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   990
            Width           =   2025
         End
         Begin VB.TextBox txtNoPay 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   90
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   420
            Width           =   2025
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Computed Loan Balance"
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
            Left            =   -825
            TabIndex        =   74
            Top             =   780
            Width           =   2970
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Pay left"
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
            Left            =   90
            TabIndex        =   72
            Top             =   180
            Width           =   1905
         End
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Started"
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
         Left            =   780
         TabIndex        =   51
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Balance"
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
         Left            =   270
         TabIndex        =   69
         Top             =   3750
         Width           =   1575
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction Option"
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
         Left            =   390
         TabIndex        =   57
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Installment Amount"
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
         Left            =   150
         TabIndex        =   67
         Top             =   3300
         Width           =   1650
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Principal Amount"
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
         Left            =   390
         TabIndex        =   65
         Top             =   2910
         Width           =   1455
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref #"
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
         Left            =   3690
         TabIndex        =   53
         Top             =   540
         Width           =   435
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   435
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   6495
         _Version        =   655364
         _ExtentX        =   11456
         _ExtentY        =   767
         _StockProps     =   14
         Caption         =   "ADD/EDIT LOANS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Type"
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
         Left            =   975
         TabIndex        =   55
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
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
         Left            =   825
         TabIndex        =   59
         Top             =   1695
         Width           =   1020
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Maturity Date"
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
         Left            =   720
         TabIndex        =   63
         Top             =   2460
         Width           =   1125
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Granted"
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
         Left            =   705
         TabIndex        =   61
         Top             =   2085
         Width           =   1110
      End
   End
   Begin VB.PictureBox Picture4 
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
      Height          =   6315
      Left            =   2520
      ScaleHeight     =   6315
      ScaleWidth      =   11145
      TabIndex        =   11
      Top             =   660
      Width           =   11145
      Begin MSComctlLib.ListView lvLoans 
         Height          =   1875
         Left            =   30
         TabIndex        =   13
         Top             =   180
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   3307
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAddMonthly 
         Caption         =   "Add Installment"
         Height          =   315
         Left            =   8580
         TabIndex        =   15
         Top             =   2085
         Width           =   1845
      End
      Begin MSComctlLib.ListView lvLoansDetails 
         Height          =   3795
         Left            =   0
         TabIndex        =   17
         Top             =   2430
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   6694
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAddAdjustment 
         Caption         =   "Loan Adjustments (+/-)"
         Height          =   315
         Left            =   2580
         TabIndex        =   16
         Top             =   2085
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CommandButton cmdLoanOption 
         Caption         =   "Loan Option"
         Height          =   315
         Left            =   6750
         TabIndex        =   95
         Top             =   2085
         Width           =   1845
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Deduction Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   30
         TabIndex        =   14
         Top             =   2130
         Width           =   2280
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Loans"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   0
         TabIndex        =   12
         Top             =   -30
         Width           =   1605
      End
   End
   Begin VB.PictureBox Picture2 
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
      Left            =   11505
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   87
      Top             =   6915
      Width           =   1440
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "Loans.frx":17D61
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":17EB3
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "Loans.frx":181F1
         MousePointer    =   99  'Custom
         Picture         =   "Loans.frx":18343
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label XLABID 
      BackColor       =   &H8000000D&
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   13860
      TabIndex        =   48
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label XLABPAYTYPE 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13860
      TabIndex        =   47
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label XLABACCOUNTNO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13860
      TabIndex        =   77
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label XLABTRANNO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   10
      Top             =   1530
      Width           =   1215
   End
   Begin VB.Label XLABLOANID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13860
      TabIndex        =   76
      Top             =   735
      Width           =   1215
   End
   Begin VB.Label XLABDETID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12120
      TabIndex        =   9
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmHRMSLoans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ADD IN THE TABLE HRMS_LOANMAS:
'1. ISACTIVE NVARCHAR(1)
'2. REMARKS NVARCHAR(100)

Option Explicit
Dim lngselected                                        As Integer
Attribute lngselected.VB_VarUserMemId = 1141047297
Dim rsEmpInfo                                          As ADODB.Recordset
Dim RSLOAN                                             As ADODB.Recordset
Dim ADDOREDIT                                          As String
Dim EMPLIVIL                                           As String
Attribute EMPLIVIL.VB_VarUserMemId = 1073938440
Dim LOANTYPE                                           As String
Attribute LOANTYPE.VB_VarUserMemId = 1073938442
Dim changebyNoPay                                      As Boolean

Function FindOTCodeDescription(OTCODE As Integer)
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Pay_Desc From HRMS_OTCodes Where Pay_Code = " & OTCODE & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindOTCodeDescription = Null2String(RSTMP!PAY_DESC)
    Else
        FindOTCodeDescription = ""
    End If
    Set RSTMP = Nothing
End Function

Function GETLOANDESCRIPTION(XXX As String) As String
    Dim RS                                             As ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT DESCRIPTION FROM HRMS_LoanCode WHERE CODE=" & N2Str2Null(Repleys(XXX)))
    If Not (RS.EOF Or RS.BOF) Then
        GETLOANDESCRIPTION = RS!Description
    End If

End Function

Function GETLOANCODE(XID) As String
    Dim RS                                             As ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT CODE FROM HRMS_LoanCode WHERE ID=" & XID)
    If Not (RS.EOF Or RS.BOF) Then
        GETLOANCODE = RS!CODE
    End If
End Function

Function SelectCombo(C As ComboBox) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim I                                              As Long
    For I = 0 To C.ListCount - 1
        If UCase(C.list(I)) = UCase(Trim(C.Text)) Then
            C.ListIndex = I
            Exit Function
        End If
    Next
    If C.Style = 0 Then
        C.Text = ""
    End If
    C.ListIndex = -1

End Function

Function GETLOANTRANNO() As String
    Dim rsID                                           As ADODB.Recordset
    Set rsID = New ADODB.Recordset
    Set rsID = gconDMIS.Execute("Select MAX( ISNULL(TRANNO, 0) ) as IDFIELD from HRMS_LOANMAS")
    'rsID.FIELDS(2).Value
    If rsID.FIELDS(0).Value = 0 Then
        GETLOANTRANNO = Format(1, "000000")
    Else
        GETLOANTRANNO = Format(val(N2Str2Zero(rsID![IDFIELD])) + 1, "000000")

    End If
    Set rsID = Nothing


End Function

Sub FILL_LOAN_DETAILS(XXX As String)
    Dim rsLoanmasDet                                   As ADODB.Recordset
    Set rsLoanmasDet = New ADODB.Recordset
    rsLoanmasDet.Open "SELECT convert(varchar, DEYT,101), PAYTYPE , LOANDESCRIPTION, AMOUNT, ID ,TRANNO ,LOANTYPE FROM HRMS_LOANMASDET  WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(txtEmployeeNumber) & "  AND TRANNO ='" & XXX & "' ORDER BY DEYT DESC", gconDMIS
    If Not rsLoanmasDet.EOF And Not rsLoanmasDet.BOF Then
        rsLoanmasDet.MoveFirst
        Listview_Loadval lvLoansDetails.ListItems, rsLoanmasDet
    End If

End Sub

Sub INITCBO()
    Dim RS                                             As ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT * FROM HRMS_LoanCode ORDER BY DESCRIPTION ASC")
    While Not RS.EOF
        cboLOAN_LOANTYPE.AddItem Null2String(RS!Description)
        cboLOAN_LOANTYPE.ItemData(cboLOAN_LOANTYPE.NewIndex) = RS!ID
        RS.MoveNext
    Wend



    cboDeductionOption.AddItem "1st Cut-off Period"
    cboDeductionOption.AddItem "2nd Cut-off Period"
    cboDeductionOption.AddItem "Every Cut-off Period"
    cboDeductionOption.ListIndex = 0
    'Stop
End Sub

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & EMPINFOEMPNO.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select * from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & frmHRMSEmpInfo.LABID.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set rsEmpInfo = New ADODB.Recordset
        
        'update:jbf: for managers to include in loan
        'rsEmpInfo.Open "select * from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL order by lastname,firstname,middlename asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
         rsEmpInfo.Open "select * from HRMS_EmpInfo WHERE RESIGNED IS NULL order by lastname,firstname,middlename asc", gconDMIS, adOpenForwardOnly, adLockReadOnly

    
    End If

End Sub

Sub InitGrid()
    With lvLoans
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "RefNo", "750", 0
        .ColumnHeaders.Add , , "Loan Type", "1100", 0
        .ColumnHeaders.Add , , "Account No", "0", 0
        .ColumnHeaders.Add , , "D.Granted", "1200", 0
        .ColumnHeaders.Add , , "D.Maturity", "1200", 0
        .ColumnHeaders.Add , , "Amount", "1200", 1
        .ColumnHeaders.Add , , "Monthly Ded", "0", 1
        .ColumnHeaders.Add , , "Installment", "1400", 1
        .ColumnHeaders.Add , , "Balance", "1400", 1
        .ColumnHeaders.Add , , "ID", "0", 1
        .ColumnHeaders.Add , , "Beg Balance", "1200", 1


    End With
    With lvLoansDetails
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Date", "1200"
        .ColumnHeaders.Add , , "Pay Type.", "0"
        .ColumnHeaders.Add , , "Description.", "4800"
        .ColumnHeaders.Add , , "Amount", "2000", 2
    End With
    '
    '    With grdLoanMasDet
    '        .Rows = 2
    '        .Cols = 4
    '        .Row = 0
    '        .Col = 0
    '
    '
    '        .ColAlignment(0) = 2
    '        .ColWidth(0) = 1200
    '
    '        .Col = 1
    '        .ColAlignment(1) = 2

    '        .ColWidth(1) = 2600
    '
    '        .Col = 2

    '        .ColWidth(2) = 1300
    '
    '        .ColWidth(3) = 0
    '
    '    End With

End Sub

Sub InitMemvars()
    cboLOAN_LOANTYPE = ""
    txtLoan_Acctno = ""
    txtLoan_AmountLoaned = "0.00"
    txtLoan_DateGranted.Value = LOGDATE
    txtLoan_DateStarted = DateSerial(PAY_YEAR, PAY_MONTH, Day(LOGDATE))
    txtLoan_MaturityDate.Value = LOGDATE
    txtLoan_SMonthlyDed = "0.00"
    txtLoan_SMonthlyDed = "0.00"
    txtLoan_LoanBalance = "0.00"
    txtNoPay = "0"
    txtTran_AccountNo = ""
    txtTran_Amount = "0.00"

    txtTran_Date = LOGDATE

    txtBEG_Bal = "0.00"

    XLABACCOUNTNO = ""
    XLABDETID = ""
    XLABPAYTYPE = ""
    XLABTRANNO = ""
    XLABLOANID = 0
End Sub

Sub StoreMemVars()

    On Error GoTo Errorcode
    Dim CNT                                            As Integer
    Dim Kode                                           As String
    Dim VYTDOvertime                                   As Double
    Dim vTOT                                           As Double
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Set RSLOAN = New ADODB.Recordset
        RSLOAN.Open "select * from HRMS_LoanMas where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " order by ID desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        txtName.Text = Cap1st(Null2String(rsEmpInfo!lastname)) & ", " & Cap1st(Null2String(rsEmpInfo!FIRSTNAME)) & " " & Cap1st(Null2String(rsEmpInfo!MIDDLENAME))

        If EMPINFOSHOW = True Then
            picSearch.Enabled = False
        Else
            picSearch.Enabled = True
        End If
        txtEmployeeNumber = Null2String(rsEmpInfo!EMPNO)
        txtEmpLevel = Null2String(rsEmpInfo!EMPLEVEL)
        StoreLoanMemVars txtEmployeeNumber, txtEmpLevel

        If lvLoans.ListItems.count > 0 Then
            lvLoans.ListItems(1).Selected = True
            lvLoans.ListItems(1).EnsureVisible
            lvLoans_Click
        End If

    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                     As ADODB.Recordset
    lstEmployees.Sorted = False: lstEmployees.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    
    'EDITED JBF
    'Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno ,ID from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL order by lastname+', '+firstname asc")
    '****************
    
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno ,ID from HRMS_EmpInfo WHERE RESIGNED IS NULL order by lastname+', '+firstname asc")
    
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lstEmployees.ListItems, rsEMPINFO2
        lstEmployees.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)

    Dim rsEMPINFO2                                     As ADODB.Recordset
    lstEmployees.Sorted = False: lstEmployees.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    If optSearch_EmployeeName.Value = True Then
        'Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno,ID  from HRMS_EmpInfo  where EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL and lastname+', '+firstname like'" & XXX & "%' order by lastname+', '+firstname asc")
        Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno,ID  from HRMS_EmpInfo  where  RESIGNED IS NULL and lastname+', '+firstname like'" & XXX & "%' order by lastname+', '+firstname asc")
    ElseIf optSearch_Empno.Value = True Then
        'Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno,ID  from HRMS_EmpInfo  where EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL and EMPNO like'" & XXX & "%' order by EMPNO asc")
        Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno,ID  from HRMS_EmpInfo  where  RESIGNED IS NULL and EMPNO like'" & XXX & "%' order by EMPNO asc")
    
    Else
        'Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno,ID  from HRMS_EmpInfo  where EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL and EMPNO IN (SELECT EMPNO FROM HRMS_LOANMAS WHERE LOANTYPE LIKE " & N2Str2Null(XXX & "%") & ") order by lastname+', '+firstname asc")
        Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno,ID  from HRMS_EmpInfo  where RESIGNED IS NULL and EMPNO IN (SELECT EMPNO FROM HRMS_LOANMAS WHERE LOANTYPE LIKE " & N2Str2Null(XXX & "%") & ") order by lastname+', '+firstname asc")
    End If



    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lstEmployees.ListItems, rsEMPINFO2
        lstEmployees.Refresh
    End If
End Sub

Sub updatedetails2()

End Sub

Sub StoreLoanMemVars(xxxEmpno, xxxEmpLevel)
    On Error GoTo Errorcode

    Dim CNT, crt                                       As Integer
    Dim LonType                                        As String
    Dim rsLoanMas                                      As ADODB.Recordset

    'grdLoans.Rows = 1
    lvLoans.ListItems.Clear
    lvLoansDetails.ListItems.Clear
    XLABACCOUNTNO = ""
    XLABDETID = 0
    XLABPAYTYPE = ""
    txtLoanDescription = ""
    XLABLOANID = 0
    XLABTRANNO = ""
    Set rsLoanMas = New ADODB.Recordset

    If COMPANY_CODE = "HARI" Then
        rsLoanMas.Open "SELECT TRANNO, LoanType, AcctNo, CONVERT(VARCHAR,DateGranted,101), CONVERT(VARCHAR,MaturityDate,101),AmountLoaned,MonthlyDed,SmonthlyDed,round(LoanBalance,2) , ID,BEG_BAL ,ISACTIVE FROM HRMS_LOANMAS WHERE EMPLEVEL = '" & xxxEmpLevel & "' AND EMPNO = '" & xxxEmpno & "' ORDER BY DATEGRANTED DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsLoanMas.Open "SELECT TRANNO, LoanType, AcctNo, CONVERT(VARCHAR,DateGranted,101), CONVERT(VARCHAR,MaturityDate,101),AmountLoaned,MonthlyDed,SmonthlyDed,round(LoanBalance,2) , ID,BEG_BAL FROM HRMS_LOANMAS WHERE EMPLEVEL = '" & xxxEmpLevel & "' AND EMPNO = '" & xxxEmpno & "' ORDER BY DATEGRANTED DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly

    End If

    If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
        rsLoanMas.MoveFirst
        Listview_Loadval lvLoans.ListItems, rsLoanMas
    End If
    Dim I                                              As Integer
    Dim j                                              As Integer
    If lvLoans.ListItems.count > 0 And COMPANY_CODE = "HARI" Then
        For I = 1 To lvLoans.ListItems.count - 1
            If lvLoans.ListItems(I).ListSubItems(11) = "N" Then
                For j = 1 To lvLoans.ColumnHeaders.count - 1
                    lvLoans.ListItems(I).ListSubItems(j).Bold = True
                    lvLoans.ListItems(I).ListSubItems(j).ForeColor = vbRed
                Next

            End If
        Next
    End If

    If lvLoans.ListItems.count > 0 Then
        cmdAddAdjustment.Enabled = True
        cmdAddMonthly.Enabled = True
        lstEmployees.Enabled = True
    Else
        cmdAddAdjustment.Enabled = False
        cmdAddMonthly.Enabled = False
    End If

    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub updatedetails()
    On Error GoTo adder:
    If NumericVal(txtNoPay) = 0 Then txtNoPay = 1
    txtLoan_SMonthlyDed = FormatNumber(NumericVal(txtLoan_AmountLoaned) / NumericVal(txtNoPay)) / 2
    txtLoan_SMonthlyDed = FormatNumber(NumericVal(txtLoan_AmountLoaned) / (NumericVal(txtNoPay)))
    If ADDOREDIT = "ADD" Then
        txtLoan_LoanBalance = txtLoan_AmountLoaned
    End If
    Exit Sub
adder:
    Exit Sub
    Err.Clear
End Sub

Private Sub cboLOAN_LOANTYPE_LostFocus()
    SelectCombo cboLOAN_LOANTYPE
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    ADDOREDIT = "ADD"
    InitMemvars
    picDeductions.Visible = True: picDeductions.ZOrder 0
    txtTranno = GETLOANTRANNO()
    Picture1.Visible = False
    Picture2.Visible = True
    picSearch.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    lstEmployees.Enabled = True
    ADDOREDIT = ""
    picDeductions.Visible = False
    Picture1.Visible = True
    Picture2.Visible = False
    picSearch.Enabled = True
    'StoreLoanMemVars txtEmployeeNumber, txtEmpLevel
      Picture4.Enabled = True
End Sub

Private Sub cmdDelete_Click()
    If lvLoans.SelectedItem Is Nothing Then ShowNothingToDeleteMsg: Exit Sub
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Delete", "EMPLOYEE MAINTAIN LOANS") = False Then Exit Sub

    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from HRMS_LoanMas where TRANNO = '" & XLABTRANNO & "'"
        gconDMIS.Execute "delete from HRMS_LoanMasDet where TRANNO = '" & XLABTRANNO & "'"

        LogAudit "X", "DELETE LOAN DETAIL OF EMPLOYEE", EMPLOYEE_NO
        ShowDeletedMsg
        StoreLoanMemVars txtEmployeeNumber, txtEmpLevel
    Else
        ShowNothingToDeleteMsg
    End If

    StoreMemVars

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdAddMonthly_Click()
    On Error GoTo Errorcode:
    If lvLoans.SelectedItem Is Nothing Then Exit Sub
    Dim fild                                           As String
    Dim rsloan_Balance                                 As ADODB.Recordset
    Dim LoanBalance                                    As Double
    Dim Installments_amount                            As Double
    XLABDETID = 0
    Option1.Visible = False
    Option2.Visible = False
    lstEmployees.Enabled = False
    ADDOREDIT = "ADD"
    picLoan_Details.Visible = True
    picLoan_Details.ZOrder 0
    cmdTranDelete.Enabled = False
    txtTran_AccountNo = XLABACCOUNTNO
    txtTranLoanCode = lvLoans.SelectedItem.ListSubItems(1).Text
    txtTran_DetTranno = XLABTRANNO
    Set rsloan_Balance = gconDMIS.Execute("SELECT LOANBALANCE,MONTHLYDED,SMONTHLYDED  FROM HRMS_LOANMAS WHERE TRANNO ='" & XLABTRANNO & "'")
    txtLoan_Cut_Off.Text = CUTTOFF_CODE
    txtLoan_Month = PAY_MONTH
    txtLoan_Year = PAY_YEAR
    txtTran_Date.Value = DateSerial(PAY_YEAR, PAY_MONTH, 1)
    txtTran_Date.Enabled = True

    If Not (rsloan_Balance.EOF Or rsloan_Balance.BOF) Then
        LoanBalance = N2Str2Zero(rsloan_Balance!LoanBalance)
        txtLoanBal = LoanBalance
        Installments_amount = N2Str2Zero(rsloan_Balance!SMONTHLYDED)
    End If

    If LoanBalance > Installments_amount Then
        txtTran_Amount = FormatNumber(Installments_amount)
    Else
        txtTran_Amount = FormatNumber(LoanBalance)
    End If

    txtLoanDescription = "INSTALLMENT PAYMENT FOR " & txtTranLoanCode

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdAddAdjustment_Click()

    On Error GoTo Errorcode:
    Dim fild                                           As String
    XLABDETID = 0
    lstEmployees.Enabled = False
    ADDOREDIT = "ADD"

    Option1.Visible = True
    Option2.Visible = True

    picLoan_Details.Visible = True
    picLoan_Details.ZOrder 0
    cmdTranDelete.Enabled = False
    txtTran_Amount = 0

    txtTran_AccountNo = XLABACCOUNTNO
    txtTranLoanCode = lvLoans.SelectedItem.ListSubItems(1).Text
    txtLoanDescription = " ADJUSTMENT FOR " & txtTranLoanCode
    XLABPAYTYPE = "A"
    XLABPAYTYPE.Visible = True

    Dim rsloan_Balance2                                As ADODB.Recordset

    Set rsloan_Balance2 = gconDMIS.Execute("SELECT LOANBALANCE FROM HRMS_LOANMAS WHERE TRANNO ='" & XLABTRANNO & "'")
    txtLoanBal = N2Str2Zero(rsloan_Balance2!LoanBalance)
    txtTran_DetTranno = XLABTRANNO
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    If lvLoans.SelectedItem Is Nothing Then Exit Sub
    If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE MAINTAIN LOANS") = False Then Exit Sub
    Dim fild                                           As String

    XLABLOANID = lvLoans.SelectedItem.ListSubItems(9).Text
    XLABTRANNO = lvLoans.SelectedItem.Text

    If XLABLOANID = "" Then: Exit Sub

    lstEmployees.Enabled = False

    Picture1.Visible = False: Picture2.Visible = True
    Dim LonType                                        As String
    Dim rsLoanMas                                      As ADODB.Recordset

    Set rsLoanMas = New ADODB.Recordset
    rsLoanMas.Open "SELECT * FROM HRMS_LOANMAS WHERE ID=" & XLABLOANID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsLoanMas.EOF And Not rsLoanMas.BOF Then
        changebyNoPay = False
        LonType = GETLOANDESCRIPTION(rsLoanMas!LOANTYPE)
        cboLOAN_LOANTYPE = LonType
        txtTranno = Null2String(rsLoanMas!TRANNO)

        txtLoan_Acctno = Null2String(rsLoanMas!acctno)
        txtLoan_AmountLoaned = FormatNumber(N2Str2Zero(rsLoanMas!AMOUNTLOANED))
        txtLoan_DateGranted = Null2String(rsLoanMas!DATEGRANTED)
        txtLoan_DateStarted = Null2String(rsLoanMas!DATESTARTED)
        txtLoan_MaturityDate = Null2String(rsLoanMas!MATURITYDATE)

        txtLoan_SMonthlyDed = FormatNumber(N2Str2Zero(rsLoanMas!SMONTHLYDED))
        txtBEG_Bal = FormatNumber(N2Str2Zero(rsLoanMas!beg_bal))


        If NumericVal(rsLoanMas!DEDUCTION_OPTION) = 1 Then
            cboDeductionOption.ListIndex = 0
        ElseIf NumericVal(rsLoanMas!DEDUCTION_OPTION) = 2 Then
            cboDeductionOption.ListIndex = 1
        ElseIf NumericVal(rsLoanMas!DEDUCTION_OPTION) = 3 Then
            cboDeductionOption.ListIndex = 2
        End If
        If N2Str2Zero(rsLoanMas!SMONTHLYDED) > 0 Or (N2Str2Zero(rsLoanMas!beg_bal) <> 0 And N2Str2Zero(rsLoanMas!SMONTHLYDED) <> 0) Then
            txtNoPay = FormatNumber(N2Str2Zero(rsLoanMas!beg_bal) / N2Str2Zero(rsLoanMas!SMONTHLYDED))
        End If

        txtLoan_LoanBalance = FormatNumber(N2Str2Zero(rsLoanMas!LoanBalance))
        
        picSearch.Enabled = False
        ADDOREDIT = "EDIT"
        picDeductions.Visible = True
        picDeductions.ZOrder 0
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
    On Error Resume Next
    txtsearch.SetFocus
    lstEmployees.Enabled = True
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
    'john fenix 04/29/2009
    
'    If Function_Access(LOGID, "Acess_Print", "EMPLOYEE MAINTAIN LOANS") = False Then Exit Sub
    On Error GoTo Errorcode:
    rptDeductions.Reset
    rptDeductions.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptDeductions.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    Screen.MousePointer = 11
    If MsgQuestionBox("Employees loan will be printed,Are you sure?", "Confirm Printing") = True Then
    Screen.MousePointer = 11
    
    'PrintSQLReport rptDeductions, HRMS_REPORT_PATH & "Loan Balance.rpt", "{HRMS_LOANMAS.EMPNO} = '" & txtEmployeeNumber & "'", DMIS_REPORT_Connection, 1
    'PrintSQLReport rptDeductions, HRMS_REPORT_PATH & "Loan Balance.rpt", "", DMIS_REPORT_Connection, 1
     PrintSQLReport rptDeductions, HRMS_REPORT_PATH & "Loan Balances Detail.rpt", "{HRMS_LOANMAS.EMPNO} = '" & txtEmployeeNumber & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    Else
        Exit Sub
    End If
    Screen.MousePointer = 0

Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    If cboDeductionOption.ListIndex = -1 Then
        MsgBox "Please Select Proper Deduction Option", vbInformation
        Exit Sub
    End If
    If cboLOAN_LOANTYPE.ListIndex = -1 Then
    'If cboLOAN_LOANTYPE.Text = "" Then
        MessagePop InfoVoid, "Selection Required", "Please Select Proper Deduction Option"
        cboLOAN_LOANTYPE.SetFocus
        Exit Sub
    End If

    Dim VCBOLOAN_LOANTYPE                              As String
    Dim VTXTLOAN_ACCTNO                                As String
    Dim VTXTLOAN_AMOUNTLOANED                          As String
    Dim VTXTLOAN_DATEGRANTED                           As String
    Dim VTXTLOAN_MATURITYDATE                          As String
    Dim VEMPLEVEL                                      As String
    Dim VEMPNO                                         As String
    Dim VTXTLOAN_SMONTHLYDED                           As String
    Dim VTXTLOAN_LOANBALANCE                           As Double
    Dim VTRANNO                                        As String
    Dim DEDUCTION_OPTION                               As Integer
    Dim VTXTBEG_BAL                                    As Double
    Dim VTXTLOAN_DATESTARTED                           As String

    VTXTLOAN_DATESTARTED = N2Str2Null(txtLoan_DateStarted)
    VCBOLOAN_LOANTYPE = N2Str2Null(GETLOANCODE(cboLOAN_LOANTYPE.ItemData(cboLOAN_LOANTYPE.ListIndex)))
    VTXTLOAN_ACCTNO = N2Str2Null(txtLoan_Acctno)
    VTXTLOAN_AMOUNTLOANED = Round(NumericVal(txtLoan_AmountLoaned), 2)
    VTXTLOAN_DATEGRANTED = N2Date2Null(txtLoan_DateGranted)
    VTXTLOAN_MATURITYDATE = N2Date2Null(txtLoan_MaturityDate)
    VEMPLEVEL = N2Str2Null(txtEmpLevel)
    VEMPNO = N2Str2Null(txtEmployeeNumber)
    VTXTLOAN_SMONTHLYDED = Round(NumericVal(txtLoan_SMonthlyDed), 2)
    VTXTLOAN_LOANBALANCE = Round(NumericVal(txtLoan_LoanBalance), 2)
    VTRANNO = N2Str2Null(txtTranno)
    VTXTBEG_BAL = NumericVal(txtBEG_Bal)
    DEDUCTION_OPTION = cboDeductionOption.ListIndex + 1
    If ADDOREDIT = "ADD" Then

        gconDMIS.Execute "INSERT INTO HRMS_LOANMAS " & _
                         "(TRANNO,EMPNO ,EMPLEVEL,ACCTNO,LOANTYPE,DATEGRANTED,DATESTARTED ,MATURITYDATE,AMOUNTLOANED ,MONTHLYDED ,SMONTHLYDED ,LOANBALANCE ,DEDUCTION_OPTION ,BEG_BAL )" & _
                       " VALUES (" & _
                       " " & VTRANNO & "," & VEMPNO & ", " & VEMPLEVEL & ", " & VTXTLOAN_ACCTNO & ", " & VCBOLOAN_LOANTYPE & "," & VTXTLOAN_DATEGRANTED & ", " & VTXTLOAN_DATESTARTED & ", " & VTXTLOAN_MATURITYDATE & ", " & VTXTLOAN_AMOUNTLOANED & _
                         ", " & VTXTLOAN_SMONTHLYDED * 2 & ", " & VTXTLOAN_SMONTHLYDED & _
                         ",  " & VTXTLOAN_LOANBALANCE & "," & DEDUCTION_OPTION & "," & VTXTBEG_BAL & ")"
    Else

        gconDMIS.Execute "UPDATE HRMS_LOANMAS SET " & _
                         "EMPLEVEL = " & VEMPLEVEL & ", " & _
                         "EMPNO = " & VEMPNO & ", " & _
                         "TRANNO= " & VTRANNO & ", " & _
                         "ACCTNO= " & VTXTLOAN_ACCTNO & ", " & _
                         "LOANTYPE= " & VCBOLOAN_LOANTYPE & ", " & _
                         "BEG_BAL= " & VTXTBEG_BAL & ", " & _
                         "DEDUCTION_OPTION= " & DEDUCTION_OPTION & ", " & _
                         "DATEGRANTED= " & VTXTLOAN_DATEGRANTED & ", " & _
                         "DATESTARTED = " & VTXTLOAN_DATESTARTED & ", " & _
                         "MATURITYDATE=" & VTXTLOAN_MATURITYDATE & ", " & _
                         "AMOUNTLOANED = " & VTXTLOAN_AMOUNTLOANED & ", " & _
                         "MONTHLYDED = " & VTXTLOAN_SMONTHLYDED * 2 & ", " & _
                         "SMONTHLYDED = " & VTXTLOAN_SMONTHLYDED & ", " & _
                         "LOANBALANCE = " & VTXTLOAN_LOANBALANCE & "  " & _
                       " WHERE  ID=" & XLABLOANID

    End If
    Picture1.Visible = True: Picture2.Visible = False
    picSearch.Enabled = True
    StoreLoanMemVars txtEmployeeNumber, txtEmpLevel
    picDeductions.Visible = False
    picSearch.Enabled = True
    Picture4.Enabled = True


End Sub

Private Sub cmdTranDelete_Click()
    If ShowConfirmDelete = False Then: Exit Sub
    lstEmployees.Enabled = True
    cmdTranClose.Value = True
    gconDMIS.Execute "DELETE FROM HRMS_LoanMasDet WHERE ID=" & XLABDETID
    gconDMIS.Execute "UPDATE HRMS_LOANMAS SET " & _
                     "LOANBALANCE= ISNULL(LOANBALANCE,0) + " & Round(NumericVal(txtTran_Amount), 2) & _
                   " WHERE TRANNO=" & N2Str2Null(XLABTRANNO)
    gconDMIS.Execute "DELETE FROM HRMS_PAYROLL_DET WHERE TRANNO=" & N2Str2Null(XLABTRANNO) & "  AND cut_off=" & txtLoan_Cut_Off & " AND PAY_MONTH=" & txtLoan_Month & " AND PAY_YEAR=" & txtLoan_Year
    StoreLoanMemVars txtEmployeeNumber, txtEmpLevel
    On Error GoTo errcode
    If lngselected > 0 And lngselected <= lvLoans.ListItems.count Then
        lvLoans.ListItems(lngselected).Selected = True
        lvLoans.ListItems(lngselected).EnsureVisible
        lvLoans_Click
    End If

    Exit Sub
errcode:
    Err.Clear

End Sub

Private Sub cmdTranClose_Click()
    picLoan_Details.Visible = False
    lstEmployees.Enabled = True
End Sub

Private Sub cmdTranSave_Click()
    Dim VAMOUNT                                        As Double
    Dim RSCHECK                                        As ADODB.Recordset
    If ADDOREDIT = "ADD" Then
        If (MONTH(txtTran_Date) <> PAY_MONTH Or YEAR(txtTran_Date) <> PAY_YEAR) Then
            MessagePop InfoVoid, "Invalid Date", "Please Select Date Between Cut-Off Period"
            txtTran_Date.SetFocus
            Exit Sub
        End If
        Set RSCHECK = gconDMIS.Execute("SELECT count(*) FROM HRMS_LOANMASDET WHERE CUT_OFF =" & CUTTOFF_CODE & " AND PAY_MONTH=" & PAY_MONTH & " AND PAY_YEAR=" & PAY_YEAR & " AND TRANNO= " & N2Str2Null(txtTran_DetTranno))
        If RSCHECK(0).Value > 0 Then
            MessagePop InfoVoid, "Duplicate Loan Entry", "Loan Details for Same Cut Off Exists For the Cut-Off Period. Please Select Another Cut-Off Code or Edit Information", 3000, 0
            Exit Sub
        End If
    End If

    VAMOUNT = Round(NumericVal(txtTran_Amount), 2)
    If VAMOUNT = 0 Then
        MsgBox "Loan Amount Cannot Be Zero.", vbInformation, "Zero Amount!"
        Exit Sub
    End If
    If ADDOREDIT = "ADD" Then
        If VAMOUNT > NumericVal(txtLoanBal) Then
            MsgBox "Installment Amount Exceeds Loan Balance.", vbInformation
            txtTran_Amount.SetFocus
            Exit Sub
        End If
    End If

    Dim VEMPNO, VEMPLEVEL, VACCTNO, VLOANTYPE, VDEYT, VTRANNO, VLOANDESCRIPTION, VLOANCODE, VPAYTYPE
    VPAYTYPE = N2Str2Null(XLABPAYTYPE)

    VEMPNO = N2Str2Null(txtEmployeeNumber)
    VEMPLEVEL = N2Str2Null(txtEmpLevel)
    VACCTNO = N2Str2Null(txtTran_AccountNo)
    VLOANTYPE = N2Str2Null(txtTranLoanCode)



    VDEYT = N2Date2Null(txtTran_Date)
    VTRANNO = N2Str2Null(txtTran_DetTranno)
    GENFROM = N2Date2Null(GENFROM)
    GENTO = N2Date2Null(GENTO)
    VLOANCODE = N2Str2Null(txtTranLoanCode)

    VLOANDESCRIPTION = N2Str2Null(txtLoanDescription)

    If ADDOREDIT = "ADD" Then

        gconDMIS.Execute ("INSERT INTO HRMS_LOANMASDET " & _
                          "(TRANNO, EMPNO ,EMPLEVEL,ACCTNO,LOANTYPE,AMOUNT,DEYT,LOANDESCRIPTION,PAYTYPE,CUT_OFF,PAY_MONTH,PAY_YEAR)" & _
                        " VALUES (" & VTRANNO & "," & VEMPNO & ", " & VEMPLEVEL & ", " & VACCTNO & ", " & VLOANTYPE & "," & VAMOUNT & "," & VDEYT & "," & VLOANDESCRIPTION & "," & VPAYTYPE & "," & CUTTOFF_CODE & "," & PAY_MONTH & "," & PAY_YEAR & ")   ")

        gconDMIS.Execute ("INSERT INTO HRMS_PAYROLL_DET (TRANNO, EMPLEVEL,EMPNO,TRANTYPE,DET_AMOUNT,ISADD,PAYPERIOD_FROM,PAYPERIOD_TO,DET_CODE,DET_DESC,CUT_OFF,PAY_MONTH,PAY_YEAR) VALUES (" & _
                          VTRANNO & "," & VEMPLEVEL & "," & VEMPNO & ",'L'," & NumericVal(txtTran_Amount) & ",0," & GENFROM & "," & GENTO & "," & VLOANCODE & "," & VLOANDESCRIPTION & "," & CUTTOFF_CODE & "," & PAY_MONTH & "," & PAY_YEAR & ")")

    Else

        gconDMIS.Execute "UPDATE HRMS_LOANMASDET SET " & _
                       " EMPLEVEL = " & VEMPLEVEL & ", " & _
                         "EMPNO= " & VEMPNO & ", " & _
                         "TRANNO= " & VTRANNO & ", " & _
                         "ACCTNO= " & VACCTNO & ", " & _
                         "LOANTYPE= " & VLOANTYPE & ", " & _
                         "AMOUNT= " & VAMOUNT & ", " & _
                         "LOANDESCRIPTION= " & VLOANDESCRIPTION & ", " & _
                         "CUT_OFF= " & txtLoan_Cut_Off & ", " & _
                         "PAY_MONTH= " & txtLoan_Month & ", " & _
                         "PAY_YEAR= " & txtLoan_Year & ", " & _
                         "DEYT= " & VDEYT & " " & _
                       " WHERE  ID=" & XLABDETID

        gconDMIS.Execute "UPDATE HRMS_PAYROLL_DET SET " & _
                       " TRANNO = " & VTRANNO & ", " & _
                         "EMPLEVEL= " & VEMPLEVEL & ", " & _
                         "TRANTYPE='L', " & _
                         "DET_AMOUNT= " & NumericVal(txtTran_Amount) & ", " & _
                         "PAYPERIOD_FROM= " & GENFROM & ", " & _
                         "PAYPERIOD_TO= " & GENTO & ", " & _
                         "CUT_OFF= " & txtLoan_Cut_Off & ", " & _
                         "PAY_MONTH= " & txtLoan_Month & ", " & _
                         "PAY_YEAR= " & txtLoan_Year & ", " & _
                         "DET_DESC= " & VLOANDESCRIPTION & ", " & _
                         "DET_CODE= " & VLOANCODE & " " & _
                       " WHERE TRANNO=" & VTRANNO & " AND TRANTYPE='L' AND CUT_OFF=" & CUTTOFF_CODE & " AND PAY_MONTH=" & PAY_MONTH & " AND PAY_YEAR=" & PAY_YEAR

        Dim sumSSAl                                    As Double
        Dim sumCSAl                                    As Double
        Dim sumPSAl                                    As Double
        Dim sumHDMF                                    As Double
        Dim sumOTHERLOANS                              As Double

        sumSSAl = 0
        sumCSAl = 0
        sumPSAl = 0
        sumHDMF = 0
        sumOTHERLOANS = 0


        Dim rsupdateloanpayroll                        As ADODB.Recordset
        Set rsupdateloanpayroll = New ADODB.Recordset
        Set rsupdateloanpayroll = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL_DET WHERE EMPNO = " & VEMPNO & " AND TRANTYPE='L' AND CUT_OFF=" & CUTTOFF_CODE & " AND PAY_MONTH=" & PAY_MONTH & " AND PAY_YEAR=" & PAY_YEAR)
        If Not (rsupdateloanpayroll.EOF And rsupdateloanpayroll.BOF) Then
            rsupdateloanpayroll.MoveFirst
            Do While Not rsupdateloanpayroll.EOF
                If Null2String(rsupdateloanpayroll!DET_CODE) = "SSAL" Then
                    sumSSAl = sumSSAl + N2Str2Zero(rsupdateloanpayroll!DET_AMOUNT)
                ElseIf Null2String(rsupdateloanpayroll!DET_CODE) = "CSAL" Then
                    sumCSAl = sumCSAl + N2Str2Zero(rsupdateloanpayroll!DET_AMOUNT)
                ElseIf Null2String(rsupdateloanpayroll!DET_CODE) = "PSAL" Then
                    sumPSAl = sumPSAl + N2Str2Zero(rsupdateloanpayroll!DET_AMOUNT)
                ElseIf Null2String(rsupdateloanpayroll!DET_CODE) = "HDMF" Then
                    sumHDMF = sumHDMF + N2Str2Zero(rsupdateloanpayroll!DET_AMOUNT)
                Else
                    sumOTHERLOANS = sumOTHERLOANS + N2Str2Zero(rsupdateloanpayroll!DET_AMOUNT)
                End If

                rsupdateloanpayroll.MoveNext
            Loop
        End If

        gconDMIS.Execute "UPDATE HRMS_PAYROLL SET " & _
                       " SSSSALLOAN = " & sumSSAl & "," & _
                       " SSSCALLOAN = " & sumCSAl & "," & _
                       " PAGSALLOAN = " & sumPSAl & "," & _
                       " PAGHDMFLOAN = " & sumHDMF & "," & _
                       " OTHERLOAN = " & sumOTHERLOANS & _
                       " WHERE EMPNO = " & VEMPNO & " AND CUT_OFF= " & CUTTOFF_CODE & " AND PAY_MONTH=" & PAY_MONTH & " AND PAY_YEAR=" & PAY_YEAR

        gconDMIS.Execute "UPDATE HRMS_PAYROLL SET " & _
                       " NETPAY = GROSS-(SSSSalLoan + SSSCalLoan + PagSalLoan + PAGHDMFLoan+ OtherLoan+TAX + PAGIBIG + PHILHEALTHE+ SSSE+UNDERTIME+Others +Advance +ABSENT)" & _
                       " WHERE EMPNO = " & VEMPNO & " AND CUT_OFF= " & CUTTOFF_CODE & " AND PAY_MONTH=" & PAY_MONTH & " AND PAY_YEAR=" & PAY_YEAR
    End If



    Dim RSLOAN_PAYMENT                                 As ADODB.Recordset
    Dim TOTAL_LOANPAYMENT                              As Double
    Set RSLOAN_PAYMENT = gconDMIS.Execute("SELECT SUM(AMOUNT) AMOUNT FROM HRMS_LOANMASDET WHERE TRANNO=" & VTRANNO)

    If Not (RSLOAN_PAYMENT.EOF Or RSLOAN_PAYMENT.BOF) Then

        TOTAL_LOANPAYMENT = NumericVal(RSLOAN_PAYMENT("AMOUNT"))

    End If


    gconDMIS.Execute "UPDATE HRMS_LOANMAS SET " & _
                     "LOANBALANCE= ISNULL(BEG_BAL,0)  - " & TOTAL_LOANPAYMENT & _
                   " WHERE TRANNO=" & VTRANNO

    cmdTranClose.Value = True

    StoreLoanMemVars txtEmployeeNumber, txtEmpLevel
    FILL_LOAN_DETAILS txtTran_DetTranno
    On Error GoTo errcode

    If lngselected > 0 And lngselected <= lvLoans.ListItems.count Then

        lvLoans.ListItems(lngselected).Selected = True
        lvLoans.ListItems(lngselected).EnsureVisible
        lvLoans_Click
    End If
    lstEmployees.Enabled = True
    Exit Sub
errcode:
    Err.Clear

End Sub

Private Sub cmdLoanOption_Click()
    If COMPANY_CODE <> "HARI" Then Exit Sub
    If XLABLOANID = "" Then Exit Sub
    Dim RSDETX                                         As ADODB.Recordset
    Set RSDETX = gconDMIS.Execute("SELECT ISACTIVE,  REMARKS FROM HRMS_LOANMAS WHERE ID=" & XLABLOANID)
    If Not RSDETX.EOF Or Not RSDETX.BOF Then
        txtLoanRemarks = Null2String(RSDETX!REMARKS)
        labselitem = lvLoans.SelectedItem.Index
        If Null2String(RSDETX!ISACTIVE) = "N" Then
            optLoanOption_Disable.Value = True
        Else
            optLoanOption_Enable.Value = True
        End If
    End If
    Set RSDETX = Nothing
    picLoanOption.Visible = True
    picLoanOption.ZOrder 0
    Picture4.Enabled = False
    picSearch.Enabled = False



End Sub

Private Sub Command2_Click()
    picLoanOption.Visible = False

    Picture4.Enabled = True
    picSearch.Enabled = True
End Sub

Private Sub Command3_Click()
    If MsgBox("Are you Sure You Want to Edit Loan Option", vbInformation + vbYesNo) = vbNo Then Exit Sub
    If optLoanOption_Disable.Value = True Then
        gconDMIS.Execute ("update hrms_loanmas set ISACTIVE='N' ,REMARKS=" & N2Str2Null(txtLoanRemarks) & " WHERE ID=" & XLABLOANID)
    ElseIf optLoanOption_Enable.Value = True Then
        gconDMIS.Execute ("update hrms_loanmas set ISACTIVE='Y' ,REMARKS=" & N2Str2Null(txtLoanRemarks) & " WHERE ID=" & XLABLOANID)
    Else
        MsgBox "No Loan Option To Apply", vbInformation
        Exit Sub
    End If
    StoreLoanMemVars txtEmployeeNumber, txtEmpLevel
    On Error Resume Next
    If labselitem <> "" Then
        lvLoans.ListItems(NumericVal(labselitem)).Selected = True
        lvLoans.ListItems(NumericVal(labselitem)).EnsureVisible
        lvLoans_Click
    End If
    Command2.Value = True
End Sub

Private Sub lvLoans_Click()
    If lvLoans.SelectedItem Is Nothing Then Exit Sub

    lvLoans_ItemClick lvLoans.SelectedItem

End Sub

Private Sub lvLoans_DblClick()
    If lvLoans.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
    Picture4.Enabled = False
End Sub

Private Sub lvLoans_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    lngselected = ITEM.Index
    XLABTRANNO = ITEM.Text
    XLABLOANID = ITEM.ListSubItems(9).Text
    RSLOAN.Find ("ID=" & ITEM.ListSubItems(9).Text)
    XLABACCOUNTNO = ITEM.ListSubItems(2).Text
    cmdAddAdjustment.Enabled = True
    cmdAddMonthly.Enabled = True
    lvLoansDetails.ListItems.Clear
    LOANBAL = lvLoans.SelectedItem.ListSubItems(8).Text

    txtSTORE_BEG_BAL = lvLoans.SelectedItem.ListSubItems(10).Text
    FILL_LOAN_DETAILS (XLABTRANNO)

    If COMPANY_CODE = "HARI" Then
        If lvLoans.SelectedItem.ListSubItems(11).Text = "N" Then
            'cmdLoanOption.Enabled = True
        Else
            'cmdLoanOption.Enabled = False
        End If
    Else


    End If

End Sub

Private Sub lvLoansDetails_DblClick()
    If lvLoansDetails.SelectedItem Is Nothing Then Exit Sub
    On Error GoTo Errorcode:
    XLABDETID = lvLoansDetails.SelectedItem.ListSubItems(4).Text
    If XLABDETID <> "" Then
        Dim RSDET                                      As ADODB.Recordset
        Dim RSLOANDETX                                 As ADODB.Recordset
        lstEmployees.Enabled = False
        ADDOREDIT = "EDIT"
        picLoan_Details.Visible = True
        picLoan_Details.ZOrder 0

        Set RSDET = New ADODB.Recordset
        RSDET.Open "select * from HRMS_loanmasdet where ID = " & XLABDETID, gconDMIS

        If Not (RSDET.EOF Or RSDET.BOF) Then

            Set RSLOANDETX = gconDMIS.Execute("SELECT SUM(AMOUNT) AS AMT FROM HRMS_LOANMASDET WHERE TRANNO='" & lvLoansDetails.SelectedItem.ListSubItems(5).Text & "'")

            If Not (RSLOANDETX.EOF Or RSLOANDETX.BOF) Then
                '    txtLoanBal = FormatNumber(NumericVal(LOANBAL) - NumericVal(RSLOANDETX!amt))

                txtLoanBal = FormatNumber(NumericVal(txtSTORE_BEG_BAL) - NumericVal(RSLOANDETX!amt))

            End If

            txtTran_AccountNo = Null2String(RSDET!acctno)
            txtTran_Amount = N2Str2Zero(RSDET!AMOUNT)
            txtTran_Date = Null2String(RSDET!DEYT)
            XLABPAYTYPE = Null2String(RSDET!PAYTYPE)
            txtLoanDescription = Null2String(RSDET!loandescription)
            txtTran_DetTranno = Null2String(RSDET!TRANNO)
            txtTranLoanCode = Null2String(RSDET!LOANTYPE)
            txtLoan_Cut_Off = Null2String(RSDET!CUT_OFF)
            txtLoan_Month = Null2String(RSDET!PAY_MONTH)
            txtLoan_Year = Null2String(RSDET!PAY_YEAR)
            cmdTranDelete.Enabled = True
            txtTran_Date.Enabled = False
        End If

    End If
    Exit Sub

Errorcode:
    ShowVBError

End Sub

Private Sub Option1_Click()
    txtTran_Amount = Abs(txtTran_Amount)
End Sub

Private Sub Option2_Click()
    txtTran_Amount = -1 * Abs(txtTran_Amount)
End Sub

Private Sub optSearch_EmployeeName_Click()
    txtsearch.SetFocus
End Sub

Private Sub optSearch_Empno_Click()
    txtsearch.SetFocus
End Sub

Private Sub optSearch_LoanCode_Click()
    txtsearch.SetFocus
End Sub

Private Sub txtBEG_Bal_Change()
    If ADDOREDIT = "ADD" Then

        On Error GoTo Errorcode
        If changebyNoPay = False Then
            If NumericVal(txtBEG_Bal) = 0 And NumericVal(txtLoan_SMonthlyDed) = 0 Then Exit Sub
            txtNoPay = Round(NumericVal(txtBEG_Bal) / NumericVal(txtLoan_SMonthlyDed), 2)
            txtLoan_LoanBalance = txtBEG_Bal
        End If
    End If
    Exit Sub
Errorcode:
    If Err.NUMBER = 6 Then
        Err.Clear
    End If

End Sub

Private Sub txtBEG_Bal_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtBEG_Bal_GotFocus()
    If NumericVal(txtBEG_Bal.Text) <= 0 Then txtBEG_Bal = ""
    changebyNoPay = False
    txtBEG_Bal.SelStart = 0
    txtBEG_Bal.SelLength = Len(txtBEG_Bal.Text)
End Sub

Private Sub txtBEG_Bal_LostFocus()
    txtBEG_Bal = FormatNumber(NumericVal(txtBEG_Bal))
End Sub

Private Sub txtLoan_LoanBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtLoan_LoanBalance_LostFocus()
    txtLoan_LoanBalance = FormatNumber(NumericVal(txtLoan_LoanBalance))
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            InitMemvars
            cmdAdd.Value = True
        Case vbKeyEscape
            If picLoan_Details.Visible = True Then
                picLoan_Details.Visible = False
                lstEmployees.Enabled = True
            Else
                cmdCancel.Value = True
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
    Else
        EMPLIVIL = "'E'"
    End If

    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    FillGrid


    rsrefresh
    txtsearch.Text = ""

    INITCBO

    InitGrid
    InitMemvars
    StoreMemVars
    cmdCancel_Click



    Screen.MousePointer = 0


    If lstEmployees.ListItems.count > 1 Then
        lstEmployees.ListItems(1).Selected = True


    End If

    If RTrim(LTrim(LOGCODE)) = "NET" Then
        XLABACCOUNTNO.Visible = True
        XLABDETID.Visible = True
        XLABPAYTYPE.Visible = True
        XLABTRANNO.Visible = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ADDOREDIT = ""

End Sub

Private Sub lstEmployees_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsEmpInfo.MoveFirst
    rsEmpInfo.Find ("id=" & ITEM.ListSubItems(2).Text)
    InitMemvars
    StoreMemVars
End Sub

Private Sub lstEmployees_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstEmployees
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

Private Sub lstEmployees_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtLoan_AmountLoaned_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtLoan_AmountLoaned_GotFocus()
    If NumericVal(txtLoan_AmountLoaned.Text) <= 0 Then txtLoan_AmountLoaned = ""
    txtLoan_AmountLoaned.SelStart = 0
    txtLoan_AmountLoaned.SelLength = Len(txtLoan_AmountLoaned.Text)
End Sub

Private Sub txtLoan_AmountLoaned_LostFocus()
    txtLoan_AmountLoaned = FormatNumber(NumericVal(txtLoan_AmountLoaned))
End Sub

Private Sub txtLoan_SMonthlyDed_Change()
    If ADDOREDIT = "" And IsNumeric(txtLoan_SMonthlyDed) = False Then Exit Sub
    If NumericVal(txtLoan_SMonthlyDed) > 0 Then
        txtNoPay = NumericVal(txtBEG_Bal) / NumericVal(txtLoan_SMonthlyDed)
    End If
End Sub

Private Sub txtLoan_SMonthlyDed_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtLoan_SMonthlyDed_GotFocus()
    If NumericVal(txtLoan_SMonthlyDed.Text) <= 0 Then txtLoan_SMonthlyDed = ""
    txtLoan_SMonthlyDed.SelStart = 0
    txtLoan_SMonthlyDed.SelLength = Len(txtLoan_SMonthlyDed.Text)
End Sub

Private Sub txtLoan_SMonthlyDed_LostFocus()
    txtLoan_SMonthlyDed = FormatNumber(NumericVal(txtLoan_SMonthlyDed))
End Sub

Private Sub txtNoPay_Change()
    If ADDOREDIT = "" Then Exit Sub
    If changebyNoPay = True Then
        txtBEG_Bal = FormatNumber(Round((NumericVal(txtLoan_SMonthlyDed) * NumericVal(txtNoPay)), 2))
    End If
End Sub

Private Sub txtNoPay_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtNoPay_GotFocus()
    If NumericVal(txtNoPay.Text) <= 0 Then txtNoPay = ""
    changebyNoPay = True
    txtNoPay.SelStart = 0
    txtNoPay.SelLength = Len(txtNoPay.Text)
End Sub

Private Sub txtNoPay_LostFocus()
    txtNoPay = FormatNumber(NumericVal(txtNoPay))
End Sub

Private Sub txtsearch_Change()
    If Trim(txtsearch.Text) = "" Then FillGrid Else FillSearchGrid (txtsearch.Text)
End Sub

Private Sub txtTran_Amount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtTran_Amount_GotFocus()


    If NumericVal(txtTran_Amount.Text) <= 0 Then txtTran_Amount = ""
    txtTran_Amount.SelStart = 0
    txtTran_Amount.SelLength = Len(txtTran_Amount.Text)
End Sub

Private Sub txtTran_Amount_LostFocus()
    txtTran_Amount = FormatNumber(NumericVal(txtTran_Amount))
End Sub

Private Sub txtTran_Amount_Minus_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub





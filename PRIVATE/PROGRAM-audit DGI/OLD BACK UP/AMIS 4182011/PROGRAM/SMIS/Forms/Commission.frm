VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Trans_Commission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Commissions"
   ClientHeight    =   8580
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Commission.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   9015
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7470
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   122
      Top             =   7620
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   750
         MouseIcon       =   "Commission.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   60
         MouseIcon       =   "Commission.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   3660
      ScaleHeight     =   900
      ScaleWidth      =   5490
      TabIndex        =   125
      Top             =   7620
      Width           =   5490
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   4560
         MouseIcon       =   "Commission.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3870
         MouseIcon       =   "Commission.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   3180
         MouseIcon       =   "Commission.frx":1B62
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":1CB4
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   2490
         MouseIcon       =   "Commission.frx":1FC7
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":2119
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Fin&d"
         Height          =   795
         Left            =   1800
         MouseIcon       =   "Commission.frx":2444
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":2596
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   1110
         MouseIcon       =   "Commission.frx":2890
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":29E2
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   420
         MouseIcon       =   "Commission.frx":2D3A
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":2E8C
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7605
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   8985
      TabIndex        =   8
      Top             =   0
      Width           =   9015
      Begin VB.Frame Frame4 
         Caption         =   "UNIT INCOME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   945
         Left            =   90
         TabIndex        =   50
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtAmountCommission 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   55
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txtDealerMarginAmount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txtUnitGrossIncome 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   54
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "UNIT COMMISSION"
            Height          =   210
            Left            =   1980
            TabIndex        =   52
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "DEALER UNIT MARGIN"
            Height          =   210
            Left            =   3900
            TabIndex        =   53
            Top             =   210
            Width           =   1620
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "GROSS INCOME"
            Height          =   210
            Left            =   120
            TabIndex        =   51
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         Enabled         =   0   'False
         ForeColor       =   &H00C4F4CD&
         Height          =   1350
         Left            =   0
         TabIndex        =   9
         Top             =   -120
         Width           =   9075
         Begin VB.Label lblC_Transaction 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   4050
            TabIndex        =   16
            Top             =   930
            Width           =   2490
         End
         Begin VB.Label lblC_Bank 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   6570
            TabIndex        =   17
            Top             =   930
            Width           =   2370
         End
         Begin VB.Label lblC_SAE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   60
            TabIndex        =   15
            Top             =   930
            Width           =   3975
         End
         Begin VB.Label lblC_Unit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   60
            TabIndex        =   13
            Top             =   540
            Width           =   7215
         End
         Begin VB.Label lblC_InvoiceDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "XXXXXXXXXX"
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
            Left            =   7290
            TabIndex        =   12
            Top             =   150
            Width           =   1650
         End
         Begin VB.Label lblC_ClientName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   60
            TabIndex        =   10
            Top             =   150
            Width           =   5985
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Invoice Date"
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
            Height          =   375
            Left            =   6060
            TabIndex        =   11
            Top             =   150
            Width           =   1185
         End
         Begin VB.Label LABIGNKEY 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MB100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   360
            Left            =   7290
            TabIndex        =   14
            Top             =   540
            Width           =   1650
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1125
         Left            =   30
         TabIndex        =   32
         Top             =   1140
         Width           =   8985
         Begin VB.TextBox txtCOM_COST_NETCOST 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6720
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   690
            Width           =   1905
         End
         Begin VB.TextBox txtCOM_COST_SUBSIDY 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4110
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   690
            Width           =   1905
         End
         Begin VB.TextBox txtCOM_COST_UNIT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1020
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   660
            Width           =   2205
         End
         Begin VB.TextBox txtCOM_NETAMOUNT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6720
            TabIndex        =   38
            Text            =   "0.00"
            Top             =   210
            Width           =   1905
         End
         Begin VB.TextBox txtCOM_DIS 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4110
            TabIndex        =   36
            Text            =   "0.00"
            Top             =   210
            Width           =   1905
         End
         Begin VB.TextBox txtCOM_SRP 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1020
            TabIndex        =   34
            Text            =   "0.00"
            Top             =   180
            Width           =   2205
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL"
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
            Left            =   6150
            TabIndex        =   43
            Top             =   810
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "SUBSIDY"
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
            Left            =   3360
            TabIndex        =   41
            Top             =   810
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "UNIT COST"
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
            Left            =   120
            TabIndex        =   39
            Top             =   750
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "SRP"
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
            Left            =   6330
            TabIndex        =   37
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "DISCOUNT"
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
            Left            =   3270
            TabIndex        =   35
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "UNIT PRICE"
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
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   870
         End
      End
      Begin VB.Frame pic1 
         Enabled         =   0   'False
         Height          =   705
         Left            =   120
         TabIndex        =   45
         Top             =   2190
         Width           =   8715
         Begin VB.TextBox txtDICommission 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6750
            TabIndex        =   49
            Top             =   210
            Width           =   1785
         End
         Begin VB.TextBox txtUnitCommission 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1860
            TabIndex        =   47
            Top             =   180
            Width           =   1695
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "UNIT COMMISSION %"
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
            TabIndex        =   46
            Top             =   270
            Width           =   1650
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Dealer Income %"
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
            Left            =   5280
            TabIndex        =   48
            Top             =   300
            Width           =   1365
         End
      End
      Begin VB.Frame FRA_INS 
         Caption         =   "INSURANCE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   945
         Left            =   90
         TabIndex        =   64
         Top             =   4710
         Width           =   5775
         Begin VB.TextBox txt_IncomeIns 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   68
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txt_IncomeInsCommission 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   69
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txt_IncomeInsMargin 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "DEALERS MARGIN"
            Height          =   210
            Left            =   3780
            TabIndex        =   67
            Top             =   210
            Width           =   1350
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "COMMISSION"
            Height          =   210
            Left            =   2010
            TabIndex        =   66
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "GROSS INCOME"
            Height          =   210
            Left            =   90
            TabIndex        =   65
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.Frame FRA_ACC 
         Caption         =   "ACCESSORIES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   945
         Left            =   90
         TabIndex        =   71
         Top             =   5640
         Width           =   5775
         Begin VB.TextBox txt_IncomeAcc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   75
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txt_IncomeAccCommission 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   76
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txt_IncomeAccMargin 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   77
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "DEALERS MARGIN"
            Height          =   210
            Left            =   3900
            TabIndex        =   72
            Top             =   180
            Width           =   1350
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "COMMISSION"
            Height          =   210
            Left            =   1980
            TabIndex        =   74
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "GROSS INCOME"
            Height          =   210
            Left            =   120
            TabIndex        =   73
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.Frame FRA_REG 
         Caption         =   "REGISTRATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   945
         Left            =   90
         TabIndex        =   96
         Top             =   6570
         Width           =   5775
         Begin VB.TextBox txt_IncomeReg 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   100
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txt_IncomeRegCommission 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   101
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txt_IncomeRegMargin 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "DEALERS MARGIN"
            Height          =   210
            Left            =   3900
            TabIndex        =   99
            Top             =   210
            Width           =   1350
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "COMMISSION"
            Height          =   210
            Left            =   1980
            TabIndex        =   98
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "GROSS INCOME"
            Height          =   210
            Left            =   60
            TabIndex        =   97
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1755
         Left            =   5550
         ScaleHeight     =   1725
         ScaleWidth      =   3375
         TabIndex        =   18
         Top             =   0
         Width           =   3405
         Begin VB.Label LABTRAN 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BPO"
            Height          =   225
            Left            =   3390
            TabIndex        =   27
            Top             =   690
            Width           =   675
         End
         Begin VB.Label LABVI_NO 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            Height          =   285
            Left            =   3390
            TabIndex        =   21
            Top             =   90
            Width           =   675
         End
         Begin VB.Label LABID 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   285
            Left            =   3390
            TabIndex        =   24
            Top             =   390
            Width           =   675
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Monthly Amortization"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   90
            TabIndex        =   28
            Top             =   930
            Width           =   1785
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Rate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   90
            TabIndex        =   25
            Top             =   630
            Width           =   390
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Term"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   60
            TabIndex        =   19
            Top             =   120
            Width           =   450
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Finance Balanced"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   90
            TabIndex        =   22
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblC_Term 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Term"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Left            =   1950
            TabIndex        =   20
            Top             =   90
            Width           =   1395
         End
         Begin VB.Label txtBalToFinanced 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Term"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   270
            Left            =   1950
            TabIndex        =   23
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label lblC_Rate 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Term"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   270
            Left            =   1950
            TabIndex        =   26
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label lblC_monthlyAmort 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Term"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Left            =   1950
            TabIndex        =   29
            Top             =   930
            Width           =   1395
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Downpayment"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   90
            TabIndex        =   30
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label txtDownpayment 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Term"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   1950
            TabIndex        =   31
            Top             =   1230
            Width           =   1395
         End
      End
      Begin VB.Frame pic3 
         Height          =   4635
         Left            =   5910
         TabIndex        =   78
         Top             =   2880
         Width           =   2955
         Begin VB.CommandButton Command1 
            Caption         =   "Free Beeies Detail"
            Height          =   405
            Left            =   1230
            TabIndex        =   95
            Top             =   3690
            Width           =   1635
         End
         Begin VB.TextBox txtCM_Days 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   90
            Text            =   "0.00"
            Top             =   2250
            Width           =   1605
         End
         Begin VB.TextBox txt_FREE_STD 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   84
            Text            =   "0.00"
            Top             =   990
            Width           =   1605
         End
         Begin VB.TextBox txt_FREE_OTHER 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   88
            Text            =   "0.00"
            Top             =   1830
            Width           =   1605
         End
         Begin VB.TextBox txt_FREE_ADD 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   86
            Text            =   "0.00"
            Top             =   1425
            Width           =   1605
         End
         Begin VB.TextBox txtCM_COSOFMONEY 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   92
            Text            =   "0.00"
            Top             =   2655
            Width           =   1605
         End
         Begin VB.TextBox txtDatePullOut 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1530
            TabIndex        =   82
            Text            =   "0.00"
            Top             =   450
            Width           =   1305
         End
         Begin VB.TextBox txtDateReleased 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   81
            Text            =   "0.00"
            Top             =   450
            Width           =   1305
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   94
            Text            =   "0.00"
            Top             =   3210
            Width           =   1605
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "DAYS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   750
            TabIndex        =   89
            Top             =   2310
            Width           =   390
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ADD.FREEBIES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   135
            TabIndex        =   85
            Top             =   1560
            Width           =   1035
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "STD. FREEBIES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   83
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "OTHERS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   540
            TabIndex        =   87
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "COST OF MONEY"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   390
            TabIndex        =   91
            Top             =   2640
            Width           =   765
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "PULL OUT DATE"
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
            Left            =   1530
            TabIndex        =   80
            Top             =   210
            Width           =   1260
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "DATE RELEASED"
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
            Left            =   120
            TabIndex        =   79
            Top             =   210
            Width           =   1275
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            X1              =   120
            X2              =   2835
            Y1              =   915
            Y2              =   915
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   450
            TabIndex        =   93
            Top             =   3300
            Width           =   705
         End
      End
      Begin VB.Frame FRA_PO 
         Caption         =   "BANK PO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   945
         Left            =   90
         TabIndex        =   57
         Top             =   3780
         Width           =   5775
         Begin VB.TextBox txt_IncomePO 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   61
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txt_IncomePOComission 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1980
            TabIndex        =   62
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox txt_IncomePOMargin 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "DEALERS MARGIN"
            Height          =   210
            Left            =   3900
            TabIndex        =   60
            Top             =   210
            Width           =   1350
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "COMMISSION"
            Height          =   210
            Left            =   1980
            TabIndex        =   58
            Top             =   180
            Width           =   960
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "GROSS INCOME"
            Height          =   210
            Left            =   90
            TabIndex        =   59
            Top             =   210
            Width           =   1170
         End
      End
   End
   Begin VB.PictureBox picFreeBeeies 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4005
      Left            =   300
      ScaleHeight     =   3975
      ScaleWidth      =   8385
      TabIndex        =   103
      Top             =   1800
      Visible         =   0   'False
      Width           =   8415
      Begin XtremeReportControl.ReportControl lvFree 
         Height          =   3420
         Left            =   60
         TabIndex        =   107
         Top             =   450
         Width           =   8250
         _Version        =   655364
         _ExtentX        =   14552
         _ExtentY        =   6032
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         MultipleSelection=   0   'False
         SkipGroupsFocus =   0   'False
         FreezeColumnsCount=   3
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7950
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ADD FREE BEEIES"
         Height          =   315
         Left            =   60
         TabIndex        =   106
         Top             =   60
         Width           =   1995
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   405
         Left            =   0
         TabIndex        =   104
         Top             =   0
         Width           =   8385
         _Version        =   655364
         _ExtentX        =   14790
         _ExtentY        =   714
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
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
      End
   End
   Begin VB.PictureBox picAccessories 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   2422
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2385
      ScaleWidth      =   4140
      TabIndex        =   108
      Top             =   2595
      Visible         =   0   'False
      Width           =   4170
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Taxable"
         Height          =   255
         Left            =   1350
         TabIndex        =   116
         Top             =   1590
         Width           =   885
      End
      Begin VB.TextBox txtFreeBeeiesAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1350
         MaxLength       =   7
         TabIndex        =   115
         Text            =   "0.00"
         Top             =   1170
         Width           =   2625
      End
      Begin VB.CommandButton cmdCancelDetailProduct 
         Caption         =   "X"
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
         Index           =   1
         Left            =   3690
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.ComboBox cboFreeBeeies 
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
         Height          =   345
         Left            =   1350
         TabIndex        =   113
         Top             =   750
         Width           =   2670
      End
      Begin VB.CommandButton cmdCancelDetailProduct 
         Height          =   495
         Index           =   0
         Left            =   3360
         MouseIcon       =   "Commission.frx":31EB
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":333D
         Style           =   1  'Graphical
         TabIndex        =   119
         TabStop         =   0   'False
         ToolTipText     =   "Exit Entry"
         Top             =   1680
         Width           =   555
      End
      Begin VB.CommandButton cmdOkMaterials 
         Height          =   495
         Left            =   2820
         MouseIcon       =   "Commission.frx":36A3
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":37F5
         Style           =   1  'Graphical
         TabIndex        =   120
         TabStop         =   0   'False
         ToolTipText     =   "Save Entry"
         Top             =   1680
         Width           =   555
      End
      Begin VB.CommandButton cmdDelFree 
         Height          =   495
         Left            =   2280
         MouseIcon       =   "Commission.frx":3B45
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":3C97
         Style           =   1  'Graphical
         TabIndex        =   118
         TabStop         =   0   'False
         ToolTipText     =   "Delete Entry"
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label LABDETID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   225
         Left            =   240
         TabIndex        =   117
         Top             =   1800
         Visible         =   0   'False
         Width           =   105
      End
      Begin XtremeShortcutBar.ShortcutCaption capAccessories 
         Height          =   330
         Left            =   0
         TabIndex        =   109
         Top             =   0
         Width           =   4155
         _Version        =   655364
         _ExtentX        =   7329
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Free Beeies"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label64 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Free Beeies"
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
         Height          =   225
         Left            =   270
         TabIndex        =   112
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
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
         Height          =   225
         Index           =   3
         Left            =   540
         TabIndex        =   114
         Top             =   1230
         Width           =   705
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1410
         TabIndex        =   111
         Top             =   390
         Width           =   3795
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   8595
      Left            =   0
      ScaleHeight     =   8595
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9015
      Begin XtremeReportControl.ReportControl ReportControl1 
         Height          =   7530
         Left            =   60
         TabIndex        =   7
         Top             =   900
         Width           =   8880
         _Version        =   655364
         _ExtentX        =   15663
         _ExtentY        =   13282
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         MultipleSelection=   0   'False
         SkipGroupsFocus =   0   'False
         FreezeColumnsCount=   3
      End
      Begin VB.CommandButton Command3 
         Caption         =   "X"
         Height          =   315
         Left            =   8550
         TabIndex        =   121
         Top             =   30
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "DATE"
         Height          =   375
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "CS#"
         Height          =   375
         Left            =   3450
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "CLIENT"
         Height          =   375
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "INV #"
         Height          =   375
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtSearchCommission 
         Height          =   375
         Left            =   5190
         TabIndex        =   6
         Top             =   450
         Width           =   3645
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   345
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9195
         _Version        =   655364
         _ExtentX        =   16219
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Search"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_Commission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS                                                                As ADODB.Recordset
Dim Free_Type                                                         As String
Private WithEvents SearchMaster                                       As frmSMIS_Mis_SearchMaster
Attribute SearchMaster.VB_VarHelpID = -1
Dim AddorEdit                                                         As String
Dim stdFree                                                           As Double
Dim affree                                                            As Double
Dim otherfree                                                         As Double

Function ComputeUnitCommission()
    Dim NETSALESPRICE
    Dim NetCostPrice
    Dim NETCOSTOFMONEY
    Dim NETSUBSIDY
    Dim TOTALCOMMISSION
    Dim DEALERMARGIN
    Dim UNITCOMMISSION
    Dim GROSSINCOME
    Dim TOTALCOSTOFMONEY
    NETSALESPRICE = NumericVal(txtCOM_NETAMOUNT) / 1.12
    NetCostPrice = NumericVal(txtCOM_COST_UNIT) / 1.12
    NETCOSTOFMONEY = NumericVal(txtCM_COSOFMONEY)
    NETSUBSIDY = NumericVal(txtCOM_COST_SUBSIDY)
    UNITCOMMISSION = NumericVal(txtUnitCommission / 100)
    TOTALCOSTOFMONEY = NumericVal(Text1)
    GROSSINCOME = NumericVal(txt_IncomePOMargin) + NumericVal(txt_IncomeAccMargin) + NumericVal(txt_IncomeRegMargin) + NumericVal(txt_IncomeInsMargin)
    ''UNIT COMMISSION
    TOTALCOMMISSION = (NETSALESPRICE - NetCostPrice - TOTALCOSTOFMONEY) * UNITCOMMISSION

    txtAmountCommission = FormatNumber(TOTALCOMMISSION)
    ''DEALER MARGIN
    DEALERMARGIN = NETSALESPRICE - NetCostPrice - TOTALCOSTOFMONEY + NETSUBSIDY - TOTALCOMMISSION
    txtDealerMarginAmount = FormatNumber(DEALERMARGIN)
    ''FINANCING INCOME
    txt_IncomePO = FormatNumber(txtBalToFinanced * (txtDICommission / 100))
    ''GROSS INCOME
    txtUnitGrossIncome = FormatNumber(GROSSINCOME + DEALERMARGIN)

End Function

Sub ConfigGrid()
    ReportControlAddColumnHeader lvFree, "DESCRIPTION, COST, NETAMOUNT,FREE BEEIES TYPE "
    ResizeColumnHeader lvFree, "6,4,4"
    ReportControlPaintManager lvFree
    ReportControlAddColumnHeader ReportControl1, "Date, CS# , VI#, Customer Name, Model"
    ResizeColumnHeader ReportControl1, "10,10,10,40,30"
    ReportControlPaintManager ReportControl1


End Sub

Sub GetFreeBeeies()
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select type from smis_vacc where AccessoriesName=" & N2Str2Null(cboFreeBeeies))
    Label26 = ""
    Free_Type = ""
    If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
        Free_Type = Null2String(TEMPRS!Type)
    End If

    If Free_Type = "ST" Then
        Label26 = "STANDARD FREEBEEIES"
    ElseIf Free_Type = "AF" Then
        Label26 = "ADDITIONAL FREEBEEIES"
    End If

End Sub

Sub InitMemVars()
    picSaves.Visible = False
    picAdds.Visible = True
    Dim txt                                                           As Control
    For Each txt In Me.ControlS
        If TypeOf txt Is TextBox Then
            txt.Text = "0.00"
        End If
    Next
    txtDateReleased = ""
    txtDatePullOut = ""
    txtSearchCommission = ""
End Sub

Sub LoadCustomerDetail(VI_NO)
    Dim TEMPRS                                                        As ADODB.Recordset

    Set TEMPRS = gconDMIS.Execute("Select * from SMIS_SALESORDER where VI_NO='" & VI_NO & "'")

    If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
        lblC_InvoiceDate = Null2String(TEMPRS!InvoicedDate)
        lblC_ClientName = Null2String(TEMPRS!CustName)
        If Null2String(TEMPRS!TERM) = "BPO" Then
            lblC_Transaction = "BANK PO"
            lblC_Bank = Null2String(TEMPRS!financingco)
            lblC_Term = FormatNumber(NumericVal(TEMPRS!MONTHSAMORT))
            lblC_monthlyAmort = FormatNumber(NumericVal(TEMPRS!NETMOAMORT))
            txtBalToFinanced = FormatNumber(NumericVal(TEMPRS!BALTOFINANCED))
            lblC_Rate = FormatNumber(NumericVal(TEMPRS!AOR))
            txtDownpayment = FormatNumber(NumericVal(TEMPRS!DownPayment))
        ElseIf Null2String(TEMPRS!TERM) = "F" Then
            lblC_Transaction = "FINANCING"
            lblC_Bank = Null2String(TEMPRS!financingco)
            lblC_Term = FormatNumber(NumericVal(TEMPRS!MONTHSAMORT))
            lblC_monthlyAmort = FormatNumber(NumericVal(TEMPRS!NETMOAMORT))
            txtBalToFinanced = FormatNumber(NumericVal(TEMPRS!BALTOFINANCED))
            lblC_Rate = FormatNumber(NumericVal(TEMPRS!AOR))
            txtDownpayment = FormatNumber(NumericVal(TEMPRS!DownPayment))
        Else
            lblC_Transaction = "CASH/COMPANY PO"
            lblC_Bank = ""
            lblC_Bank = ""
            lblC_Term = "0"
            lblC_monthlyAmort = "0.00"
            txtBalToFinanced = "0.00"
            lblC_Rate = "0.00"
            txtDownpayment = "0.00"
        End If

        txt_IncomeIns = FormatNumber(NumericVal(TEMPRS!INSURANCE))
        LABTRAN = Null2String(TEMPRS!TERM)
        LABVI_NO = Null2String(TEMPRS!VI_NO)

        lblC_SAE = Null2String(TEMPRS!salesae)
        If Null2String(TEMPRS!DATERELEASED) <> "" Then: txtDateReleased = Format(TEMPRS!DATERELEASED, "mmm dd yyyy")
        txtCOM_SRP = FormatNumber(NumericVal(TEMPRS!SALESPRICE))
        txtCOM_DIS = FormatNumber(NumericVal(TEMPRS!DISCOUNT))
        txtCOM_NETAMOUNT = FormatNumber(NumericVal(TEMPRS!NETSALESPRICE))
    Else
        MsgBox "Vehicle Information Missing! ", vbCritical
    End If
End Sub

Sub LoadMrrDetail(MRRID)
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim SQL                                                           As String
    lvFree.Records.DeleteAll
    Set TEMPRS = gconDMIS.Execute("SELECT DESCRIPTION , COST , CASE WHEN TAXABLE = 1 THEN  COST/1.12 ELSE COST END ,TYPE , ID  , TAXABLE FROM SMIS_MRRINV_DETAIL WHERE  IGNKEYNO='" & MRRID & "'")
    flex_FillReportView TEMPRS, lvFree
    stdFree = 0: otherfree = 0: affree = 0
    SQL = "SELECT     "
    SQL = SQL & " SUM(case when taxable = 1 then  Cost/1.12 else cost end ) as NETQ ,TYPE "
    SQL = SQL & " From SMIS_MrrInv_Detail"
    SQL = SQL & " WHERE     (IgnKeyNo = '" & MRRID & "')"
    SQL = SQL & " group by type "
    Set TEMPRS = gconDMIS.Execute(SQL)
    While Not TEMPRS.EOF
        If Null2String(TEMPRS!Type) = "ST" Then
            txt_FREE_STD = FormatNumber(NumericVal(TEMPRS!NETQ))
        ElseIf Null2String(TEMPRS!Type) = "AF" Then
            txt_FREE_ADD = FormatNumber(NumericVal(TEMPRS!NETQ))
        Else
            txt_FREE_OTHER = FormatNumber(NumericVal(TEMPRS!NETQ))
        End If
        TEMPRS.MoveNext
    Wend

End Sub

Sub LoadVehicleDetail(MRRID)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("Select * from SMIS_MRRINV where IGNKEY='" & MRRID & "'")
    If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
        If Null2String(TEMPRS!PullOutDate) <> "" Then: txtDatePullOut = Format(TEMPRS!PullOutDate, "mmm dd yyyy")
        txtCOM_COST_UNIT = FormatNumber(NumericVal(TEMPRS!PurchPrice))
        txtCOM_COST_SUBSIDY = FormatNumber(NumericVal(TEMPRS!MMPCSUBs))
        txtCOM_COST_NETCOST = FormatNumber(NumericVal(txtCOM_COST_UNIT) - NumericVal(txtCOM_COST_SUBSIDY))
        LABIGNKEY = Null2String(TEMPRS!ignkey)
        lblC_Unit = Null2String(TEMPRS!YEER) + " " + Null2String(TEMPRS!Make) + " " + Null2String(TEMPRS!DESCRIPT)
    Else
        MsgBox "Vehicle Information Missing! Please Update Your Recieving Entry", vbCritical
    End If
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM SMIS_COMMISSION order by id desc", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemVars()
    If Not RS.EOF Or Not RS.BOF Then
        LoadVehicleDetail (Null2String(RS!IGNKEYNO))
        LoadCustomerDetail (Null2String(RS!VI_NO))
        LoadMrrDetail (LABIGNKEY)
        If IsDate(txtDateReleased) = True And IsDate(txtDatePullOut) = True Then
            txtCM_Days = DateDiff("d", txtDatePullOut, txtDateReleased)
        End If
        labid = Null2String(RS!ID)
        LABIGNKEY = Null2String(RS!IGNKEYNO)
        LABTRAN = Null2String(RS!TERM)
        LABVI_NO = Null2String(RS!VI_NO)
        If txtCM_Days <= 4 Then
            txtCM_Days = 0
        End If
        txtCM_COSOFMONEY = NumericVal(RS!COSTOFMONEY)
        txtUnitCommission = NumericVal(RS!UNITCOMM)




        txt_IncomePO = FormatNumber(NumericVal(RS!FIN_GROSS))
        txt_IncomePOComission = FormatNumber(NumericVal(RS!FIN_AGENT))
        txt_IncomePOMargin = FormatNumber(NumericVal(RS!FIN_NET))


        txt_IncomeIns = FormatNumber(NumericVal(RS!INS_GROSS))
        txt_IncomeInsCommission = FormatNumber(NumericVal(RS!INS_AGENT))
        txt_IncomeInsMargin = FormatNumber(NumericVal(RS!INS_NET))

        txt_IncomeAcc = FormatNumber(NumericVal(RS!ACC_GROSS))
        txt_IncomeAccCommission = FormatNumber(NumericVal(RS!ACC_AGENT))
        txt_IncomeAccMargin = FormatNumber(NumericVal(RS!ACC_NET))

        txt_IncomeReg = FormatNumber(NumericVal(RS!REG_GROSS))
        txt_IncomeRegCommission = FormatNumber(NumericVal(RS!REG_AGENT))
        txt_IncomeRegMargin = FormatNumber(NumericVal(RS!REG_NET))

        txtDICommission = FormatNumber(NumericVal(RS!DI))
        txtDatePullOut_Change
        ComputeCostOfMoney
    Else
        ShowNoRecord
        If MsgBox(" There are No Records for the Commission ! Do you Want to Add New?", vbQuestion + vbYesNo) = vbYes Then
            cmdAdd.Value = True
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub cboFreeBeeies_Change()
    GetFreeBeeies
End Sub

Private Sub cboFreeBeeies_Click()
    GetFreeBeeies
End Sub

Private Sub ComputeCostOfMoney()
    On Error Resume Next
    txtCM_COSOFMONEY = FormatNumber((txtCOM_COST_NETCOST * txtCM_Days * 0.1) / 360)
    If txtCM_COSOFMONEY <= 0 Then: txtCM_COSOFMONEY = "0.00"
    Text1 = FormatNumber(NumericVal(txtCM_COSOFMONEY) + NumericVal(txt_FREE_OTHER) + NumericVal(txt_FREE_ADD) + NumericVal(txt_FREE_STD))
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "COMMISSION") = False Then Exit Sub
    SearchMaster.SearchForRELEASED
    SearchMaster.Show 1
End Sub

Private Sub cmdCancel_Click()
    pic1.Enabled = False: pic3.Enabled = False
    picSaves.Visible = False: picAdds.Visible = True: AddorEdit = ""
    Picture1.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdCancelDetailProduct_Click(Index As Integer)
    ShowHidePictureBox2 picAccessories, False, Picture1
    LABDETID = 0
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "COMMISSION") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute " DELETE FROM SMIS_COMMISSION WHERE ID=" & labid
        rsRefresh
        StoreMemVars
        LogAudit "X", "Vehicle Commission", "CS:" & LABIGNKEY & "CLIENT:" & lblC_ClientName
        MessagePop InfoOk, "DELETED", "RECORD SUCESSFULLY DELETE", 1000, 2
    End If
End Sub

Private Sub cmdDelFree_Click()
    If ShowConfirmDelete = True Then
        Free_Type = ""
        gconDMIS.Execute ("DELETE FROM SMIS_MRRINV_DETAIL WHERE ID=" & LABDETID)
        LABDETID = 0
        LoadMrrDetail LABIGNKEY
        ShowHidePictureBox2 picAccessories, False, Picture1
    End If
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "COMMISSION") = False Then Exit Sub
    pic1.Enabled = True: pic3.Enabled = True
    Picture1.Enabled = True
    picSaves.Visible = True: picAdds.Visible = False
    AddorEdit = "EDIT"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    ShowHidePictureBox2 Picture2, True
    On Error Resume Next
    txtSearchCommission.SetFocus
End Sub

Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdOkMaterials_Click()

    If NumericVal(txtFreeBeeiesAmount) <= 0 Then
        ShowIsRequiredMsg " Free Beeies Amount"
        txtFreeBeeiesAmount.SetFocus
        Exit Sub
    End If

    If LABDETID = 0 Then
        Dim TEMPRS                                                    As ADODB.Recordset
        Set TEMPRS = gconDMIS.Execute("Select COUNT(*) from SMIS_MRRINV_DETAIL WHERE IgnKeyNo='" & LABIGNKEY & "'and description=" & N2Str2Null(cboFreeBeeies))
        If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
            If TEMPRS.Fields(0).Value > 0 Then
                MsgBox " Free Beeies Detail Already Exists"
                Exit Sub
            End If
        End If
    End If

    Dim SQL                                                           As String
    If Free_Type = "" Then: Free_Type = "ST"
    If LABDETID = 0 Then
        SQL = "insert into smis_mrrinv_detail(ignkeyno,description,cost,type,taxable)values("
        SQL = SQL & N2Str2Null(LABIGNKEY) & "," & N2Str2Null(cboFreeBeeies) & "," & NumericVal(txtFreeBeeiesAmount) & "," & N2Str2Null(Free_Type) & "," & Check1.Value & ")"
    Else
        SQL = "update smis_mrrinv_detail SET Cost=" & NumericVal(txtFreeBeeiesAmount) & "" _
            & ",taxable=" & Check1.Value & " WHERE ID=" & LABDETID
    End If
    gconDMIS.Execute SQL
    LoadMrrDetail LABIGNKEY
    ShowHidePictureBox2 picAccessories, False, Picture1
    LABDETID = 0
End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdSave_Click()
    If IsDate(txtDatePullOut) = False Then
        ShowIsRequiredMsg "Pull Out Date"
        txtDatePullOut.SetFocus
        Exit Sub
    End If
    If IsDate(txtDateReleased) = False Then
        ShowIsRequiredMsg "Release Date"
        txtDateReleased.SetFocus
        Exit Sub
    End If


    If IsNumeric(txtDICommission) = True Then
        '    If txtDICommission <= 0 Then
        '        ShowIsRequiredMsg "Dealer Income"'

        'txtDICommission.SetFocus
        'Exit Sub
        'End If
    Else
        ShowIsRequiredMsg "Dealer Income"

        txtDICommission.SetFocus
        Exit Sub
    End If

    Dim SQL                                                           As String
    SQL = ""
    If AddorEdit = "ADD" Then

        SQL = " INSERT INTO SMIS_COMMISSION ("
        SQL = SQL & " TDAYS, IGNKEYNO , VI_NO , UNITCOST, SUBSIDY, COSTOFMONEY, NETCOST, DISCOUNT, "
        SQL = SQL & " STD_FREEBEEIES,ADD_FREEBEEIES,OTHER,TERM,"
        SQL = SQL & " UNITCOMM , DI ,"
        SQL = SQL & " FIN_GROSS,FIN_AGENT,FIN_NET,"
        SQL = SQL & " INS_GROSS,INS_AGENT,INS_NET,"
        SQL = SQL & " ACC_GROSS,ACC_AGENT,ACC_NET,"
        SQL = SQL & " REG_GROSS , REG_AGENT, REG_NET,"
        SQL = SQL & " USRECODE ,LASTUPDATE ) VALUES( "
        SQL = SQL & NumericVal(txtCM_Days) & "," & N2Str2Null(LABIGNKEY) & "," & N2Str2Null(LABVI_NO) & "," & NumericVal(txtCOM_COST_UNIT) & "," & NumericVal(txtCOM_COST_SUBSIDY) & "," & NumericVal(txtCM_COSOFMONEY) & "," & NumericVal(txtCOM_COST_NETCOST) & "," & NumericVal(txtCOM_DIS) & ","
        SQL = SQL & NumericVal(txt_FREE_STD) & "," & NumericVal(txt_FREE_ADD) & "," & NumericVal(txt_FREE_OTHER) & ", '" & LABTRAN & "' ,"
        SQL = SQL & NumericVal(txtUnitCommission) & "," & NumericVal(txtDICommission) & ","
        SQL = SQL & NumericVal(txt_IncomePO) & "," & NumericVal(txt_IncomePOComission) & "," & NumericVal(txt_IncomePOMargin) & ","
        SQL = SQL & NumericVal(txt_IncomeIns) & "," & NumericVal(txt_IncomeInsCommission) & "," & NumericVal(txt_IncomeInsMargin) & ","
        SQL = SQL & NumericVal(txt_IncomeAcc) & "," & NumericVal(txt_IncomeAccCommission) & "," & NumericVal(txt_IncomeAccMargin) & ","
        SQL = SQL & NumericVal(txt_IncomeReg) & "," & NumericVal(txt_IncomeRegCommission) & "," & NumericVal(txt_IncomeRegMargin) & ","
        SQL = SQL & N2Str2Null(LOGCODE) & ",'" & Date & "')"
        LogAudit "A", "Vehicle Commission", "CS:" & LABIGNKEY & "CLIENT:" & lblC_ClientName
        MessagePop RecSaveOk, "ADDED", "RECORD SUCESSFULLY ADDED", 1000, 2

    Else
        SQL = " UPDATE SMIS_COMMISSION"
        SQL = SQL & " SET IGNKEYNO ='" & LABIGNKEY & "'"
        SQL = SQL & " ,VI_NO ='" & LABVI_NO & "'"
        SQL = SQL & " ,UNITCOST =" & NumericVal(txtCOM_COST_UNIT)
        SQL = SQL & " ,UNITCOMM =" & NumericVal(txtUnitCommission)
        SQL = SQL & " ,DI =" & NumericVal(txtDICommission)
        SQL = SQL & " ,TDAYS =" & NumericVal(txtCM_Days)

        SQL = SQL & " ,SUBSIDY =" & NumericVal(txtCOM_COST_SUBSIDY)
        SQL = SQL & " ,COSTOFMONEY =" & NumericVal(txtCM_COSOFMONEY)
        SQL = SQL & " ,NETCOST =" & NumericVal(txtCOM_COST_NETCOST)
        SQL = SQL & " ,DISCOUNT =" & NumericVal(txtCOM_DIS)
        SQL = SQL & " ,STD_FREEBEEIES =" & NumericVal(txt_FREE_STD)
        SQL = SQL & " ,ADD_FREEBEEIES =" & NumericVal(txt_FREE_ADD)
        SQL = SQL & " ,OTHER =" & NumericVal(txt_FREE_OTHER)
        SQL = SQL & " ,TERM ='" & LABTRAN & "'"

        SQL = SQL & " ,FIN_GROSS =" & NumericVal(txt_IncomePO)
        SQL = SQL & " ,FIN_AGENT =" & NumericVal(txt_IncomePOComission)
        SQL = SQL & " ,FIN_NET =" & NumericVal(txt_IncomePOMargin)

        SQL = SQL & " ,INS_GROSS =" & NumericVal(txt_IncomeIns)
        SQL = SQL & " ,INS_AGENT =" & NumericVal(txt_IncomeInsCommission)
        SQL = SQL & " ,INS_NET =" & NumericVal(txt_IncomeInsMargin)

        SQL = SQL & " ,ACC_GROSS =" & NumericVal(txt_IncomeAcc)
        SQL = SQL & " ,ACC_AGENT =" & NumericVal(txt_IncomeAccCommission)
        SQL = SQL & " ,ACC_NET =" & NumericVal(txt_IncomeAccMargin)

        SQL = SQL & " ,REG_GROSS =" & NumericVal(txt_IncomeReg)
        SQL = SQL & " ,REG_AGENT =" & NumericVal(txt_IncomeRegCommission)
        SQL = SQL & " ,REG_NET =" & NumericVal(txt_IncomeRegMargin)

        SQL = SQL & " ,USRECODE =" & N2Str2Null(LOGCODE)
        SQL = SQL & " ,LASTUPDATE ='" & Date & "'"
        SQL = SQL & " WHERE ID=" & labid
        LogAudit "E", "Vehicle Commission", "CS:" & LABIGNKEY & "CLIENT:" & lblC_ClientName
        MessagePop RecSaveInfo, "UPDATED", "RECORD SUCESSFULLY UPDATED", 1000, 2

    End If
    gconDMIS.Execute SQL

    gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET datereleased='" & txtDateReleased & "' , PULLOUTDATE='" & txtDatePullOut & "' where IGNKEY='" & LABIGNKEY & "'")

    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET datereleased='" & txtDateReleased & "' where VI_NO='" & LABVI_NO & "'")

    RS.Requery
    If AddorEdit = "EDIT" Then
        RS.Find ("ID=" & labid)
    End If
    cmdCancel.Value = True
End Sub

Private Sub Command1_Click()
    ShowHidePictureBox2 picFreeBeeies, True
End Sub

Private Sub Command2_Click()
    ShowHidePictureBox2 picFreeBeeies, False
End Sub

Private Sub Command3_Click()
    ShowHidePictureBox2 Picture2, False
End Sub

Private Sub Command5_Click()
    cboFreeBeeies = "": txtFreeBeeiesAmount = "0.00": Check1.Value = 0: Label26 = "": Free_Type = ""
    cboFreeBeeies.Enabled = True
    ShowHidePictureBox2 picAccessories, True, Picture1: cmdDelFree.Enabled = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set SearchMaster = New frmSMIS_Mis_SearchMaster
    ConfigGrid
    InitMemVars
    rsRefresh
    Combo_Loadval cboFreeBeeies, gconDMIS.Execute("Select AccessoriesName from  SMIS_VACC ")
    Picture1.Enabled = False
    StoreMemVars





End Sub

Private Sub LABTRAN_Change()
    If LABTRAN = "BPO" Then
        FRA_PO.Caption = "BANK PO"
    ElseIf LABTRAN = "F" Then
        FRA_PO.Caption = "FINANCING"
    Else
        FRA_PO.Caption = "CASH"
    End If
End Sub

Private Sub lvFree_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then Exit Sub
    cboFreeBeeies.Enabled = False
    LABDETID = Row.Record(4).Value
    cmdDelFree.Enabled = True
    cboFreeBeeies.Text = Row.Record(0).Value
    txtFreeBeeiesAmount.Text = Row.Record(1).Value
    If IsNull(Row.Record(5).Value) = True Then

    Else
        If Row.Record(5).Value = False Then
            Check1.Value = 0
        Else
            Check1.Value = 1
        End If
    End If
    ShowHidePictureBox2 picAccessories, True, Picture1
End Sub

Private Sub Option1_Click()
    txtSearchCommission.SetFocus
End Sub

Private Sub Option2_Click()
    txtSearchCommission.SetFocus
End Sub

Private Sub Option3_Click()
    txtSearchCommission.SetFocus
End Sub

Private Sub Option4_Click()
    txtSearchCommission.Text = MonthName(Month(LOGDATE), True)
End Sub

Private Sub ReportControl1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorCode

    If KeyCode = 13 Then
        If ReportControl1.SelectedRows.Count > 0 Then
            RS.MoveFirst
            RS.Find ("ID=" & ReportControl1.SelectedRows(0).Record(6).Value)
            StoreMemVars
            ShowHidePictureBox2 Picture2, False
        End If
    ElseIf KeyCode = vbKeyUp Then
        If ReportControl1.SelectedRows(0).Index = 0 Then
            txtSearchCommission.SetFocus
        End If
    End If


    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub ReportControl1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    ReportControl1_KeyDown 13, 0
End Sub

Private Sub SearchMaster_NoSelectionMade()
    If RS.EOF Or RS.BOF Then
        Unload Me
    End If
End Sub

Private Sub SearchMaster_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    InitMemVars
    LoadVehicleDetail (Null2String(oCusRs!ignkey))
    LoadCustomerDetail (Null2String(oCusRs!VI_NO))
    LoadMrrDetail (Null2String(oCusRs!ignkey))
    txtDatePullOut_Change
    ComputeCostOfMoney

    Unload SearchMaster
    Picture1.Enabled = True
    cmdEdit.Value = True
    AddorEdit = "ADD"
End Sub

Private Sub txtCM_Days_Change()
    If AddorEdit = "EDIT" Then
        ComputeCostOfMoney
    End If
End Sub

Private Sub txtSearchCommission_Change()
    On Error GoTo ErrorCode:
    Dim SQL

    Dim FILTER                                                        As String
    If Option1.Value = True Then
        FILTER = " WHERE SMIS_COMMISSION.VI_NO LIKE '%" & Repleys(txtSearchCommission) & "%'"
    ElseIf Option2.Value = True Then
        FILTER = " WHERE SMIS_COMMISSION.IGNKEYNO LIKE '" & Repleys(txtSearchCommission) & "%'"
    ElseIf Option3.Value = True Then
        FILTER = " WHERE SMIS_SALESORDER.CUSTNAME LIKE '" & Repleys(txtSearchCommission) & "%'"
    ElseIf Option4.Value = True Then
        FILTER = " WHERE SMIS_SALESORDER.DATERELEASED LIKE '" & Repleys(txtSearchCommission) & "%'"
    End If

    SQL = " SELECT top 100 " _
        & " SMIS_SALESORDER.DATERELEASED ," _
        & " SMIS_COMMISSION.IGNKEYNO, " _
        & " SMIS_COMMISSION.VI_NO , " _
        & " SMIS_SALESORDER.CUSTNAME , " _
        & " SMIS_SALESORDER.MODELDESCRIPTION, " _
        & " SMIS_SALESORDER.COLOR, " _
        & " SMIS_COMMISSION.ID " _
        & " From SMIS_COMMISSION " _
        & " INNER JOIN SMIS_SALESORDER ON " _
        & " SMIS_COMMISSION.VI_NO = SMIS_SALESORDER.VI_NO  " & FILTER

    flex_FillReportView gconDMIS.Execute(SQL), ReportControl1
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub txtSearchCommission_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrorCode

    If KeyCode = vbKeyDown Then
        If ReportControl1.Records.Count > 0 Then
            ReportControl1.Rows(0).Selected = True
            ReportControl1.SetFocus
        End If
    End If


    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub txt_FREE_ADD_GotFocus()
    If NumericVal(txt_FREE_ADD.Text) <= 0 Then txt_FREE_ADD = ""
End Sub

Private Sub txt_FREE_ADD_LostFocus()
    If NumericVal(txt_FREE_ADD.Text) <= 0 Then txt_FREE_ADD = "0.00"
    txt_FREE_ADD = FormatNumber(txt_FREE_ADD)
End Sub

Private Sub txt_FREE_OTHER_GotFocus()
    If NumericVal(txt_FREE_OTHER.Text) <= 0 Then txt_FREE_OTHER = ""
End Sub

Private Sub txt_FREE_OTHER_LostFocus()
    If NumericVal(txt_FREE_OTHER.Text) <= 0 Then txt_FREE_OTHER = "0.00"
    txt_FREE_OTHER = FormatNumber(txt_FREE_OTHER)
End Sub

Private Sub txt_FREE_STD_GotFocus()
    If NumericVal(txt_FREE_STD.Text) <= 0 Then txt_FREE_STD = ""
End Sub

Private Sub txt_FREE_STD_LostFocus()
    If NumericVal(txt_FREE_STD.Text) <= 0 Then txt_FREE_STD = "0.00"
    txt_FREE_STD = FormatNumber(txt_FREE_STD)
End Sub

Private Sub txt_IncomeAcc_Change()
    txt_IncomeAccMargin = ToDoubleNumber(NumericVal(txt_IncomeAcc) - NumericVal(txt_IncomeAccCommission))
End Sub

Private Sub txt_IncomeAcc_GotFocus()
    If NumericVal(txt_IncomeAcc.Text) <= 0 Then txt_IncomeAcc = ""
End Sub

Private Sub txt_IncomeAcc_LostFocus()
    If NumericVal(txt_IncomeAcc.Text) <= 0 Then txt_IncomeAcc = "0.00"
    txt_IncomeAcc = FormatNumber(txt_IncomeAcc)
End Sub

Private Sub txt_IncomeAccCommission_Change()
    txt_IncomeAccMargin = ToDoubleNumber(NumericVal(txt_IncomeAcc) - NumericVal(txt_IncomeAccCommission))
End Sub

Private Sub txt_IncomeAccCommission_GotFocus()
    If NumericVal(txt_IncomeAccCommission.Text) <= 0 Then txt_IncomeAccCommission = ""
End Sub

Private Sub txt_IncomeAccCommission_LostFocus()
    If NumericVal(txt_IncomeAccCommission.Text) <= 0 Then txt_IncomeAccCommission = "0.00"
    txt_IncomeAccCommission = FormatNumber(txt_IncomeAccCommission)
End Sub

Private Sub txt_IncomeIns_Change()
    txt_IncomeInsMargin = ToDoubleNumber(NumericVal(txt_IncomeIns) - NumericVal(txt_IncomeInsCommission))
End Sub

Private Sub txt_IncomeIns_GotFocus()
    If NumericVal(txt_IncomeIns.Text) <= 0 Then txt_IncomeIns = ""
End Sub

Private Sub txt_IncomeIns_LostFocus()
    If NumericVal(txt_IncomeIns.Text) <= 0 Then txt_IncomeIns = "0.00"
    txt_IncomeIns = FormatNumber(txt_IncomeIns)
End Sub

Private Sub txt_IncomeInsCommission_Change()
    txt_IncomeInsMargin = ToDoubleNumber(NumericVal(txt_IncomeIns) - NumericVal(txt_IncomeInsCommission))
End Sub

Private Sub txt_IncomeInsCommission_GotFocus()
    If NumericVal(txt_IncomeInsCommission.Text) <= 0 Then txt_IncomeInsCommission = ""
End Sub

Private Sub txt_IncomeInsCommission_LostFocus()
    If NumericVal(txt_IncomeInsCommission.Text) <= 0 Then txt_IncomeInsCommission = "0.00"
    txt_IncomeInsCommission = FormatNumber(txt_IncomeInsCommission)
End Sub

Private Sub txt_IncomePO_Change()
    txt_IncomePOMargin = ToDoubleNumber(NumericVal(txt_IncomePO) - NumericVal(txt_IncomePOComission))
End Sub

Private Sub txt_IncomePO_GotFocus()
    If NumericVal(txt_IncomePO.Text) <= 0 Then txt_IncomePO = ""
End Sub

Private Sub txt_IncomePO_LostFocus()
    If NumericVal(txt_IncomePO.Text) <= 0 Then txt_IncomePO = "0.00"
    txt_IncomePO = FormatNumber(txt_IncomePO)
End Sub

Private Sub txt_IncomePOComission_Change()
    txt_IncomePOMargin = ToDoubleNumber(NumericVal(txt_IncomePO) - NumericVal(txt_IncomePOComission))
End Sub

Private Sub txt_IncomePOComission_GotFocus()
    If NumericVal(txt_IncomePOComission.Text) <= 0 Then txt_IncomePOComission = ""
End Sub

Private Sub txt_IncomePOComission_LostFocus()
    If NumericVal(txt_IncomePOComission.Text) <= 0 Then txt_IncomePOComission = "0.00"
    txt_IncomePOComission = FormatNumber(txt_IncomePOComission)
End Sub

Private Sub txt_IncomeReg_Change()
    txt_IncomeRegMargin = ToDoubleNumber(NumericVal(txt_IncomeReg) - NumericVal(txt_IncomeRegCommission))
End Sub

Private Sub txt_IncomeReg_GotFocus()
    If NumericVal(txt_IncomeReg.Text) <= 0 Then txt_IncomeReg = ""
End Sub

Private Sub txt_IncomeReg_LostFocus()
    If NumericVal(txt_IncomeReg.Text) <= 0 Then txt_IncomeReg = "0.00"
    txt_IncomeReg = FormatNumber(txt_IncomeReg)
End Sub

Private Sub txt_IncomeRegCommission_Change()
    txt_IncomeRegMargin = ToDoubleNumber(NumericVal(txt_IncomeReg) - NumericVal(txt_IncomeRegCommission))
End Sub

Private Sub txt_IncomeRegCommission_GotFocus()
    If NumericVal(txt_IncomeRegCommission.Text) <= 0 Then txt_IncomeRegCommission = ""
End Sub

Private Sub txt_IncomeRegCommission_LostFocus()
    If NumericVal(txt_IncomeRegCommission.Text) <= 0 Then txt_IncomeRegCommission = "0.00"
    txt_IncomeRegCommission = FormatNumber(txt_IncomeRegCommission)
End Sub

Private Sub txtFreeBeeiesAmount_GotFocus()
    If NumericVal(txtFreeBeeiesAmount.Text) <= 0 Then txtFreeBeeiesAmount = ""
End Sub

Private Sub txtFreeBeeiesAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFreeBeeiesAmount_LostFocus()
    If NumericVal(txtFreeBeeiesAmount.Text) <= 0 Then txtFreeBeeiesAmount = "0.00"
    txtFreeBeeiesAmount = FormatNumber(txtFreeBeeiesAmount)
End Sub

Private Sub txtCM_COSOFMONEY_GotFocus()
    If NumericVal(txtCM_COSOFMONEY.Text) <= 0 Then txtCM_COSOFMONEY = ""
End Sub

Private Sub txtCM_COSOFMONEY_LostFocus()
    If NumericVal(txtCM_COSOFMONEY.Text) <= 0 Then txtCM_COSOFMONEY = "0.00"
    txtCM_COSOFMONEY = FormatNumber(txtCM_COSOFMONEY)
End Sub

Private Sub txtCM_Days_GotFocus()
    If NumericVal(txtCM_Days.Text) <= 0 Then txtCM_Days = ""
End Sub

Private Sub txtCM_Days_LostFocus()
    If NumericVal(txtCM_Days.Text) <= 0 Then txtCM_Days = "0"
    txtCM_Days = FormatNumber(txtCM_Days)
End Sub

Private Sub txtDatePullOut_Change()
    If IsDate(txtDatePullOut) And IsDate(txtDateReleased) Then
        txtCM_Days = DateDiff("d", txtDatePullOut, txtDateReleased)
        If txtCM_Days < 4 Then: txtCM_Days = 0
    End If
End Sub

Private Sub txtDatePullOut_GotFocus()
    If IsDate(txtDatePullOut) = True Then
        txtDatePullOut = FormatDateTime(txtDatePullOut, vbShortDate)
    End If
End Sub

Private Sub txtDatePullOut_LostFocus()
    If IsDate(txtDatePullOut) And IsDate(txtDateReleased) = True Then
        txtDatePullOut = Format(txtDatePullOut, "mmm dd yyyy")
        If DateDiff("d", txtDatePullOut, txtDateReleased) < 0 Then
            MessagePop RecSaveError, "Invalid Entry", "Invalid Date", 500
            'txtDatePullOut.SetFocus

        End If
    End If
End Sub

Private Sub txtDateReleased_Change()
    If IsDate(txtDatePullOut) And IsDate(txtDateReleased) Then
        txtCM_Days = DateDiff("d", txtDatePullOut, txtDateReleased)
        If txtCM_Days < 4 Then: txtCM_Days = 0
    End If
End Sub

Private Sub txtDateReleased_GotFocus()
    If IsDate(txtDateReleased) = True Then
        txtDateReleased = FormatDateTime(txtDateReleased, vbShortDate)
    End If
End Sub

Private Sub txtDateReleased_LostFocus()
    If IsDate(txtDateReleased) And IsDate(txtDatePullOut) = True Then
        txtDateReleased = Format(txtDateReleased, "mmm dd yyyy")
        If DateDiff("d", txtDatePullOut, txtDateReleased) < 0 Then
            MessagePop RecSaveError, "Invalid Entry", "Invalid Date", 500
            txtDateReleased.SetFocus
        End If
    End If
End Sub

Private Sub txtDICommission_GotFocus()
    If NumericVal(txtDICommission.Text) <= 0 Then txtDICommission = ""
End Sub

Private Sub txtDICommission_LostFocus()
    If NumericVal(txtDICommission.Text) <= 0 Then txtDICommission = "0.00"
    txtDICommission = FormatNumber(txtDICommission)
End Sub

Private Sub txtUnitCommission_GotFocus()
    If NumericVal(txtUnitCommission.Text) <= 0 Then txtUnitCommission = ""
End Sub

Private Sub txtUnitCommission_LostFocus()
    If NumericVal(txtUnitCommission.Text) <= 0 Then txtUnitCommission = "0.00"
    txtUnitCommission = FormatNumber(txtUnitCommission)
End Sub


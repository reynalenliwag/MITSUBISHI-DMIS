VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CMIS Main Menu"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmCMISMainMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   9780
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6705
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   11715
      _Version        =   655364
      _ExtentX        =   20664
      _ExtentY        =   11827
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   4
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   18
      Item(0).Control(0)=   "cmdTrans_LTOFund"
      Item(0).Control(1)=   "cmdInquiry_ViewCashPosition(1)"
      Item(0).Control(2)=   "cmdTrans_CashierCashCount"
      Item(0).Control(3)=   "cmdProcess_CutOff"
      Item(0).Control(4)=   "cmdTrans_CheckEnchasment"
      Item(0).Control(5)=   "cmdTrans_BankDeposit"
      Item(0).Control(6)=   "cmdTrans_PettyCash"
      Item(0).Control(7)=   "cmdNONVAT"
      Item(0).Control(8)=   "cmdORVAT"
      Item(0).Control(9)=   "Label70"
      Item(0).Control(10)=   "Label31"
      Item(0).Control(11)=   "Label2"
      Item(0).Control(12)=   "Label9"
      Item(0).Control(13)=   "Label10"
      Item(0).Control(14)=   "Label12"
      Item(0).Control(15)=   "Label14"
      Item(0).Control(16)=   "Label15"
      Item(0).Control(17)=   "Label1"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   15
      Item(1).Control(0)=   "Command1"
      Item(1).Control(1)=   "cmdTables_OtherTransaction"
      Item(1).Control(2)=   "cmdTable_Transaction"
      Item(1).Control(3)=   "cmdTable_Bank"
      Item(1).Control(4)=   "cmdTable_CheckClassification"
      Item(1).Control(5)=   "cmdTable_PettyCashType"
      Item(1).Control(6)=   "cmdTables_Employee"
      Item(1).Control(7)=   "Label17"
      Item(1).Control(8)=   "cmdOtherTransaction"
      Item(1).Control(9)=   "Label18"
      Item(1).Control(10)=   "Label20"
      Item(1).Control(11)=   "Label21"
      Item(1).Control(12)=   "Label22"
      Item(1).Control(13)=   "Label29"
      Item(1).Control(14)=   "Picture2"
      Item(2).Caption =   "Reports"
      Item(2).ControlCount=   28
      Item(2).Control(0)=   "cmdCancelledOR"
      Item(2).Control(1)=   "cmdReport_PettyCashReplishment"
      Item(2).Control(2)=   "cmdReport_LTOSummary"
      Item(2).Control(3)=   "cmdReport_PettyCashSummary"
      Item(2).Control(4)=   "cmdReport_SalesDiscount"
      Item(2).Control(5)=   "cmdReport_OutPutTax"
      Item(2).Control(6)=   "cmdReport_CashOnHand"
      Item(2).Control(7)=   "cmdReport_CreditCardBankDeposit"
      Item(2).Control(8)=   "cmdInquiry_ViewCashPosition(0)"
      Item(2).Control(9)=   "cmdReport_CreditCardListingReport"
      Item(2).Control(10)=   "cmdCashTallyReport"
      Item(2).Control(11)=   "cmdReport_CustomerOverPayment"
      Item(2).Control(12)=   "cmdReport_InflowPerDateRange"
      Item(2).Control(13)=   "Label16"
      Item(2).Control(14)=   "Label47"
      Item(2).Control(15)=   "Label48"
      Item(2).Control(16)=   "Label43"
      Item(2).Control(17)=   "Label49"
      Item(2).Control(18)=   "Label51"
      Item(2).Control(19)=   "Label52"
      Item(2).Control(20)=   "Label55"
      Item(2).Control(21)=   "Label57"
      Item(2).Control(22)=   "Label58"
      Item(2).Control(23)=   "Label59"
      Item(2).Control(24)=   "Label61"
      Item(2).Control(25)=   "Label62"
      Item(2).Control(26)=   "Label63"
      Item(2).Control(27)=   "cmdReport_CorporateTax"
      Item(3).Caption =   "Other Setups"
      Item(3).ControlCount=   6
      Item(3).Control(0)=   "cmdReminders"
      Item(3).Control(1)=   "cmd(67)"
      Item(3).Control(2)=   "cmdSignatories"
      Item(3).Control(3)=   "Label30"
      Item(3).Control(4)=   "Label69"
      Item(3).Control(5)=   "Label41"
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4725
         Left            =   -65020
         ScaleHeight     =   4725
         ScaleWidth      =   3825
         TabIndex        =   70
         Top             =   1740
         Visible         =   0   'False
         Width           =   3825
         Begin VB.CommandButton cmdTables_ChartOfAcccounts 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   60
            MouseIcon       =   "frmCMISMainMenu.frx":0D92
            MousePointer    =   99  'Custom
            Picture         =   "frmCMISMainMenu.frx":0EE4
            Style           =   1  'Graphical
            TabIndex        =   75
            Tag             =   "1032"
            ToolTipText     =   "Chart Of Accounts"
            Top             =   915
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdTables_InsuranceCompany 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   60
            MouseIcon       =   "frmCMISMainMenu.frx":160C
            MousePointer    =   99  'Custom
            Picture         =   "frmCMISMainMenu.frx":175E
            Style           =   1  'Graphical
            TabIndex        =   74
            Tag             =   "1030"
            ToolTipText     =   "Insurance Company Listing"
            Top             =   30
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdTables_BankInformation 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   60
            MouseIcon       =   "frmCMISMainMenu.frx":1EC6
            MousePointer    =   99  'Custom
            Picture         =   "frmCMISMainMenu.frx":2018
            Style           =   1  'Graphical
            TabIndex        =   73
            Tag             =   "1045"
            ToolTipText     =   "Bank Information"
            Top             =   3540
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdTables_Particular 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   60
            MouseIcon       =   "frmCMISMainMenu.frx":27D9
            MousePointer    =   99  'Custom
            Picture         =   "frmCMISMainMenu.frx":292B
            Style           =   1  'Graphical
            TabIndex        =   72
            Tag             =   "1044"
            ToolTipText     =   "Particulars"
            Top             =   2655
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdTables_Payee 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   60
            MouseIcon       =   "frmCMISMainMenu.frx":2FC0
            MousePointer    =   99  'Custom
            Picture         =   "frmCMISMainMenu.frx":3112
            Style           =   1  'Graphical
            TabIndex        =   71
            Tag             =   "1043"
            ToolTipText     =   "Payee"
            Top             =   1785
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Information"
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
            Left            =   1050
            TabIndex        =   80
            Top             =   3840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Particulars"
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
            Left            =   1050
            TabIndex        =   79
            Top             =   2955
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Payee"
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
            Left            =   1050
            TabIndex        =   78
            Top             =   2085
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Chart Of Accounts"
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
            Left            =   1050
            TabIndex        =   77
            Top             =   1200
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance Company Listing"
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
            Left            =   1050
            TabIndex        =   76
            Top             =   285
            Visible         =   0   'False
            Width           =   2310
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   9180
         Top             =   6270
      End
      Begin VB.CommandButton cmdSignatories 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":3791
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":38E3
         Style           =   1  'Graphical
         TabIndex        =   64
         Tag             =   "1046"
         ToolTipText     =   "Signatories"
         Top             =   1800
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   67
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":3D44
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":3E96
         Style           =   1  'Graphical
         TabIndex        =   63
         Tag             =   "1088"
         ToolTipText     =   "Password Maintenance "
         Top             =   2700
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReminders 
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":47BA
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":490C
         Style           =   1  'Graphical
         TabIndex        =   62
         Tag             =   "1102"
         ToolTipText     =   "Reminders"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_CorporateTax 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -64960
         MouseIcon       =   "frmCMISMainMenu.frx":5187
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":52D9
         Style           =   1  'Graphical
         TabIndex        =   60
         Tag             =   "1078"
         ToolTipText     =   "Corporate Tax Report"
         Top             =   1740
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_InflowPerDateRange 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":59C3
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":5B15
         Style           =   1  'Graphical
         TabIndex        =   45
         Tag             =   "1065"
         ToolTipText     =   "Cash InFlow Per Date Of Range"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_CustomerOverPayment 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":623D
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":638F
         Style           =   1  'Graphical
         TabIndex        =   44
         Tag             =   "1066"
         ToolTipText     =   "Customer Over Payment Report"
         Top             =   1740
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdCashTallyReport 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":6B18
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":6C6A
         Style           =   1  'Graphical
         TabIndex        =   43
         Tag             =   "1067"
         ToolTipText     =   "Cash Tally Report"
         Top             =   2580
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_CreditCardListingReport 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":7345
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":7497
         Style           =   1  'Graphical
         TabIndex        =   42
         Tag             =   "1068"
         ToolTipText     =   "Credit Card Listing Report On Hand"
         Top             =   5055
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdInquiry_ViewCashPosition 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   0
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":7CE1
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":7E33
         Style           =   1  'Graphical
         TabIndex        =   41
         Tag             =   "1071"
         ToolTipText     =   "View Cash Position"
         Top             =   4245
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_CreditCardBankDeposit 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":84AA
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":85FC
         Style           =   1  'Graphical
         TabIndex        =   40
         Tag             =   "1069"
         ToolTipText     =   "Credit Card Bank Deposit Report"
         Top             =   5880
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_CashOnHand 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":8DF3
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":8F45
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "1076"
         ToolTipText     =   "Cash On Hand Report"
         Top             =   3420
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_OutPutTax 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -64960
         MouseIcon       =   "frmCMISMainMenu.frx":9685
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":97D7
         Style           =   1  'Graphical
         TabIndex        =   38
         Tag             =   "1077"
         ToolTipText     =   "Output Tax Report"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_SalesDiscount 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -64960
         MouseIcon       =   "frmCMISMainMenu.frx":9F16
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":A068
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "1079"
         ToolTipText     =   "Sales Discount Report"
         Top             =   2595
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_PettyCashSummary 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -64960
         MouseIcon       =   "frmCMISMainMenu.frx":A8BF
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":AA11
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "1081"
         ToolTipText     =   "Petty Cash Summary Report"
         Top             =   3420
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_LTOSummary 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -64960
         MouseIcon       =   "frmCMISMainMenu.frx":B2EC
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":B43E
         Style           =   1  'Graphical
         TabIndex        =   35
         Tag             =   "1082"
         ToolTipText     =   "L.T.O. Summary Report"
         Top             =   4245
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdReport_PettyCashReplishment 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -64960
         MouseIcon       =   "frmCMISMainMenu.frx":BCB9
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":BE0B
         Style           =   1  'Graphical
         TabIndex        =   34
         Tag             =   "1081"
         ToolTipText     =   "Petty Cash Replenishment Summary "
         Top             =   5895
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdCancelledOR 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -64960
         MouseIcon       =   "frmCMISMainMenu.frx":C589
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":C6DB
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "1048"
         ToolTipText     =   "NON VAT O.R."
         Top             =   5070
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTables_Employee 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":CE74
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":CFC6
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "1028"
         ToolTipText     =   "Employee"
         Top             =   4365
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTable_PettyCashType 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":D675
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":D7C7
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "1024"
         ToolTipText     =   "Petty Cash/L.T.O. Type"
         Top             =   2625
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTable_CheckClassification 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":DE46
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":DF98
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "1025"
         ToolTipText     =   "Check Classification"
         Top             =   3495
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTable_Bank 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":E618
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":E76A
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "1021"
         ToolTipText     =   "Bank"
         Top             =   1770
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTable_Transaction 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":EE4C
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":EF9E
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "1020"
         ToolTipText     =   "Transaction"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTables_OtherTransaction 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -64960
         MouseIcon       =   "frmCMISMainMenu.frx":F53B
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":F68D
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "1023"
         ToolTipText     =   "Other Transaction"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   -69340
         MouseIcon       =   "frmCMISMainMenu.frx":FC32
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":FD84
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "1069"
         ToolTipText     =   "Credit Card Bank Deposit Report"
         Top             =   5235
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdORVAT 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   660
         MouseIcon       =   "frmCMISMainMenu.frx":1057B
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":106CD
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "1047"
         ToolTipText     =   "O.R. with VAT"
         Top             =   900
         Width           =   720
      End
      Begin VB.CommandButton cmdNONVAT 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   660
         MouseIcon       =   "frmCMISMainMenu.frx":10E85
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":10FD7
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "1048"
         ToolTipText     =   "NON VAT O.R."
         Top             =   1770
         Width           =   720
      End
      Begin VB.CommandButton cmdTrans_PettyCash 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   660
         MouseIcon       =   "frmCMISMainMenu.frx":11770
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":118C2
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "1054"
         ToolTipText     =   "Petty Cash Entry"
         Top             =   2655
         Width           =   720
      End
      Begin VB.CommandButton cmdTrans_BankDeposit 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   660
         MouseIcon       =   "frmCMISMainMenu.frx":12019
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":1216B
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "1056"
         ToolTipText     =   "Bank Deposit"
         Top             =   4425
         Width           =   720
      End
      Begin VB.CommandButton cmdTrans_CheckEnchasment 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   660
         MouseIcon       =   "frmCMISMainMenu.frx":127B1
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":12903
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "1057"
         ToolTipText     =   "Check Encashment"
         Top             =   5310
         Width           =   720
      End
      Begin VB.CommandButton cmdProcess_CutOff 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5040
         MouseIcon       =   "frmCMISMainMenu.frx":12F67
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":130B9
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "1058"
         ToolTipText     =   "O.R. Cut-Off Entry"
         Top             =   1770
         Width           =   720
      End
      Begin VB.CommandButton cmdTrans_CashierCashCount 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5040
         MouseIcon       =   "frmCMISMainMenu.frx":13854
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":139A6
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "1059"
         ToolTipText     =   "Cashier Cash Count"
         Top             =   900
         Width           =   720
      End
      Begin VB.CommandButton cmdInquiry_ViewCashPosition 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   5025
         MouseIcon       =   "frmCMISMainMenu.frx":14059
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":141AB
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "1071"
         ToolTipText     =   "View Cash Position"
         Top             =   2655
         Width           =   720
      End
      Begin VB.CommandButton cmdTrans_LTOFund 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   660
         MouseIcon       =   "frmCMISMainMenu.frx":14822
         MousePointer    =   99  'Custom
         Picture         =   "frmCMISMainMenu.frx":14974
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "1054"
         ToolTipText     =   "Petty Cash Entry"
         Top             =   3540
         Width           =   720
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Signatories"
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
         Left            =   -68350
         TabIndex        =   67
         Top             =   2055
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Maintenance "
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
         Left            =   -68350
         TabIndex        =   66
         Top             =   2910
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reminders"
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
         Left            =   -68350
         TabIndex        =   65
         Top             =   1230
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O.R. with V.A.T."
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
         Left            =   1650
         TabIndex        =   61
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Replenishment Summary "
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
         Left            =   -64000
         TabIndex        =   59
         Top             =   6120
         Visible         =   0   'False
         Width           =   3150
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corporate Tax Report"
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
         Left            =   -64000
         TabIndex        =   58
         Top             =   1995
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.T.O. Summary Report"
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
         Left            =   -64000
         TabIndex        =   57
         Top             =   4530
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Summary Report"
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
         Left            =   -64000
         TabIndex        =   56
         Top             =   3660
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output Tax Report"
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
         Left            =   -64000
         TabIndex        =   55
         Top             =   1170
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Discount Report"
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
         Left            =   -64000
         TabIndex        =   54
         Top             =   2805
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash On Hand Report"
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
         Left            =   -68350
         TabIndex        =   53
         Top             =   3600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Cash Position"
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
         Left            =   -68350
         TabIndex        =   52
         Top             =   4500
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card Bank Deposit Report"
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
         Left            =   -68350
         TabIndex        =   51
         Top             =   6045
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card Listing Report On Hand"
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
         Left            =   -68350
         TabIndex        =   50
         Top             =   5325
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Tally Report"
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
         Left            =   -68350
         TabIndex        =   49
         Top             =   2835
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash InFlow Per Date Of Range"
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
         Left            =   -68350
         TabIndex        =   48
         Top             =   1185
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Over Payment Report"
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
         Left            =   -68350
         TabIndex        =   47
         Top             =   1995
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelled OR Report"
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
         Left            =   -64000
         TabIndex        =   46
         Top             =   5250
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         Left            =   -68350
         TabIndex        =   32
         Top             =   4620
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction"
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
         Left            =   -68350
         TabIndex        =   31
         Top             =   1215
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         Left            =   -68350
         TabIndex        =   30
         Top             =   2115
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Classification"
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
         Left            =   -68350
         TabIndex        =   29
         Top             =   3825
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Type"
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
         Left            =   -68350
         TabIndex        =   28
         Top             =   2880
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label cmdOtherTransaction 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Transaction"
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
         Left            =   -63970
         TabIndex        =   27
         Top             =   1215
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card Company"
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
         Left            =   -68350
         TabIndex        =   26
         Top             =   5520
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cashier Cash Count"
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
         Left            =   6030
         TabIndex        =   18
         Top             =   1185
         Width           =   1680
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process New Cut-Off Entry"
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
         Left            =   6030
         TabIndex        =   17
         Top             =   2085
         Width           =   2265
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Entry"
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
         Left            =   1650
         TabIndex        =   16
         Top             =   2925
         Width           =   1395
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Deposit"
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
         Left            =   1650
         TabIndex        =   15
         Top             =   4695
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Encashment"
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
         Left            =   1650
         TabIndex        =   14
         Top             =   5565
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NON VAT O.R."
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
         Left            =   1650
         TabIndex        =   13
         Top             =   2010
         Width           =   1155
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Cash Position"
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
         Left            =   6030
         TabIndex        =   12
         Top             =   2925
         Width           =   1635
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LTO FUND Entry"
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
         Left            =   1650
         TabIndex        =   11
         Top             =   3780
         Width           =   1320
      End
   End
   Begin VB.Label lblCutOff 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "September 07, 2010"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7110
      TabIndex        =   69
      Top             =   6720
      Width           =   2565
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Cut Off Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4590
      TabIndex        =   68
      Top             =   6720
      Width           =   2280
   End
   Begin VB.Label Label6 
      Caption         =   "FORCE CANCEL OF NON-VAT O.R."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   900
      TabIndex        =   1
      Top             =   4650
      Width           =   4965
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   6930
      Top             =   6690
      Width           =   2925
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   0
      Top             =   6690
      Width           =   6945
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'FUNCTION FEATURE   : added right acess settings
'DATE STARTED       : 10/15/2007
'LAST UPDATED       : 10/15/2007
'WHO UPDATED        : AXP
'UPDATING CODE      : AXP1015200716:59
'REQUEST NO         : AXP
Private Sub cmd_Click(Index As Integer)
    '''''''''''''''''''''''''''''''''''Files''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case cmd(Index).Tag
        Case FILES_TRANSACTION
            If Module_Access(LOGID, "FILES TRANSACTION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "A"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES TRANSACTION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "TRANSACTION"
            frmCMISSBookEntry.Caption = "TRANSACTION CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_BANK
            If Module_Access(LOGID, "FILES BANK", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "B"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES BANK"
            frmCMISSBookEntry.labCODE.Caption = "FILES CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "BANK NAME"
            frmCMISSBookEntry.Caption = "BANK CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_BRANCH
            If Module_Access(LOGID, "FILES BRANCH", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "C"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES BRANCH"
            frmCMISSBookEntry.labCODE.Caption = "FILES BRANCH"
            frmCMISSBookEntry.labDESCNAME.Caption = "BRANCH NAME"
            frmCMISSBookEntry.Caption = "BRANCH CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_OTHERTRANSACTION
            If Module_Access(LOGID, "FILES OTHER TRANSACTION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "D"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES OTHER TRANSACTION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "OTHER TRAN."
            frmCMISSBookEntry.Caption = "OTHER TRANSACTION CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_PETTYCASHLTOTYPE
            If Module_Access(LOGID, "FILES PETTY LTO TYPE", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "E"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES PETTY LTO TYPE"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "PETTY/LTO"
            frmCMISSBookEntry.Caption = "PETTY CASH/L.T.O. MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_CHECKCLASSIFICATION
            If Module_Access(LOGID, "FILES CHECK CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "F"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES CHECK CLASSIFICATION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "CHECK CLASS"
            frmCMISSBookEntry.Caption = "CHECK CODE CLASSIFICATION MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_PAYMENTCLASSIFICATION
            If Module_Access(LOGID, "FILES PAYMENT CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "F"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES PAYMENT CLASSIFICATION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "CHECK CLASS"
            frmCMISSBookEntry.Caption = "CHECK CODE CLASSIFICATION MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_DEPARTMENT
            If Module_Access(LOGID, "FILES DEPARTMENT", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "H"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES DEPARTMENT"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "DEPT. NAME"
            frmCMISSBookEntry.Caption = "DEPARMENT CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_EMPLOYEE
            If Module_Access(LOGID, "FILES EMPLOYEE", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "I"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES EMPLOYEE"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "EMP. NAME"
            frmCMISSBookEntry.Caption = "EMPLOYEES CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_REPLENISHMENTENTRY
            If Module_Access(LOGID, "FILES REPLENISHMENT ENTRY", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "J"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES REPLENISHMENT ENTRY"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "DESCRIPTION"
            frmCMISSBookEntry.Caption = "REPLENISHMENT CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_INSURANCECOMPANYLISTING
            If Module_Access(LOGID, "FILES INSURANCE COMPANY LISTING", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "K"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES INSURANCE COMPANY LISTING"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "COMPANY"
            frmCMISSBookEntry.Caption = "INSURANCE COMPANY CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_INTEROFFICECOLLECTION
            If Module_Access(LOGID, "FILES INTER OFFICE COLLECTION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "L"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES INTER OFFICE COLLECTION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "OFFICE"
            frmCMISSBookEntry.Caption = "INTER OFFICE CODE MAINTENANCE"
            frmCMISSBookEntry.Show

        Case FILES_CHARTSOFACCOUNT
            If Module_Access(LOGID, "FILES CHARTS OF ACCOUNT", "DATA ENTRY") = False Then Exit Sub

        Case FILES_CUSTOMERDEPOSIT
            If Module_Access(LOGID, "FILES CUSTOMER DEPOSIT", "DATA ENTRY") = False Then Exit Sub
        Case FILES_HIST_ORHISTORYFILES
            If Module_Access(LOGID, "FILES O.R. HISTORY FILES", "DATA ENTRY") = False Then Exit Sub
        Case FILES_HIST_BANKDEPOSITSHISTORYFILE
            If Module_Access(LOGID, "FILES BANK DEPOSITS HISTORY FILE", "DATA ENTRY") = False Then Exit Sub
        Case FILES_HIST_CASHENCASHMENTHISTORYFILE
            If Module_Access(LOGID, "FILES CASH ENCASHMENT HISTORY FILE", "DATA ENTRY") = False Then Exit Sub
        Case FILES_HIST_PETTYCASHHISTORYFILE
            If Module_Access(LOGID, "FILES PETTY CASH HISTORY FILE", "DATA ENTRY") = False Then Exit Sub
        Case FILES_HIST_LTOHISTORYFILE
            If Module_Access(LOGID, "FILES L.T.O. HISTORY FILE", "DATA ENTRY") = False Then Exit Sub
        Case FILES_PAIDAPP_REPAIRORDER
            If Module_Access(LOGID, "FILES PAID REPAIR ORDER", "DATA ENTRY") = False Then Exit Sub
        Case FILES_PAIDAPP_PERCUSTOMER
            If Module_Access(LOGID, "FILES PAID APP PER CUSTOMER", "DATA ENTRY") = False Then Exit Sub
        Case FILES_PAIDAPP_CASHANDCHARGEINVOICE
            If Module_Access(LOGID, "FILES PAIDAPP CASHANDCHARGEINVOICE", "DATA ENTRY") = False Then Exit Sub
        Case FILES_PAIDAPP_VEHICLESALESINVOICE
            If Module_Access(LOGID, "FILES PAIDAPP VEHICLE SALES INVOICE", "DATA ENTRY") = False Then Exit Sub
        Case FILES_PAYEE
            If Module_Access(LOGID, "FILES PAYEE", "DATA ENTRY") = False Then Exit Sub
            'frmCMISPayee.Show
        Case FILES_PARTICULARS
            If Module_Access(LOGID, "FILES PARTICULARS", "DATA ENTRY") = False Then Exit Sub
            'frmCMISParticulars.Show
        Case FILES_BANKINFORMATION
            If Module_Access(LOGID, "FILES BANKINFORMATION", "DATA ENTRY") = False Then Exit Sub
            'frmCMISBankInfo.Show
        Case FILES_SIGNATORIES
            If Module_Access(LOGID, "FILES SIGNATORIES", "DATA ENTRY") = False Then Exit Sub
            'frmCMISSignatories.Show
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''TRANSACTIONS''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case TRANS_ORE_ORWITHVAT
            If Module_Access(LOGID, "TRANSACTION O.R. WITH VAT", "TRANSACTION") = False Then Exit Sub
            OR_VAT_NONVAT = "VAT"
            frmCMISOREntry.LocalAcess = "TRANSACTION O.R. WITH VAT"
            frmCMISOREntry.Show
            frmCMISOREntry.Caption = "Official Receipt Data Entry [With VAT]"
        Case TOOL_ORWITHVAT
            If Module_Access(LOGID, "TRANSACTION O.R. WITHOUT VAT", "TRANSACTION") = False Then Exit Sub
            frmCMISOREntry.LocalAcess = "TRANSACTION O.R. WITHOUT VAT"
            OR_VAT_NONVAT = "NON-VAT"
            frmCMISOREntry.Show
            frmCMISOREntry.Caption = "Official Receipt Data Entry [NON VAT]"
        Case TRANS_ORE_NONVATOR, TOOL_NONVATOR
            If Module_Access(LOGID, "TRANSACTION O.R. WITH NON VAT", "TRANSACTION") = False Then Exit Sub
            frmCMISOREntry.LocalAcess = "TRANSACTION O.R. WITH NON VAT"
            OR_VAT_NONVAT = "NONVAT"
            frmCMISOREntry.Show
            frmCMISOREntry.Caption = "Official Receipt Data Entry [NON VAT]"
        Case TRANS_ORE_OLDOFFICIALRECEIPTS
            If Module_Access(LOGID, "TRANSACTION O.R. OLD OFFICIAL RECEIPTS", "TRANSACTION") = False Then Exit Sub
            frmCMISOREntry.LocalAcess = "TRANSACTION O.R. OLD OFFICIAL RECEIPTS"
            frmCMISOREntry.Show
            frmCMISOREntry.Caption = "Old Official Receipts"
        Case TRANS_FORCE_ORWITHVAT
            If Module_Access(LOGID, "TRANSACTION O.R. FORCE OR WITHVAT", "TRANSACTION") = False Then Exit Sub
            CANCEL_OR_VAT_NONVAT = "VAT"
            frmCMISCancelOREntry.Show
        Case TRANS_FORCE_NONVATOR
            If Module_Access(LOGID, "TRANSACTION O.R. FORCE OR NON WITHVAT", "TRANSACTION") = False Then Exit Sub
            CANCEL_OR_VAT_NONVAT = "NONVAT"
            frmCMISCancelOREntry.Show
        Case TRANS_FORCE_CARDWITHOR
            If Module_Access(LOGID, "TRANSACTION O.R. FORCE CARDWITHOR", "TRANSACTION") = False Then Exit Sub
            CANCEL_OR_VAT_NONVAT = "CARD_OR"
            frmCMISCancelOREntry.Show
        Case TRANS_FORCE_OLDOFFICIALRECEIPTS

            If Module_Access(LOGID, "TRANSACTION O.R. FORCE OLD OFFICIAL RECEIPTS", "TRANSACTION") = False Then Exit Sub
        Case TRANS_PETTYCASHENTRY, TOOL_PETTYCASHENTRY


        Case TRANS_LTOFUNDENTRY, TOOL_LTOFUNDENTRY

            If Module_Access(LOGID, "TRANSACTION LTO FUND ENTRY", "TRANSACTION") = False Then Exit Sub
            frmCMISLTOFUND.Show
        Case TRANS_BANKDEPOSIT, TOOL_BANKDEPOSIT

            If Module_Access(LOGID, "TRANSACTION BANKDEPOSIT", "TRANSACTION") = False Then Exit Sub
            frmCMISBankDeposit.Show

        Case TRANS_CHECKENCASHMENT, TOOL_CHECKENCASHMENT

            If Module_Access(LOGID, "TRANSACTION CHECK ENCASHMENT", "TRANSACTION") = False Then Exit Sub
            frmCMISCheckEncashment.Show
        Case TRANS_OFFICIALRECEIPTCUTOFFENTRY, TOOL_OFFICIALRECEIPTCUTOFFENTRY


            'frmCMISUpdateCUTOFFMaster.Show vbModal
        Case TRANS_CASHIERCASHCOUNT, TOOL_CASHIERCASHCOUNT
            If Module_Access(LOGID, "TRANSACTION CASHIER CASH COUNT", "TRANSACTION") = False Then Exit Sub
            frmCMISCashCount.Show
        Case TRANS_PETTYCASH

            If Module_Access(LOGID, "TRANSACTION PETTYCASH", "TRANSACTION") = False Then Exit Sub
            frmCMISPettyCash.Show
            ' frmCMISExpenses.Show
            'BG HERE
            '

            'frmCMISPettyCash.Show
        Case TRANS_ENCASHMENT

            If Module_Access(LOGID, "TRANSACTION ENCASHMENT", "TRANSACTION") = False Then Exit Sub
            ' frmCMISEncashment.Show
        Case TRANS_DEPOSITS

            If Module_Access(LOGID, "TRANSACTION DEPOSITS", "TRANSACTION") = False Then Exit Sub
            'frmCMISDeposits.Show
        Case TRANS_CASHCOUNT

            If Module_Access(LOGID, "TRANSACTION CASHCOUNT", "TRANSACTION") = False Then Exit Sub

            frmCMISCashCount.Show
        Case TRANS_ORSYSTEM

            If Module_Access(LOGID, "TRANSACTION ORSYSTEM", "TRANSACTION") = False Then Exit Sub
            'frmCMISORSystem.Show vbModal
            '''''''''''''''''''''''''''''''''''Maintainence''''''''''''''''''''''''''''''''''''''''''''''''
        Case MATAIN_SYSTEMCONFIGURATION

            If Module_Access(LOGID, "MAINTAIN SYSTEM CONFIGURATION", "SYSTEM") = False Then Exit Sub
            frmCMISProfile.Show
        Case MAINTAIN_COMPANYPROFILE

            If Module_Access(LOGID, "MAINTAIN COMPANYPROFILE", "SYSTEM") = False Then Exit Sub
            'frmCMISProfile.Show
        Case MAINTAIN_PASSWORDMAINTENANCE

            If Module_Access(LOGID, "MAINTIAN USER MAINTENANCE", "SYSTEM") = False Then Exit Sub
            frmAccMaintenance.Show

            'frmPass.Show

        Case MAINTAIN_ADV_EDITCASHPOSITION

            If Module_Access(LOGID, "MAINTIAN ADVANCED EDITCASHPOSITION", "SYSTEM") = False Then Exit Sub
            frmEDITViewCashPosition.Show

            '''''''''''''''''''''''''''''''''''Report''''''''''''''''''''''''''''''''''''''''''''''''
        Case REPORT_CASHINFLOWPERDATEOFRANGE, TOOL_CASHINFLOWPERDATEOFRANGE
            If Module_Access(LOGID, "REPORT CASH INFLOW PER DATE OF RANGE", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Cash In Flow Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show

        Case REPORT_CUSTOMEROVERPAYMENTREPORT, TOOL_CUSTOMEROVERPAYMENTREPORT
            If Module_Access(LOGID, "REPORT CUSTOMER OVER PAYMENT REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Customer Over-Payment Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show

        Case REPORT_CASHTALLYREPORT, TOOL_CASHTALLYREPORT

            If Module_Access(LOGID, "REPORT CASH TALLY REPORT", "REPORTS") = False Then Exit Sub
            frmCMISCutDate.Show
        Case REPORT_CRCARD_CARDLISTINGREPORTONHAND, TOOL_CREDITCARDPAYMENTREPORTS

            If Module_Access(LOGID, "REPORT CARD LISTING REPORT ON-HAND", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Card Listing Report On Hand"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_CRCARD_CARDBANKDEPOSITREPORT

            If Module_Access(LOGID, "REPORT CARD BANK DEPOSIT REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Credit Card Bank Deposit Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_DAILYTRANSMITTALREPORT, TOOL_DAILYTRANSMITTALREPORT

            If Module_Access(LOGID, "REPORT DAILY TRANSMITTAL REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Daily Transmittal Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_VIEWCASHPOSITION, TOOL_VIEWCASHPOSITION

            If Module_Access(LOGID, "REPORT VIEWCASHPOSITION", "REPORTS") = False Then Exit Sub
            frmViewCashPosition.Show vbModal
        Case REPORT_CASHREC_SUMMARYJOURNALREPORT

            If Module_Access(LOGID, "REPORT CASHRECIEPT SUMMARYJOURNALREPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Cash Receipts Journal Report - Summary"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_CASHREC_DETAILJOURNALREPORT

            If Module_Access(LOGID, "REPORT CASHRECIEPT DETAIL JOURNAL REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Cash Receipts Journal Report - Detail"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_CASHREC_APSUMMARYREPORT

            If Module_Access(LOGID, "REPORT CASHRECIEPT AP SUMMARY REPORT", "REPORTS") = False Then Exit Sub
        Case REPORT_PETTYCASHREPLENISHMENT
            If Module_Access(LOGID, "REPORT PETTY CASH REPLENISHMENT", "REPORTS") = False Then Exit Sub
        Case REPORT_CASHREC_CASHONHAND

            If Module_Access(LOGID, "REPORT CASHRECIEPT CASH ON HAND", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Cash On Hand Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_CASHREC_OUTPUTTAX

            If Module_Access(LOGID, "REPORT CASHRECIEPT OUT PUT TAX", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "OutPut Tax Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_CASHREC_CORPORATETAX

            If Module_Access(LOGID, "REPORT CASHRECIEPT CORPORATE TAX", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Corporate Tax Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_CASHREC_SALESDISCOUNT

            If Module_Access(LOGID, "REPORT CASHRECIEPT SALES DISCOUNT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Sales Discount Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_CASHREC_JOURNALWITHREFERENCECODE

            If Module_Access(LOGID, "REPORT CASHRECIEPT JOURNAL WITH REFERENCE CODE", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Journal with Reference Code"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_REPLENISH_LTOSUMMARYREPORT, TOOL_REPLENISHMENTSUMMARYREPORT

            If Module_Access(LOGID, "REPORT REPLENISH LTO SUMMARY REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "L.T.O. Summary Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_REPLENISH_PETTYCASHSUMMARYREPORT

            If Module_Access(LOGID, "REPORT REPLENISH PETTY CASH SUMMARY REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Petty Cash Summary Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        Case REPORT_SMOKETESTPAYMENTSSUMMARYREPORT

            If Module_Access(LOGID, "REPORT SMOKE TEST PAYMENTS SUMMARY REPORT", "REPORTS") = False Then Exit Sub

            '''''''''''''''''''''''''''''''''''Tools''''''''''''''''''''''''''''''''''''''''''''''''
        Case TOOL_SETDATETIME

            If Module_Access(LOGID, "TOOL SETDATETIME", "SYSTEM") = False Then Exit Sub
            'frmTOOLSSetDateTime.Show
            '''''''''''''''''''''''''''''''''''Windows''''''''''''''''''''''''''''''''''''''''''''''''
        Case WINDOW_ABOUT, TOOL_ABOUTTHEAUTHOR
            frmAbout.Show
        Case WINDOW_EXIT, TOOL_EXITSYSTEM
            Unload Me
        Case TOOL_DASHBOARD
            frmMainMenu.Show
    End Select
End Sub

Private Sub cmdCashTallyReport_Click()
    If Module_Access(LOGID, "REPORT CASH TALLY REPORT", "REPORTS") = False Then Exit Sub
    frmCMISCutDate.Show
End Sub

Private Sub cmdInquiry_ViewCashPosition_Click(Index As Integer)
    If Module_Access(LOGID, "VIEW CASH POSITION", "INQUIRY") = False Then Exit Sub

    If Module_Access(LOGID, "TRANSACTION O.R. WITH VAT", "TRANSACTION") = False Then
        'frmViewPettyCashPosition.Show
        frmViewCashPosition.Show vbModal
    Else
        
        frmViewCashPosition.Show vbModal
        'frmViewCollectionPosition.Show
    End If
End Sub

Private Sub cmdSignatories_Click()
    If Module_Access(LOGID, "SYSTEM CONFIGURATION", "SYSTEM") = False Then Exit Sub
    frmCMISProfile.Show
End Sub

Private Sub cmdTrans_BankDeposit_Click()
    If Module_Access(LOGID, "TRANSACTION BANKDEPOSIT", "TRANSACTION") = False Then Exit Sub
    frmCMISBankDeposit.Show
End Sub

Private Sub cmdTrans_CashierCashCount_Click()
    If Module_Access(LOGID, "TRANSACTION CASHIER CASH COUNT", "TRANSACTION") = False Then Exit Sub
    frmCMISCashCount.Show
End Sub

Private Sub cmdTrans_CheckEnchasment_Click()
    If Module_Access(LOGID, "TRANSACTION CHECK ENCASHMENT", "TRANSACTION") = False Then Exit Sub
    frmCMISCheckEncashment.Show
End Sub

Private Sub cmdTrans_LTOFund_Click()
    If Module_Access(LOGID, "TRANSACTION LTO FUND ENTRY", "TRANSACTION") = False Then Exit Sub
    frmCMISLTOFUND.Show
End Sub

Private Sub cmdNONVAT_Click()
    If Module_Access(LOGID, "TRANSACTION O.R. WITHOUT VAT", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    OR_VAT_NONVAT = "NON-VAT"
    Unload frmCMISOREntry
    frmCMISOREntry.LocalAcess = "TRANSACTION O.R. WITHOUT VAT"
    frmCMISOREntry.Show
    frmCMISOREntry.Caption = "Official Receipt Data Entry [NON VAT]"
End Sub

Private Sub cmdORVAT_Click()
    If Module_Access(LOGID, "TRANSACTION O.R. WITH VAT", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    OR_VAT_NONVAT = "VAT"
    Unload frmCMISOREntry
    frmCMISOREntry.LocalAcess = "TRANSACTION O.R. WITH VAT"
    frmCMISOREntry.Show
    frmCMISOREntry.Caption = "Official Receipt Data Entry [With VAT]"
End Sub

Private Sub cmdProcess_CutOff_Click()
    If Module_Access(LOGID, "OFFICIAL RECEIPT CUT-OFF ENTRY", "PROCESSING") = False Then Exit Sub
    frmCMISProcessCUTOFF.Show vbModal
End Sub

Private Sub cmdReport_CashOnHand_Click()
    If Module_Access(LOGID, "REPORT CASHRECIEPT CASH ON HAND", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "Cash On Hand Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_CorporateTax_Click()
    If Module_Access(LOGID, "REPORT CASHRECIEPT CORPORATE TAX", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "Corporate Tax Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_CreditCardBankDeposit_Click()
    If Module_Access(LOGID, "REPORT CARD BANK DEPOSIT REPORT", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "Credit Card Bank Deposit Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_CreditCardListingReport_Click()
    If Module_Access(LOGID, "REPORT CARD LISTING REPORT ON-HAND", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "Card Listing Report On Hand"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_CustomerOverPayment_Click()
    If Module_Access(LOGID, "REPORT CUSTOMER OVER PAYMENT REPORT", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "Customer Over-Payment Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_InflowPerDateRange_Click()
    If Module_Access(LOGID, "REPORT CASH INFLOW PER DATE OF RANGE", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "Cash In Flow Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_LTOSummary_Click()
    If Module_Access(LOGID, "REPORT REPLENISH LTO SUMMARY REPORT", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "L.T.O. Summary Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_OutPutTax_Click()
    If Module_Access(LOGID, "REPORT CASHRECIEPT OUT PUT TAX", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "OutPut Tax Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_PettyCashReplishment_Click()
    If Module_Access(LOGID, "REPORT PETTY CASH REPLENISHMENT", "REPORTS") = False Then Exit Sub
End Sub

Private Sub cmdReport_PettyCashSummary_Click()
    If Module_Access(LOGID, "REPORT REPLENISH PETTY CASH SUMMARY REPORT", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "Petty Cash Summary Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdReport_SalesDiscount_Click()
    If Module_Access(LOGID, "REPORT CASHRECIEPT SALES DISCOUNT", "REPORTS") = False Then Exit Sub
    CMIS_Report_Range = "Sales Discount Report"
    Unload frmCMISReportRange
    frmCMISReportRange.Show
End Sub

Private Sub cmdTable_Bank_Click()
    If Module_Access(LOGID, "FILES BANK", "DATA ENTRY") = False Then Exit Sub
    BOOKTYPE = "B"
    On Error Resume Next
    Unload frmCMISSBookEntry
    frmCMISSBookEntry.LocalAcess = "FILES BANK"
    frmCMISSBookEntry.labCODE.Caption = "FILES CODE"
    frmCMISSBookEntry.labDESCNAME.Caption = "BANK NAME"
    frmCMISSBookEntry.Caption = "BANK CODE MAINTENANCE"
    frmCMISSBookEntry.Show
End Sub

Private Sub cmdTable_CheckClassification_Click()
    If Module_Access(LOGID, "FILES CHECK CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
    BOOKTYPE = "F"
    On Error Resume Next
    Unload frmCMISSBookEntry
    frmCMISSBookEntry.LocalAcess = "FILES CHECK CLASSIFICATION"
    frmCMISSBookEntry.labCODE.Caption = "CODE"
    frmCMISSBookEntry.labDESCNAME.Caption = "CHECK CLASS"
    frmCMISSBookEntry.Caption = "CHECK CODE CLASSIFICATION MAINTENANCE"
    frmCMISSBookEntry.Show
End Sub

Private Sub cmdTable_PettyCashType_Click()
    If Module_Access(LOGID, "FILES PETTY LTO TYPE", "DATA ENTRY") = False Then Exit Sub
    BOOKTYPE = "E"
    On Error Resume Next
    Unload frmCMISSBookEntry
    frmCMISSBookEntry.LocalAcess = "FILES PETTY LTO TYPE"
    frmCMISSBookEntry.labCODE.Caption = "CODE"
    frmCMISSBookEntry.labDESCNAME.Caption = "PETTY/LTO"
    frmCMISSBookEntry.Caption = "PETTY CASH/L.T.O. MAINTENANCE"
    frmCMISSBookEntry.Show
End Sub

Private Sub cmdTable_Transaction_Click()
    If Module_Access(LOGID, "FILES TRANSACTION", "DATA ENTRY") = False Then Exit Sub
    BOOKTYPE = "A"
    On Error Resume Next
    Unload frmCMISSBookEntry
    frmCMISSBookEntry.LocalAcess = "FILES TRANSACTION"
    frmCMISSBookEntry.labCODE.Caption = "CODE"
    frmCMISSBookEntry.labDESCNAME.Caption = "TRANSACTION"
    frmCMISSBookEntry.Caption = "TRANSACTION CODE MAINTENANCE"
    frmCMISSBookEntry.Show
End Sub

Private Sub cmdTables_Employee_Click()
    If Module_Access(LOGID, "FILES EMPLOYEE", "DATA ENTRY") = False Then Exit Sub
    BOOKTYPE = "I"
    On Error Resume Next
    Unload frmCMISSBookEntry
    frmCMISSBookEntry.LocalAcess = "FILES EMPLOYEE"
    frmCMISSBookEntry.labCODE.Caption = "CODE"
    frmCMISSBookEntry.labDESCNAME.Caption = "EMP. NAME"
    frmCMISSBookEntry.Caption = "EMPLOYEES CODE MAINTENANCE"
    frmCMISSBookEntry.Show
End Sub

Private Sub cmdTables_InsuranceCompany_Click()
    If Module_Access(LOGID, "FILES INSURANCE COMPANY LISTING", "DATA ENTRY") = False Then Exit Sub
    BOOKTYPE = "K"
    On Error Resume Next
    Unload frmCMISSBookEntry
    frmCMISSBookEntry.LocalAcess = "FILES INSURANCE COMPANY LISTING"
    frmCMISSBookEntry.labCODE.Caption = "CODE"
    frmCMISSBookEntry.labDESCNAME.Caption = "COMPANY"
    frmCMISSBookEntry.Caption = "INSURANCE COMPANY CODE MAINTENANCE"
    frmCMISSBookEntry.Show
End Sub

Private Sub cmdTables_OtherTransaction_Click()
    If Module_Access(LOGID, "FILES OTHER TRANSACTION", "DATA ENTRY") = False Then Exit Sub
    BOOKTYPE = "D"
    On Error Resume Next
    Unload frmCMISSBookEntry
    frmCMISSBookEntry.LocalAcess = "FILES OTHER TRANSACTION"
    frmCMISSBookEntry.labCODE.Caption = "CODE"
    frmCMISSBookEntry.labDESCNAME.Caption = "OTHER TRAN."
    frmCMISSBookEntry.Caption = "OTHER TRANSACTION CODE MAINTENANCE"
    frmCMISSBookEntry.Show
End Sub

Private Sub cmdTrans_PettyCash_Click()
    If Module_Access(LOGID, "TRANSACTION PETTY CASH ENTRY", "TRANSACTION") = False Then Exit Sub
    frmCMISPettyCash.Show
End Sub

Private Sub cmdCancelledOR_Click()
    If Module_Access(LOGID, "CANCEL OR REPORT", "REPORTS") = False Then Exit Sub
    frmReportCancelOR.Show
End Sub

Private Sub cmdReminders_Click()
    'Upating Code       : AXP-062620071225
    frmSMIS_Log_Reminder.Show
End Sub

Private Sub Command1_Click()
    frmCreditCardCompany.Show
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    TabControl1.SelectedItem = 0
End Sub

Private Sub Timer1_Timer()
    lblCutOff = Format(CURRENT_CUTOFF_DATE, "mmmm dd, yyyy")
End Sub

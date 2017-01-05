VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CSMS Main Menu"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   2070
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmCSMSMainMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9090
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9075
      _Version        =   655364
      _ExtentX        =   16007
      _ExtentY        =   10186
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   5
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "cmdMain_BillingSystem"
      Item(0).Control(1)=   "cmdMain_JobEstimate"
      Item(0).Control(2)=   "cmdMain_ServiceCounter"
      Item(0).Control(3)=   "cmdMain_JobClock"
      Item(0).Control(4)=   "Label37"
      Item(0).Control(5)=   "Label2"
      Item(0).Control(6)=   "Label1(0)"
      Item(0).Control(7)=   "Label11"
      Item(0).Control(8)=   "TabControl3"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   26
      Item(1).Control(0)=   "cmdTab_Color"
      Item(1).Control(1)=   "cmdMake"
      Item(1).Control(2)=   "Command5"
      Item(1).Control(3)=   "cmdMaster_Model"
      Item(1).Control(4)=   "cmdMaster_Customer"
      Item(1).Control(5)=   "Label58"
      Item(1).Control(6)=   "Label54(0)"
      Item(1).Control(7)=   "Label50(0)"
      Item(1).Control(8)=   "Label4"
      Item(1).Control(9)=   "Label22"
      Item(1).Control(10)=   "cmdTable_DealerMaster"
      Item(1).Control(11)=   "cmdMaster_OtherJobs"
      Item(1).Control(12)=   "cmdServiceMaintenace"
      Item(1).Control(13)=   "cmdMaster_CannedLabor"
      Item(1).Control(14)=   "cmdMaster_LTS"
      Item(1).Control(15)=   "cmdMaster_PMS"
      Item(1).Control(16)=   "Label54(1)"
      Item(1).Control(17)=   "Label3"
      Item(1).Control(18)=   "Label50(1)"
      Item(1).Control(19)=   "Label5"
      Item(1).Control(20)=   "Label25"
      Item(1).Control(21)=   "Label26"
      Item(1).Control(22)=   "Command11"
      Item(1).Control(23)=   "Label8(1)"
      Item(1).Control(24)=   "cmdServiceInternal"
      Item(1).Control(25)=   "Label51"
      Item(2).Caption =   "Inquiry"
      Item(2).ControlCount=   16
      Item(2).Control(0)=   "cmdInq_CusVeh"
      Item(2).Control(1)=   "cmdInq_ServiceAdvisorWorkDetail"
      Item(2).Control(2)=   "Label10"
      Item(2).Control(3)=   "Label9"
      Item(2).Control(4)=   "cmdInq_Acc"
      Item(2).Control(5)=   "cmdInq_Accessoires"
      Item(2).Control(6)=   "cmdInq_Parts"
      Item(2).Control(7)=   "cmdInq_Techmonitoring"
      Item(2).Control(8)=   "cmdInq_JobEsitmateListing"
      Item(2).Control(9)=   "Label27(2)"
      Item(2).Control(10)=   "Label27(0)"
      Item(2).Control(11)=   "Label28"
      Item(2).Control(12)=   "Label35"
      Item(2).Control(13)=   "Label23"
      Item(2).Control(14)=   "cmdInq_ProspInq"
      Item(2).Control(15)=   "Label27(1)"
      Item(3).Caption =   "Reports"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControl2"
      Item(4).Caption =   "Other Setups"
      Item(4).ControlCount=   16
      Item(4).Control(0)=   "Command10"
      Item(4).Control(1)=   "cmdMain_Complaint"
      Item(4).Control(2)=   "cmdMain_ConcernResolution"
      Item(4).Control(3)=   "cmdAuditInquiry"
      Item(4).Control(4)=   "cmdCompany_Profile(14)"
      Item(4).Control(5)=   "cmdPassword"
      Item(4).Control(6)=   "Label60"
      Item(4).Control(7)=   "Label47"
      Item(4).Control(8)=   "Label44"
      Item(4).Control(9)=   "Label57"
      Item(4).Control(10)=   "label"
      Item(4).Control(11)=   "Label69"
      Item(4).Control(12)=   "Command15"
      Item(4).Control(13)=   "Label7"
      Item(4).Control(14)=   "Label18"
      Item(4).Control(15)=   "cmdctr"
      Begin VB.CommandButton cmdctr 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -65260
         MouseIcon       =   "frmCSMSMainMenu.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   140
         Tag             =   "1102"
         ToolTipText     =   "Write a Reminder"
         Top             =   1650
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -65260
         MouseIcon       =   "frmCSMSMainMenu.frx":2256
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":23A8
         Style           =   1  'Graphical
         TabIndex        =   136
         Tag             =   "1102"
         ToolTipText     =   "Write a Reminder"
         Top             =   780
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdServiceInternal 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -63820
         MouseIcon       =   "frmCSMSMainMenu.frx":342A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":357C
         Style           =   1  'Graphical
         TabIndex        =   132
         Tag             =   "1318"
         ToolTipText     =   "view Internal Service payment"
         Top             =   2655
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdMain_BillingSystem 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   330
         MouseIcon       =   "frmCSMSMainMenu.frx":45FE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":4750
         Style           =   1  'Graphical
         TabIndex        =   129
         Tag             =   "1033"
         ToolTipText     =   "View Billing System"
         Top             =   2830
         Width           =   795
      End
      Begin VB.CommandButton cmdInq_JobEsitmateListing 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -60700
         MouseIcon       =   "frmCSMSMainMenu.frx":57D2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":5924
         Style           =   1  'Graphical
         TabIndex        =   122
         Tag             =   "1031"
         ToolTipText     =   "View Job Estimate"
         Top             =   5250
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -63820
         MouseIcon       =   "frmCSMSMainMenu.frx":69A6
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":6AF8
         Style           =   1  'Graphical
         TabIndex        =   112
         Tag             =   "1020"
         ToolTipText     =   "View Customer Master List"
         Top             =   1725
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdPassword 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":7B7A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":7CCC
         Style           =   1  'Graphical
         TabIndex        =   105
         Tag             =   "1407"
         ToolTipText     =   "View Password Maintenance "
         Top             =   810
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdCompany_Profile 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":8D4E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":8EA0
         Style           =   1  'Graphical
         TabIndex        =   104
         Tag             =   "1055"
         ToolTipText     =   "View Company Profile"
         Top             =   1650
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAuditInquiry 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":9F22
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":A074
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "View Audit Inquiry"
         Top             =   2460
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdMain_ConcernResolution 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":B0F6
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":B248
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "View Concern Resolution"
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdMain_Complaint 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":C2CA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":C41C
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Write a complain"
         Top             =   3270
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":D49E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":D5F0
         Style           =   1  'Graphical
         TabIndex        =   100
         Tag             =   "1102"
         ToolTipText     =   "Write a Reminder"
         Top             =   4890
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdInq_ProspInq 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -60070
         MouseIcon       =   "frmCSMSMainMenu.frx":E672
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":E7C4
         Style           =   1  'Graphical
         TabIndex        =   98
         Tag             =   "1196"
         ToolTipText     =   "Prospect Inquiry"
         Top             =   4260
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdInq_Techmonitoring 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -65530
         MouseIcon       =   "frmCSMSMainMenu.frx":F07B
         Picture         =   "frmCSMSMainMenu.frx":100FD
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Technician Attendance Monitoring"
         Top             =   840
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdInq_Parts 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":1117F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":112D1
         Style           =   1  'Graphical
         TabIndex        =   91
         Tag             =   "1048"
         ToolTipText     =   "View Parts Inquiry"
         Top             =   2610
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdInq_Accessoires 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":12353
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":124A5
         Style           =   1  'Graphical
         TabIndex        =   90
         Tag             =   "1049"
         ToolTipText     =   "View Materials inquiry"
         Top             =   3495
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdInq_Acc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":13527
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":13679
         Style           =   1  'Graphical
         TabIndex        =   89
         Tag             =   "1295"
         ToolTipText     =   "view Accessories inquiry"
         Top             =   4380
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdInq_ServiceAdvisorWorkDetail 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":146FB
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":1484D
         Style           =   1  'Graphical
         TabIndex        =   86
         Tag             =   "1043"
         ToolTipText     =   "View Repair Order inquiry"
         Top             =   1725
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdInq_CusVeh 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":158CF
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":15A21
         Style           =   1  'Graphical
         TabIndex        =   85
         Tag             =   "1042"
         ToolTipText     =   "View Customer Vehicle Inquiry"
         Top             =   840
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdMaster_PMS 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -66790
         MouseIcon       =   "frmCSMSMainMenu.frx":16AA3
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":16BF5
         Style           =   1  'Graphical
         TabIndex        =   78
         Tag             =   "1024"
         ToolTipText     =   "View PMS Jobs"
         Top             =   2655
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdMaster_LTS 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -66790
         MouseIcon       =   "frmCSMSMainMenu.frx":17C77
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":17DC9
         Style           =   1  'Graphical
         TabIndex        =   77
         Tag             =   "1023"
         ToolTipText     =   "View General Jobs"
         Top             =   1725
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdMaster_CannedLabor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -66790
         MouseIcon       =   "frmCSMSMainMenu.frx":18E4B
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":18F9D
         Style           =   1  'Graphical
         TabIndex        =   76
         Tag             =   "1025"
         ToolTipText     =   "View Canned Labor"
         Top             =   3600
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdServiceMaintenace 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -63820
         MouseIcon       =   "frmCSMSMainMenu.frx":1A01F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":1A171
         Style           =   1  'Graphical
         TabIndex        =   75
         Tag             =   "1020"
         ToolTipText     =   "View Service Personnel Maintenance"
         Top             =   780
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdMaster_OtherJobs 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -66790
         MouseIcon       =   "frmCSMSMainMenu.frx":1B1F3
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":1B345
         Style           =   1  'Graphical
         TabIndex        =   74
         Tag             =   "1023"
         ToolTipText     =   "View Other Jobs"
         Top             =   780
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdTable_DealerMaster 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -66790
         MouseIcon       =   "frmCSMSMainMenu.frx":1C3C7
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":1C519
         Style           =   1  'Graphical
         TabIndex        =   73
         Tag             =   "1289"
         ToolTipText     =   "Dealer Master File"
         Top             =   4560
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdMaster_Customer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":1CCA2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":1CDF4
         Style           =   1  'Graphical
         TabIndex        =   67
         Tag             =   "1020"
         ToolTipText     =   "View Customer Master List"
         Top             =   780
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdMaster_Model 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":1DE76
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":1DFC8
         Style           =   1  'Graphical
         TabIndex        =   66
         Tag             =   "1026"
         ToolTipText     =   "View Vehicle Model"
         Top             =   2655
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":1F04A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":1F19C
         Style           =   1  'Graphical
         TabIndex        =   65
         Tag             =   "1407"
         ToolTipText     =   "Customer Vehicle Information Maintenance"
         Top             =   1725
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdMake 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":2021E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":20370
         Style           =   1  'Graphical
         TabIndex        =   64
         Tag             =   "1206"
         ToolTipText     =   "Vehicle Class"
         Top             =   4560
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdTab_Color 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -69670
         MouseIcon       =   "frmCSMSMainMenu.frx":208DD
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":20A2F
         Style           =   1  'Graphical
         TabIndex        =   63
         Tag             =   "1084"
         ToolTipText     =   "Vehicle Color"
         Top             =   3600
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdMain_JobClock 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   330
         MouseIcon       =   "frmCSMSMainMenu.frx":22B01
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":22C53
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "View Time Clock/Job Clock Log-In"
         Top             =   3695
         Width           =   795
      End
      Begin VB.CommandButton cmdMain_ServiceCounter 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   330
         MouseIcon       =   "frmCSMSMainMenu.frx":23CD5
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":23E27
         Style           =   1  'Graphical
         TabIndex        =   46
         Tag             =   "1030"
         ToolTipText     =   "View Service Counter"
         Top             =   1140
         Width           =   795
      End
      Begin VB.CommandButton cmdMain_JobEstimate 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   330
         MouseIcon       =   "frmCSMSMainMenu.frx":25EF9
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSMainMenu.frx":2604B
         Style           =   1  'Graphical
         TabIndex        =   45
         Tag             =   "1031"
         ToolTipText     =   "View Job Estimate"
         Top             =   2000
         Width           =   795
      End
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   5175
         Left            =   -69970
         TabIndex        =   1
         Top             =   570
         Visible         =   0   'False
         Width           =   9015
         _Version        =   655364
         _ExtentX        =   15901
         _ExtentY        =   9128
         _StockProps     =   64
         Appearance      =   2
         Color           =   4
         PaintManager.Layout=   2
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         PaintManager.FixedTabWidth=   160
         PaintManager.MinTabWidth=   100
         ItemCount       =   3
         Item(0).Caption =   "Yearly/Monthly/Weekly"
         Item(0).ControlCount=   19
         Item(0).Control(0)=   "Label41"
         Item(0).Control(1)=   "Label39"
         Item(0).Control(2)=   "Label15"
         Item(0).Control(3)=   "Label43"
         Item(0).Control(4)=   "Label34"
         Item(0).Control(5)=   "Label42"
         Item(0).Control(6)=   "Label24"
         Item(0).Control(7)=   "cmdReport_TransactionForFollowup"
         Item(0).Control(8)=   "cmdReport_WorkshopSalesWeeklyPerformance"
         Item(0).Control(9)=   "cmdReport_AfterSales"
         Item(0).Control(10)=   "cmdReport_ActualManning"
         Item(0).Control(11)=   "cmdReport_HyundaiDealerMonthlyPerformance"
         Item(0).Control(12)=   "cmdReport_UnitsReceivedWeeklyPerformance"
         Item(0).Control(13)=   "cmdReport_WorkInProgress"
         Item(0).Control(14)=   "Command6"
         Item(0).Control(15)=   "Label17(2)"
         Item(0).Control(16)=   "Label17(0)"
         Item(0).Control(17)=   "Label17(1000)"
         Item(0).Control(18)=   "Label17(1)"
         Item(1).Caption =   "Service"
         Item(1).ControlCount=   16
         Item(1).Control(0)=   "Label16"
         Item(1).Control(1)=   "Label32(0)"
         Item(1).Control(2)=   "Label30"
         Item(1).Control(3)=   "Label33"
         Item(1).Control(4)=   "cmdReport_ServiceReport"
         Item(1).Control(5)=   "cmdReport_ServiceAdvisor"
         Item(1).Control(6)=   "cmdReport_VehicleAgingReport"
         Item(1).Control(7)=   "cmdReport_AppointmentDiary"
         Item(1).Control(8)=   "cmdReport_ServiceAdivsorSales"
         Item(1).Control(9)=   "Label32(1)"
         Item(1).Control(10)=   "Command14"
         Item(1).Control(11)=   "Label6"
         Item(1).Control(12)=   "Command17"
         Item(1).Control(13)=   "Label12"
         Item(1).Control(14)=   "Command19"
         Item(1).Control(15)=   "Label1(2)"
         Item(2).Caption =   "Technician/Other"
         Item(2).ControlCount=   24
         Item(2).Control(0)=   "Label36"
         Item(2).Control(1)=   "Label13(0)"
         Item(2).Control(2)=   "cmdReport_VehicleByModel"
         Item(2).Control(3)=   "cmdReport_CustDir"
         Item(2).Control(4)=   "cmdReport_Technician"
         Item(2).Control(5)=   "Command2"
         Item(2).Control(6)=   "Label45"
         Item(2).Control(7)=   "Command9"
         Item(2).Control(8)=   "Label59"
         Item(2).Control(9)=   "Label14"
         Item(2).Control(10)=   "cmdEstimatedPMS"
         Item(2).Control(11)=   "cmdSummary"
         Item(2).Control(12)=   "Label13(1)"
         Item(2).Control(13)=   "Label13(2)"
         Item(2).Control(14)=   "Command1"
         Item(2).Control(15)=   "Label17(3)"
         Item(2).Control(16)=   "Command4"
         Item(2).Control(17)=   "Label17(4)"
         Item(2).Control(18)=   "Command12"
         Item(2).Control(19)=   "Label17(5)"
         Item(2).Control(20)=   "Command13"
         Item(2).Control(21)=   "Label17(6)"
         Item(2).Control(22)=   "Command16"
         Item(2).Control(23)=   "Label17(7)"
         Begin VB.CommandButton Command19 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65230
            MouseIcon       =   "frmCSMSMainMenu.frx":270CD
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":2721F
            Style           =   1  'Graphical
            TabIndex        =   142
            ToolTipText     =   "Generate Void R.O. Report"
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command17 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65230
            MouseIcon       =   "frmCSMSMainMenu.frx":282A1
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":283F3
            Style           =   1  'Graphical
            TabIndex        =   138
            ToolTipText     =   "Generate MPR Schedules"
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65230
            MouseIcon       =   "frmCSMSMainMenu.frx":29475
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":295C7
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Generate MPR Schedules"
            Top             =   690
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command16 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65290
            MouseIcon       =   "frmCSMSMainMenu.frx":2A649
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":2A79B
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Generate MPR Schedules"
            Top             =   4080
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65290
            MouseIcon       =   "frmCSMSMainMenu.frx":2B81D
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":2B96F
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Generate MPR Schedules"
            Top             =   3240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65290
            MouseIcon       =   "frmCSMSMainMenu.frx":2C9F1
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":2CB43
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "Generate MPR Schedules"
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65290
            MouseIcon       =   "frmCSMSMainMenu.frx":2DBC5
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":2DD17
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Generate MPR Schedules"
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -65290
            MouseIcon       =   "frmCSMSMainMenu.frx":2ED99
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":2EEEB
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Generate MPR Schedules"
            Top             =   690
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_CustDir 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":2FF6D
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":300BF
            Style           =   1  'Graphical
            TabIndex        =   21
            Tag             =   "1050"
            ToolTipText     =   "View Customer Directory Listing Report"
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_Technician 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":31141
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":31293
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Generate Technician Report"
            Top             =   3240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_VehicleByModel 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":32315
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":32467
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "1053"
            ToolTipText     =   "View Vehicle By Model"
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_ServiceReport 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":334E9
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":3363B
            Style           =   1  'Graphical
            TabIndex        =   18
            Tag             =   "1054"
            ToolTipText     =   "Generate Service Report"
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_ServiceAdvisor 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":346BD
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":3480F
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Generate Service Advisor Performance Report"
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_VehicleAgingReport 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":35891
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":359E3
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Generate Vehicle Aging Report"
            Top             =   3240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_HyundaiDealerMonthlyPerformance 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   330
            MouseIcon       =   "frmCSMSMainMenu.frx":36A65
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":36BB7
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Generate Hyundai Dealer Monthly Performance Report"
            Top             =   2400
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_ActualManning 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   330
            MouseIcon       =   "frmCSMSMainMenu.frx":37C39
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":37D8B
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "View Actual Manning Report"
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_TransactionForFollowup 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   330
            MouseIcon       =   "frmCSMSMainMenu.frx":38E0D
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":38F5F
            Style           =   1  'Graphical
            TabIndex        =   13
            Tag             =   "1052"
            ToolTipText     =   "View Transaction For Follow-Up"
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_UnitsReceivedWeeklyPerformance 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   330
            MouseIcon       =   "frmCSMSMainMenu.frx":39FE1
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":3A133
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Generate Units Received Weekly Performance Report"
            Top             =   3240
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_AfterSales 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   10020
            MouseIcon       =   "frmCSMSMainMenu.frx":3B1B5
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":3B307
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "After Sales Service Report"
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_WorkshopSalesWeeklyPerformance 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   330
            MouseIcon       =   "frmCSMSMainMenu.frx":3BC5D
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":3BDAF
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Generate Workshop Sales Weekly Performance Report"
            Top             =   4080
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_WorkInProgress 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4200
            MouseIcon       =   "frmCSMSMainMenu.frx":3CE31
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":3CF83
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "View Work In Progress"
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_AppointmentDiary 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":3E005
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":3E157
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "View Appointment Diary"
            Top             =   4080
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":3F1D9
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":3F32B
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Generate Warranty Claim Report"
            Top             =   4080
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4200
            MouseIcon       =   "frmCSMSMainMenu.frx":403AD
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":404FF
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "View MPR Schedules"
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton Command9 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":41581
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":416D3
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Generate Time Clock/Job Clock Log-In"
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdEstimatedPMS 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -60910
            MouseIcon       =   "frmCSMSMainMenu.frx":42755
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":428A7
            TabIndex        =   4
            Tag             =   "1150"
            ToolTipText     =   "Monthly Vehicle Gross Profile Report"
            Top             =   3120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdSummary 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -60910
            MouseIcon       =   "frmCSMSMainMenu.frx":43355
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":434A7
            TabIndex        =   3
            ToolTipText     =   "Ranged Customer Directory By Customer Type"
            Top             =   3960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdReport_ServiceAdivsorSales 
            Height          =   735
            Left            =   -69670
            MouseIcon       =   "frmCSMSMainMenu.frx":43DA1
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":43EF3
            Style           =   1  'Graphical
            TabIndex        =   2
            Tag             =   "1087"
            ToolTipText     =   "Generate Service invoice summary report"
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Void R.O. Report"
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
            Height          =   255
            Index           =   2
            Left            =   -64270
            TabIndex        =   143
            Top             =   2640
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Gross Profit Reports"
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
            Height          =   225
            Left            =   -64270
            TabIndex        =   139
            Top             =   1800
            Visible         =   0   'False
            Width           =   2430
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Labor Cost Reports"
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
            Height          =   225
            Left            =   -64240
            TabIndex        =   135
            Top             =   990
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "After Sales Report"
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
            Height          =   225
            Index           =   7
            Left            =   -64330
            TabIndex        =   131
            Top             =   4290
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active Inactive Customer"
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
            Height          =   225
            Index           =   6
            Left            =   -64330
            TabIndex        =   128
            Top             =   3480
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unserved Sublet Purchase Order"
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
            Height          =   225
            Index           =   5
            Left            =   -64330
            TabIndex        =   126
            Top             =   2670
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet Sales Report"
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
            Height          =   225
            Index           =   4
            Left            =   -64330
            TabIndex        =   124
            Top             =   1830
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Summary of Internal Sales"
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
            Height          =   225
            Index           =   3
            Left            =   -64330
            TabIndex        =   121
            Top             =   990
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated PMS Return"
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
            Height          =   225
            Index           =   2
            Left            =   -59920
            TabIndex        =   44
            Top             =   3390
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Model Vehicle Report"
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
            Height          =   225
            Index           =   1
            Left            =   -68740
            TabIndex        =   43
            Top             =   1830
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Monthly Performance Report"
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
            Height          =   225
            Index           =   1
            Left            =   5160
            TabIndex        =   42
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Invoice Summary Report"
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
            Height          =   225
            Index           =   1
            Left            =   -68710
            TabIndex        =   41
            Top             =   2670
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Technician Report"
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
            Height          =   225
            Left            =   -68740
            TabIndex        =   40
            Top             =   3510
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Directory Listing Report"
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
            Height          =   225
            Index           =   0
            Left            =   -68740
            TabIndex        =   39
            Top             =   990
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle By Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   2
            Left            =   -68800
            TabIndex        =   38
            Top             =   1860
            Width           =   1440
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Report"
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
            Height          =   225
            Left            =   -68710
            TabIndex        =   37
            Top             =   1005
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Advisor Performance Report"
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
            Height          =   225
            Index           =   0
            Left            =   -68710
            TabIndex        =   36
            Top             =   1860
            Visible         =   0   'False
            Width           =   3120
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Aging Report"
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
            Height          =   225
            Left            =   -68710
            TabIndex        =   35
            Top             =   3510
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Appointment Diary"
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
            Height          =   225
            Left            =   -68710
            TabIndex        =   34
            Top             =   4290
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Work In Progress"
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
            Height          =   225
            Left            =   5160
            TabIndex        =   33
            Top             =   975
            Width           =   1500
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Sales Weekly Performance Report"
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
            Height          =   225
            Left            =   1260
            TabIndex        =   32
            Top             =   4290
            Width           =   3870
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "After Sales Service Report"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   10410
            TabIndex        =   31
            Top             =   2370
            Width           =   2700
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Units Received Weekly Performance Report"
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
            Height          =   225
            Left            =   1260
            TabIndex        =   30
            Top             =   3480
            Width           =   3720
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction For Follow-Up"
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
            Height          =   225
            Left            =   1260
            TabIndex        =   29
            Top             =   975
            Width           =   2205
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actual Manning Report"
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
            Height          =   225
            Left            =   1260
            TabIndex        =   28
            Top             =   1800
            Width           =   1920
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hyundai Dealer Monthly Performance Report"
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
            Height          =   225
            Left            =   1260
            TabIndex        =   27
            Top             =   2670
            Width           =   3765
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Advisor Sales Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   -68740
            TabIndex        =   26
            Top             =   2640
            Width           =   2475
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Warranty Report"
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
            Height          =   225
            Left            =   -68740
            TabIndex        =   25
            Top             =   4380
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Time Control Analysis"
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
            Height          =   225
            Left            =   -68740
            TabIndex        =   24
            Top             =   2670
            Visible         =   0   'False
            Width           =   2550
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Summary of Customer Suggetion/Reccomendation"
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
            Height          =   735
            Left            =   -59920
            TabIndex        =   23
            Top             =   4080
            Visible         =   0   'False
            Width           =   4275
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated PMS Return"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1000
            Left            =   -64300
            TabIndex        =   22
            Top             =   990
            Width           =   1905
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl3 
         Height          =   5175
         Left            =   3840
         TabIndex        =   52
         Top             =   570
         Width           =   5205
         _Version        =   655364
         _ExtentX        =   9181
         _ExtentY        =   9128
         _StockProps     =   64
         Appearance      =   2
         Color           =   4
         PaintManager.Layout=   2
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         PaintManager.FixedTabWidth=   130
         ItemCount       =   2
         Item(0).Caption =   "Warranty && Sublet"
         Item(0).ControlCount=   8
         Item(0).Control(0)=   "cmdReport_warrantyClaim"
         Item(0).Control(1)=   "cmdMain_QualityInformation"
         Item(0).Control(2)=   "Label48"
         Item(0).Control(3)=   "Label40"
         Item(0).Control(4)=   "cmdSubletRepair"
         Item(0).Control(5)=   "cmdSubletReceiving"
         Item(0).Control(6)=   "Label53"
         Item(0).Control(7)=   "Label52"
         Item(1).Caption =   "Others Modules"
         Item(1).ControlCount=   8
         Item(1).Control(0)=   "Command3"
         Item(1).Control(1)=   "Command8"
         Item(1).Control(2)=   "Command7"
         Item(1).Control(3)=   "Label46"
         Item(1).Control(4)=   "Label56"
         Item(1).Control(5)=   "Label55"
         Item(1).Control(6)=   "cmdMain_QC"
         Item(1).Control(7)=   "Label38"
         Begin VB.CommandButton cmdMain_QC 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   -69730
            MouseIcon       =   "frmCSMSMainMenu.frx":44F75
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":450C7
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "View Quality Control Inspection"
            Top             =   720
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CommandButton cmdSubletReceiving 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   270
            MouseIcon       =   "frmCSMSMainMenu.frx":46149
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":4629B
            Style           =   1  'Graphical
            TabIndex        =   115
            Tag             =   "1025"
            ToolTipText     =   "View Sublet Repair Receive"
            Top             =   3270
            Width           =   795
         End
         Begin VB.CommandButton cmdSubletRepair 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   270
            MouseIcon       =   "frmCSMSMainMenu.frx":4731D
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":4746F
            Style           =   1  'Graphical
            TabIndex        =   114
            Tag             =   "1294"
            ToolTipText     =   "View Sublet Repair Purchase"
            Top             =   2430
            Width           =   795
         End
         Begin VB.CommandButton cmdReport_warrantyClaim 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   270
            MouseIcon       =   "frmCSMSMainMenu.frx":484F1
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":48643
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "View Warranty Claim Report"
            Top             =   1605
            Width           =   795
         End
         Begin VB.CommandButton cmdMain_QualityInformation 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   270
            MouseIcon       =   "frmCSMSMainMenu.frx":496C5
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":49817
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "View Prior Work Approval Records"
            Top             =   765
            Width           =   795
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   -69730
            MouseIcon       =   "frmCSMSMainMenu.frx":4A899
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":4A9EB
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "View Material Requisition"
            Top             =   1620
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CommandButton Command8 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   -69730
            MouseIcon       =   "frmCSMSMainMenu.frx":4BA6D
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":4BBBF
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Work Bay Masterfile"
            Top             =   3420
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CommandButton Command7 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   -69730
            MouseIcon       =   "frmCSMSMainMenu.frx":4CC41
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSMainMenu.frx":4CD93
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "View Workshop Monitoring"
            Top             =   2520
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quality Control Inspection"
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
            Height          =   225
            Left            =   -68860
            TabIndex        =   119
            Top             =   1020
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet Repair Purchase Order"
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
            Height          =   225
            Left            =   1140
            TabIndex        =   117
            Top             =   2760
            Width           =   2550
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet Repair Receiving"
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
            Height          =   225
            Left            =   1140
            TabIndex        =   116
            Top             =   3495
            Width           =   2010
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accumulated Claim List (ACL)"
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
            Height          =   225
            Left            =   1140
            TabIndex        =   62
            Top             =   1890
            Width           =   2520
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prior Work Approval (PWA)"
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
            Height          =   225
            Left            =   1140
            TabIndex        =   61
            Top             =   1050
            Width           =   2310
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Materials Requisition"
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
            Height          =   225
            Left            =   -68860
            TabIndex        =   60
            Top             =   1920
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Bay Masterfile"
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
            Height          =   225
            Left            =   -68860
            TabIndex        =   59
            Top             =   3720
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Bay Monitoring"
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
            Height          =   225
            Left            =   -68860
            TabIndex        =   58
            Top             =   2850
            Visible         =   0   'False
            Width           =   2190
         End
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Counter"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   -64240
         TabIndex        =   141
         Top             =   1920
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Loyalty File Generation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   -64300
         TabIndex        =   137
         Top             =   1020
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Service Payment Master File"
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
         Height          =   675
         Left            =   -62920
         TabIndex        =   133
         Top             =   2775
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor/Contractor Master File"
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
         Height          =   675
         Index           =   1
         Left            =   -62950
         TabIndex        =   113
         Top             =   1875
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Maintenance "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -68710
         TabIndex        =   111
         Top             =   1050
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Profile"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -68710
         TabIndex        =   110
         Top             =   1920
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Inquiry"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -68710
         TabIndex        =   109
         Top             =   2730
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complaints Form"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -68710
         TabIndex        =   108
         Top             =   3540
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Concern Resolution"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -68710
         TabIndex        =   107
         Top             =   4350
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reminders"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -68710
         TabIndex        =   106
         Top             =   5130
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Advisor Work Details NEO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Index           =   1
         Left            =   -59680
         TabIndex        =   99
         ToolTipText     =   "View Parts Listing"
         Top             =   4440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Estimate Listing"
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
         Height          =   225
         Left            =   -59740
         TabIndex        =   97
         Top             =   5490
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Technician Attendance Monitoring"
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
         Height          =   225
         Left            =   -64600
         TabIndex        =   96
         Top             =   1125
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parts Inquiry"
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
         Height          =   225
         Left            =   -68785
         TabIndex        =   95
         Top             =   2865
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material Lubricant Inquiry"
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
         Height          =   225
         Index           =   0
         Left            =   -68785
         TabIndex        =   94
         ToolTipText     =   "View Parts Listing"
         Top             =   3750
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories Inquiry"
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
         Height          =   225
         Index           =   2
         Left            =   -68785
         TabIndex        =   93
         ToolTipText     =   "View Parts Listing"
         Top             =   4650
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Order Inquiry"
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
         Height          =   225
         Left            =   -68785
         TabIndex        =   88
         Top             =   1965
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Vehicle Inquiry"
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
         Height          =   225
         Left            =   -68785
         TabIndex        =   87
         Top             =   1140
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Time Standards"
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
         Height          =   225
         Left            =   -65920
         TabIndex        =   84
         Top             =   2025
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PMS Jobs"
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
         Height          =   225
         Left            =   -65920
         TabIndex        =   83
         Top             =   2940
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Canned Labor"
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
         Height          =   225
         Left            =   -65920
         TabIndex        =   82
         Top             =   3855
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Personnel Maintenance"
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
         Height          =   585
         Index           =   1
         Left            =   -62950
         TabIndex        =   81
         Top             =   945
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Jobs"
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
         Height          =   225
         Left            =   -65920
         TabIndex        =   80
         Top             =   1110
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Master File"
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
         Height          =   225
         Index           =   1
         Left            =   -65920
         TabIndex        =   79
         Top             =   4830
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Master List"
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
         Height          =   645
         Left            =   -68800
         TabIndex        =   72
         Top             =   1125
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Height          =   225
         Left            =   -68800
         TabIndex        =   71
         Top             =   2970
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Vehicle Maintenance"
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
         Height          =   795
         Index           =   0
         Left            =   -68800
         TabIndex        =   70
         Top             =   1860
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Make"
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
         Height          =   225
         Index           =   0
         Left            =   -68800
         TabIndex        =   69
         Top             =   4860
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Color"
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
         Height          =   225
         Left            =   -68800
         TabIndex        =   68
         Top             =   3915
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Billing System"
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
         Height          =   225
         Left            =   1260
         TabIndex        =   51
         Top             =   3090
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Counter"
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
         Height          =   225
         Index           =   0
         Left            =   1260
         TabIndex        =   50
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Estimate"
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
         Height          =   225
         Left            =   1260
         TabIndex        =   49
         Top             =   2235
         Width           =   1110
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Clock"
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
         Height          =   225
         Left            =   1260
         TabIndex        =   48
         Top             =   3915
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents frm                                As frmCSMS_MasterStockInquiry
Attribute frm.VB_VarHelpID = -1

Private Sub cmdAuditInquiry_Click()
    If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub

    frmALL_AuditInquiry.Show
End Sub

Private Sub cmdCompany_Profile_Click(Index As Integer)
    frmCSMSProfile.Show
End Sub

Private Sub cmdctr_Click()
    If Module_Access(LOGID, "INVOICE COUNTER", "DATA ENTRY") = False Then Exit Sub
    If frmCTR Is Nothing Then
        frmCTR.Show
    Else
        frmCTR.WindowState = 0
        frmCTR.ZOrder 0
    End If
    
End Sub

Private Sub cmdEstimatedPMS_Click()
    MsgBox "Module Under Develop", vbInformation, "CSMS"
    Exit Sub
    If Module_Access(LOGID, "ESTIMATED PMS RETURN", "REPORTS") = False Then Exit Sub
    frmCSMS_EstimatedPMS.Show
End Sub

Private Sub cmdGJBP_Click()
    'PROCESS ON PUTING JOBTYPE ON THOSE PARTS, MATERIALS AND ACCESSORIES WHICH JOBTYPE IS NULL
    Dim rsDet                                          As New ADODB.Recordset
    Dim RSORD                                          As New ADODB.Recordset
    Dim JOBTYPE                                        As String

    'FOR THIS MONTH TRANSACTION
    Set RSORD = gconDMIS.Execute("SELECT * FROM PMIS_ORD_HD WHERE TRANTYPE = 'RIV' AND STATUS NOT IN ('C','N')")
    If Not (RSORD.BOF And RSORD.EOF) Then
        Do While RSORD.EOF
            JOBTYPE = Null2String(RSORD!SI_TYPE)

            RSORD.MoveNext
        Loop
    End If


    'FOR THE HISTORY TRANSACTION
    Set RSORD = New ADODB.Recordset
    Set RSORD = gconDMIS.Execute("SELECT * FROM PMIS_ORD_HIST WHERE STATUS <> 'C'")
    If Not (RSORD.BOF And RSORD.EOF) Then
        'Do While RSORD.EOF
        'Null2String(RSORD!SI_TYPE)

        'RSORD.MoveNext
        'Loop
    End If

    Set RSORD = Nothing
End Sub

Private Sub cmdInq_Acc_Click()
    If Module_Access(LOGID, "ACCESSORIES INQUIRY", "INQUIRY") = False Then Exit Sub
    
    Set frm = New frmCSMS_MasterStockInquiry
    Call frm.SetType("A", "Accessories Stock Inquiry")
    frm.Show
End Sub

Private Sub cmdInq_Accessoires_Click()
    If Module_Access(LOGID, "MATERIAL INQUIRY", "INQUIRY") = False Then Exit Sub
    
    Set frm = New frmCSMS_MasterStockInquiry
    Call frm.SetType("M", "Material Stock Inquiry")
    frm.Show
End Sub

Private Sub cmdInq_CusVeh_Click()
    If Module_Access(LOGID, "CUSTOMER VEHICLE INQUIRY", "INQUIRY") = False Then Exit Sub

    frmCSMSCustomerHistory.Show
    frmCSMSCustomerHistory.ZOrder 0
End Sub

Private Sub cmdInq_JobEsitmateListing_Click()
    If Module_Access(LOGID, "JOB ESTIMATE LISTING", "INQUIRY") = False Then Exit Sub

    frmCSMS_INQUIRY_JobEstimate.Show
End Sub

Private Sub cmdInq_Parts_Click()
    If Module_Access(LOGID, "PARTS INQUIRY", "INQUIRY") = False Then Exit Sub

    Set frm = New frmCSMS_MasterStockInquiry
    Call frm.SetType("P", "Parts Stock Inquiry")
    frm.Show
End Sub

Private Sub cmdInq_ServiceAdvisorWorkDetail_Click()
    If Module_Access(LOGID, "REPAIR ORDER INQUIRY", "INQUIRY") = False Then Exit Sub
    
    frmCSMS_MasterRepairInquiry.Show
End Sub

Private Sub cmdInq_Techmonitoring_Click()
    If Module_Access(LOGID, "TECHNICIAN MONITORING", "INQUIRY") = False Then Exit Sub

    frmCSMSTechnicianMonitoring.Show
    frmCSMSTechnicianMonitoring.ZOrder 0
End Sub

Private Sub cmdMain_BillingSystem_Click()
    If Module_Access(LOGID, "BILLING SYSTEM", "TRANSACTION") = False Then Exit Sub
    frmCSMSDataEntry.Show
End Sub

Private Sub cmdMain_Complaint_Click()
    If Module_Access(LOGID, "COMPLAINTS FORM", "DATA ENTRY") = False Then Exit Sub

    FrmCSMSComplaintsForm.Show
    FrmCSMSComplaintsForm.ZOrder 0
End Sub

Private Sub cmdMain_ConcernResolution_Click()
    If Module_Access(LOGID, "CONCERN RESOLUTION", "DATA ENTRY") = False Then Exit Sub

    frmCSMSConcernResolution.Show
    frmCSMSConcernResolution.ZOrder 0
End Sub

Private Sub cmdMain_JobClock_Click()
    If Module_Access(LOGID, "JOB CLOCK", "SYSTEM") = False Then Exit Sub

    frmCSMSClockINOUT.Show
    frmCSMSClockINOUT.ZOrder 0
End Sub

Private Sub cmdMain_JobEstimate_Click()
    If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub

    frmCSMSEstimateEntry.Show
    frmCSMSEstimateEntry.ZOrder 0
End Sub

Private Sub cmdMain_QC_Click()
    If Module_Access(LOGID, "QUALITY INSPECTION", "TRANSACTION") = False Then Exit Sub

    If QC_MODULE_ON = "ON" Then
        frmCSMS_QCInspection.Show
        frmCSMS_QCInspection.ZOrder 0
    Else
        MsgBox "Quality Control Inpection Module is OFF", vbInformation, "CSMS"
        Exit Sub
    End If
End Sub

Private Sub cmdMain_QualityInformation_Click()
    If Module_Access(LOGID, "QUALITY INFORMATION REPORT", "TRANSACTION") = False Then Exit Sub
    frmCSMS_CQI.Show
    frmCSMS_CQI.ZOrder 0
End Sub

Private Sub cmdMain_ServiceCounter_Click()
    If Module_Access(LOGID, "SERVICE COUNTER", "SYSTEM") = False Then Exit Sub
    frmCSMS_ServiceCounter.Show
End Sub

Private Sub cmdMake_Click()
    If Module_Access(LOGID, "MAKE", "DATA ENTRY") = False Then Exit Sub
    frmCSMS_MAKE.Show
End Sub

Private Sub cmdMaster_CannedLabor_Click()
    If Module_Access(LOGID, "CANNED LABOR", "DATA ENTRY") = False Then Exit Sub

    frmCSMSCannedlabor.Show
End Sub

Private Sub cmdMaster_Customer_Click()
    If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub

    frmAllCustomer.Show
    frmAllCustomer.ZOrder 0
End Sub

Private Sub cmdMaster_LTS_Click()
    If Module_Access(LOGID, "JOBS", "DATA ENTRY") = False Then Exit Sub

    frmCSMSReqJobs.Show
    frmCSMSReqJobs.ZOrder 0
End Sub

Private Sub cmdMaster_Model_Click()
    If Module_Access(LOGID, "MODEL", "DATA ENTRY") = False Then Exit Sub

    frmCSMSModel.Show
    frmCSMSModel.ZOrder 0
End Sub

Private Sub cmdMaster_OtherJobs_Click()
    If Module_Access(LOGID, "JOBS", "DATA ENTRY") = False Then Exit Sub

    Screen.MousePointer = 11
    frmCSMSJobs.Show
    frmCSMSJobs.ZOrder 0
    Screen.MousePointer = 0
End Sub

Private Sub cmdMaster_PMS_Click()
    If Module_Access(LOGID, "PMS JOBS", "DATA ENTRY") = False Then Exit Sub

    frmCSMSAddPms.Show
    frmCSMSAddPms.ZOrder 0
End Sub

Private Sub cmdPassword_Click()
    frmAccMaintenance.Show
    frmAccMaintenance.ZOrder 0
End Sub

Private Sub cmdReport_ActualManning_Click()
    If Module_Access(LOGID, "ACTUAL MANNING REPORT", "REPORTS") = False Then Exit Sub

    frmCSMSActualManningReport.Show
End Sub

Private Sub cmdReport_AfterSales_Click()
    If Module_Access(LOGID, "AFTER SALES REPORT", "REPORTS") = False Then Exit Sub

    frmCSMSAfterSalesServiceReport.Show
    frmCSMSAfterSalesServiceReport.ZOrder 0
End Sub

Private Sub cmdReport_AppointmentDiary_Click()
    If Module_Access(LOGID, "APPOINTMENT DIARY", "REPORTS") = False Then Exit Sub

    frmCSMSAppointmentDiary.Show
End Sub

Private Sub cmdReport_CustDir_Click()
    If Module_Access(LOGID, "CUSTOMER DIRECTORY LISTING", "REPORTS") = False Then Exit Sub
    frmCSMS_DirectoryLisiting.Show
End Sub

Private Sub cmdReport_HyundaiDealerMonthlyPerformance_Click()
    If Module_Access(LOGID, "MONTHLY PERFORMANCE REPORT", "REPORTS") = False Then Exit Sub

    frmCSMSHyundaiMonthlyPerformanceReport.Show
    frmCSMSHyundaiMonthlyPerformanceReport.ZOrder 0
End Sub

Private Sub cmdReport_PartsPickList_Click()
    If Module_Access(LOGID, "PARTS PICK LIST", "REPORTS") = False Then Exit Sub
End Sub

Private Sub cmdReport_ServiceAdivsorSales_Click()
    If Module_Access(LOGID, "SERVICE ADVISOR SALES", "REPORTS") = False Then Exit Sub

    frmCSMS_Reports_SASales.Show
    frmCSMS_Reports_SASales.ZOrder 0
End Sub

Private Sub cmdReport_ServiceAdvisor_Click()
    If Module_Access(LOGID, "SERVICE ADVISOR REPORT", "REPORTS") = False Then Exit Sub

    frmCSMSServiceAdvisorReport.Show
    frmCSMSServiceAdvisorReport.ZOrder 0
End Sub

Private Sub cmdReport_ServiceReport_Click()
    If Module_Access(LOGID, "SERVICE REPORT", "REPORTS") = False Then Exit Sub

    frmCSMSServiceReport.Show
    frmCSMSServiceReport.ZOrder 0
End Sub

Private Sub cmdReport_Technician_Click()
    If Module_Access(LOGID, "TECHNICIAN REPORT", "REPORTS") = False Then Exit Sub

    frmTechnicianReport.Show
    frmTechnicianReport.ZOrder 0
End Sub

Private Sub cmdReport_TransactionForFollowup_Click()
    If Module_Access(LOGID, "TRANSACTIONS FOR FOLLOW UP", "REPORTS") = False Then Exit Sub

    frmCSMSFor_followUp.Show
    frmCSMSFor_followUp.ZOrder 0
End Sub

Private Sub cmdReport_UnitsReceivedWeeklyPerformance_Click()
    If Module_Access(LOGID, "UNITS RECEIVE WEEKLY PERFORMANCE REPORT", "REPORTS") = False Then Exit Sub

    frmCSMSUnitsReceivedWeeklyPerformanceReport.Show
    frmCSMSUnitsReceivedWeeklyPerformanceReport.ZOrder 0
End Sub

Private Sub cmdReport_VehicleAgingReport_Click()
    If Module_Access(LOGID, "VEHICLE AGING PROGRESS REPORT", "REPORTS") = False Then Exit Sub

    frmCSMSVehicleAgingOnProcessReport.Show
    frmCSMSVehicleAgingOnProcessReport.ZOrder 0
End Sub

Private Sub cmdReport_VehicleByModel_Click()
    If Module_Access(LOGID, "VEHICLE BY MODEL", "REPORTS") = False Then Exit Sub

    frmCSMSVehicleByModel.Show
    frmCSMSVehicleByModel.ZOrder 0
End Sub

Private Sub cmdReport_warrantyClaim_Click()
    If Module_Access(LOGID, "ACCUMULATED CLAIM LIST", "TRANSACTION") = False Then Exit Sub
    frmCSMS_ACL.Show
    frmCSMS_ACL.ZOrder 0
End Sub

Private Sub cmdReport_WorkInProgress_Click()
    If Module_Access(LOGID, "WORKING IN PROGRRESS", "REPORTS") = False Then Exit Sub

    frmCSMSWorkInProgress.Show
    frmCSMSWorkInProgress.ZOrder 0
End Sub

Private Sub cmdReport_WorkshopSalesWeeklyPerformance_Click()
    If Module_Access(LOGID, "WORKSHOP SALES WEEKLY PERFORMANCE REPORT", "REPORTS") = False Then Exit Sub

    frmCSMSWorkshopSalesWeeklyPerformanceReport.Show
    frmCSMSWorkshopSalesWeeklyPerformanceReport.ZOrder 0
End Sub

Private Sub cmdServiceInternal_Click()
    If Module_Access(LOGID, "SERVICE PAYMENT METHODS", "DATA ENTRY") = False Then Exit Sub
    frmCSMSInternalPayments.Show
End Sub

Private Sub cmdServiceMaintenace_Click()
    If Module_Access(LOGID, "SERVICE PERSONNEL MAINTAINANCE", "DATA ENTRY") = False Then Exit Sub
    frmCSMS_ServicePersonnel.Show
End Sub

Private Sub cmdSubletReceiving_Click()
    If Module_Access(LOGID, "SUBLET RECEIVING", "TRANSACTION") = False Then Exit Sub
    frmCSMS_ReceivingEntry.Show
End Sub

Private Sub cmdSubletRepair_Click()
    If Module_Access(LOGID, "SUBLET PURCHASE", "TRANSACTION") = False Then Exit Sub
    frmCSMS_PurchaseOrder.Show
End Sub

Private Sub cmdSummary_Click()
    MsgBox "Module Under Develop", vbInformation, "CSMS"
End Sub

Private Sub cmdTab_Color_Click()
    If Module_Access(LOGID, "VEHICLE COLOR", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_Color.Show
    frmSMIS_Files_Color.ZOrder 0
End Sub

Private Sub cmdTab_CustInfo_Click()

End Sub

Private Sub cmdTable_DealerMaster_Click()
    If Module_Access(LOGID, "SELLING DEALER", "DATA ENTRY") = False Then Exit Sub
    frmCSMS_Files_SellingDealer.Show
End Sub

Private Sub Command1_Click()
    'UPDATE BY   : MJP 11092009 1127PM
    'DESCRIPTION : CRF 140
        If Module_Access(LOGID, "SUMMARY OF INTERNAL SALES", "REPORTS") = False Then Exit Sub
        frmCSMS_Report_SummaryofInternalSales.Show
    'UPDATE BY   : MJP 11092009 1127PM
End Sub

Private Sub Command10_Click()
    frmSMIS_Log_Reminder.Show
End Sub

Private Sub Command11_Click()
    If Module_Access(LOGID, "VENDORS", "DATA ENTRY") = False Then Exit Sub
    frmAMISMASTERFILEVendor.Show
End Sub

Private Sub Command12_Click()
    'UPDATE BY   : MJP 11082009 0300PM
    'DESCRIPTION : CRF 140
    If Module_Access(LOGID, "UNSERVED SUBLET PO", "REPORTS") = False Then Exit Sub
    frmCSMS_Report_UnservedPO.Show
    'UPDATE BY   : MJP 11082009 0300PM
End Sub

Private Sub Command13_Click()
    'UPDATE BY   : EAP 11112009 0400PM
    'DESCRIPTION : CRF 137
    If Module_Access(LOGID, "ACTIVE INACTIVE CUSTOMER", "REPORTS") = False Then Exit Sub
    frmCSMS_Reports_ActiveInactive.Show
    'UPDATE BY   : EAP 11112009 0400PM
End Sub

Private Sub Command14_Click()
     'UPDATE BY   : NVB 12/11/2009 10AM
     'DESCRIPTION : CRF
    
    If Module_Access(LOGID, "LABOR COST", "REPORTS") = False Then Exit Sub
    frmReportsLaborcost.Show
End Sub

Private Sub Command15_Click()
    'frmCSMS_Reports_History.Show
    'If COMPANY_CODE <> "HAI" Then
    '    MessagePop InfoFriend, "Module Info.", "This module is not supported by your dealer. For more information Kindly contact Netspeed Software Inc. about this Module"
    '    Exit Sub
    'End If
    
    If Module_Access(LOGID, "Generate Loyalty File", "SYSTEM") = False Then Exit Sub
    frmCSMS_Loyalty.Show

End Sub

Private Sub Command16_Click()
     'UPDATE BY   : NVB 12/11/2009 10AM
     'DESCRIPTION : CRF
        If Module_Access(LOGID, "MONTHLY SALES CUSTOMER", "REPORTS") = False Then Exit Sub
        frmCSMS_Report_AfterSales.Show
End Sub



Private Sub Command17_Click()
'Updated By: IEVB 04152011
'description:   for tcn # 13402
If Module_Access(LOGID, "SERVICE GROSS PROFUT REPORT", "REPORTS") = False Then Exit Sub
    If frmcsms_grossprofitreport Is Nothing Then
        frmcsms_grossprofitreport.Show
    Else
        frmcsms_grossprofitreport.WindowState = 0
        frmcsms_grossprofitreport.ZOrder 0
    End If
End Sub

Private Sub Command19_Click()
If Module_Access(LOGID, "VOID R.O. REPORT", "REPORTS") = False Then Exit Sub
    If frmcsms_report_voidro Is Nothing Then
        frmcsms_report_voidro.Show
    Else
        frmcsms_report_voidro.WindowState = 0
        frmcsms_report_voidro.ZOrder 0
    End If
End Sub

Private Sub Command2_Click()
    If Module_Access(LOGID, "WARRANTY REPORTS", "REPORTS") = False Then Exit Sub
    frmCSMS_WarRep.Show
End Sub

Private Sub Command3_Click()
    If Module_Access(LOGID, "MATERIALS REQUISITION SLIP", "TRANSACTION") = False Then Exit Sub
    frmPMISMAT_MRISForms.Show
End Sub

Private Sub Command4_Click()
    'UPDATE BY   : MJP 11082009 0200PM
    'DESCRIPTION : CRF 141
    If Module_Access(LOGID, "SUBLET SALES REPORTS", "REPORTS") = False Then Exit Sub
    frmCSMS_Report_SubletSales.Show
    'UPDATE BY   : MJP 11082009 0200PM
End Sub

Private Sub Command5_Click()
    If Module_Access(LOGID, "CUSTOMER VEHICLE", "DATA ENTRY") = False Then Exit Sub
    frmCSMSEditCustomerVehicle_neo.Show
End Sub

Private Sub Command6_Click()
    If Module_Access(LOGID, "SCHEDULE MPR", "REPORTS") = False Then Exit Sub
    frmCSMSScheduleMPR.Show
End Sub

Private Sub Command7_Click()
    If Module_Access(LOGID, "BAY MONITORING", "SYSTEM") = False Then Exit Sub
    frmBayMonitoring.Show 1
End Sub

Private Sub Command8_Click()
    If Module_Access(LOGID, "BAY MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    frmBayMasterFile.Show
End Sub

Private Sub Command9_Click()
    If Module_Access(LOGID, "MONTHLY TIME CONTROL ANALYSIS", "REPORTS") = False Then Exit Sub
    frmCSMSMonthlyTimeControlAnalysis.Show
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    QC_MODULE_ON = ""

    TabControl1.SelectedItem = 0
End Sub

Private Sub Label2_Click()
'On Error GoTo adoerror
    gconDMIS.Execute ("insert into ALL_Color values ('Yd','SILVER')")
Exit Sub
 
End Sub

Private Sub Picture6_Click()

End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If TabControl1.SelectedItem = 3 Then
        TabControl2.SelectedItem = 0
    End If
End Sub


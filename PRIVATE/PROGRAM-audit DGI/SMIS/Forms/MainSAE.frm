VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form MainSAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Diary"
   ClientHeight    =   8175
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   13710
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "MainSAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   914
   Begin VB.ComboBox cboYear 
      Height          =   345
      ItemData        =   "MainSAE.frx":030A
      Left            =   90
      List            =   "MainSAE.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   90
      Width           =   2115
   End
   Begin VB.ListBox lstMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   3255
      IntegralHeight  =   0   'False
      Left            =   60
      MouseIcon       =   "MainSAE.frx":030E
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   510
      Width           =   2145
   End
   Begin Crystal.CrystalReport rptMain 
      Left            =   4980
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   51
      Top             =   7740
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Object.ToolTipText     =   "Agent Code"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   6615
            MinWidth        =   6615
            Object.ToolTipText     =   "Agent Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Object.ToolTipText     =   "Login Time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2:17 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock Status"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock Status"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Scroll Lock Status"
         EndProperty
      EndProperty
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
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   7350
      Left            =   2310
      TabIndex        =   5
      Top             =   0
      Width           =   11415
      _Version        =   655364
      _ExtentX        =   20135
      _ExtentY        =   12965
      _StockProps     =   64
      AutoResizeClient=   0   'False
      Appearance      =   3
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   6
      Item(0).Caption =   "Prospects"
      Item(0).Tooltip =   "Prospects"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "lstProspects"
      Item(0).Control(1)=   "picOptionProspects"
      Item(0).Control(2)=   "Picture2"
      Item(0).Control(3)=   "txtSearch_Prospect"
      Item(1).Caption =   "Sales Order"
      Item(1).Tooltip =   "Sales Order"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "lstSalesOrder"
      Item(1).Control(1)=   "picOptSalesOrder"
      Item(1).Control(2)=   "txtSearch_SalesOrder"
      Item(2).Caption =   "Loan Application"
      Item(2).Tooltip =   "Loan Application"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "picOptLoan(0)"
      Item(2).Control(1)=   "picOptLoan(1)"
      Item(2).Control(2)=   "txtSearch_LoanApplication"
      Item(2).Control(3)=   "lstLoanApplication"
      Item(3).Caption =   "Vehicles Inquiry"
      Item(3).Tooltip =   "Vehicles Inquiry"
      Item(3).ControlCount=   3
      Item(3).Control(0)=   "lstVehicles"
      Item(3).Control(1)=   "picOptLoan(2)"
      Item(3).Control(2)=   "txtSearch_Vehicles"
      Item(4).Caption =   "Reports && Activity Inquiry"
      Item(4).ControlCount=   7
      Item(4).Control(0)=   "cmdOther_SAEPerformance_1"
      Item(4).Control(1)=   "Label36"
      Item(4).Control(2)=   "Combo1"
      Item(4).Control(3)=   "Label3"
      Item(4).Control(4)=   "Label4"
      Item(4).Control(5)=   "txtSearch_Activity"
      Item(4).Control(6)=   "lstActivity"
      Item(5).Caption =   "Reminders"
      Item(5).ControlCount=   2
      Item(5).Control(0)=   "Picture1"
      Item(5).Control(1)=   "lstReminders"
      Begin XtremeReportControl.ReportControl lstLoanApplication 
         Height          =   6210
         Left            =   -69940
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   11325
         _Version        =   655364
         _ExtentX        =   19976
         _ExtentY        =   10954
         _StockProps     =   64
         BorderStyle     =   4
         PreviewMode     =   -1  'True
      End
      Begin XtremeReportControl.ReportControl lstActivity 
         Height          =   5835
         Left            =   -69850
         TabIndex        =   80
         Top             =   1380
         Visible         =   0   'False
         Width           =   11175
         _Version        =   655364
         _ExtentX        =   19711
         _ExtentY        =   10292
         _StockProps     =   64
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
      End
      Begin XtremeReportControl.ReportControl lstReminders 
         Height          =   6120
         Left            =   -69940
         TabIndex        =   69
         Top             =   1200
         Visible         =   0   'False
         Width           =   11325
         _Version        =   655364
         _ExtentX        =   19976
         _ExtentY        =   10795
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl lstProspects 
         Height          =   6135
         Left            =   30
         TabIndex        =   7
         Top             =   1170
         Width           =   8550
         _Version        =   655364
         _ExtentX        =   15081
         _ExtentY        =   10821
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl lstVehicles 
         Height          =   6165
         Left            =   -69940
         TabIndex        =   22
         Top             =   1110
         Visible         =   0   'False
         Width           =   11280
         _Version        =   655364
         _ExtentX        =   19897
         _ExtentY        =   10874
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl lstSalesOrder 
         Height          =   6210
         Left            =   -69940
         TabIndex        =   59
         Top             =   1080
         Visible         =   0   'False
         Width           =   11310
         _Version        =   655364
         _ExtentX        =   19950
         _ExtentY        =   10954
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.CommandButton cmdOther_SAEPerformance_1 
         Height          =   645
         Left            =   -69820
         MouseIcon       =   "MainSAE.frx":0460
         MousePointer    =   99  'Custom
         Picture         =   "MainSAE.frx":05B2
         Style           =   1  'Graphical
         TabIndex        =   65
         Tag             =   "1156"
         ToolTipText     =   "Sales Executive Performance Report"
         Top             =   690
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtSearch_Activity 
         Height          =   375
         Left            =   -65320
         TabIndex        =   82
         Top             =   690
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   -61690
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   690
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   -69940
         ScaleHeight     =   525
         ScaleWidth      =   11295
         TabIndex        =   68
         Top             =   570
         Visible         =   0   'False
         Width           =   11295
         Begin VB.ComboBox cboStatus 
            Height          =   345
            Left            =   6270
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   120
            Width           =   2085
         End
         Begin VB.OptionButton optReminder 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Internal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   405
            Index           =   3
            Left            =   1060
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   60
            Width           =   975
         End
         Begin VB.OptionButton optReminder 
            BackColor       =   &H0080FFFF&
            Caption         =   "&All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   405
            Index           =   2
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   60
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.ComboBox cboPriority 
            Height          =   345
            Left            =   9120
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   90
            Width           =   2085
         End
         Begin VB.OptionButton optReminder 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Customers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   405
            Index           =   1
            Left            =   2060
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   60
            Width           =   1035
         End
         Begin VB.OptionButton optReminder 
            BackColor       =   &H8000000D&
            Caption         =   "&Prospect"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   60
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Status"
            Height          =   285
            Left            =   5610
            TabIndex        =   77
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Priority"
            Height          =   285
            Left            =   8460
            TabIndex        =   73
            Top             =   150
            Width           =   795
         End
      End
      Begin VB.PictureBox picOptSalesOrder 
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   -66550
         ScaleHeight     =   450
         ScaleWidth      =   7710
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   7710
         Begin VB.OptionButton optSO 
            BackColor       =   &H00008000&
            Caption         =   "&On Process"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   58
            Tag             =   "On Process"
            Top             =   60
            Width           =   1140
         End
         Begin VB.OptionButton optSO 
            Caption         =   "&All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   4788
            Style           =   1  'Graphical
            TabIndex        =   57
            Tag             =   "All"
            Top             =   60
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optSO 
            BackColor       =   &H8000000D&
            Caption         =   "&Invoiced"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   2
            Left            =   2424
            Style           =   1  'Graphical
            TabIndex        =   56
            Tag             =   "Invoiced"
            Top             =   60
            Width           =   1140
         End
         Begin VB.OptionButton optSO 
            BackColor       =   &H00000080&
            Caption         =   "&Cancelled"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   1
            Left            =   1242
            Style           =   1  'Graphical
            TabIndex        =   55
            Tag             =   "Cancelled"
            Top             =   60
            Width           =   1140
         End
         Begin VB.CommandButton cmdPrintSO 
            Caption         =   "Print"
            Height          =   330
            Left            =   5970
            TabIndex        =   54
            Top             =   90
            Width           =   1140
         End
         Begin VB.OptionButton optSO 
            Caption         =   "&Released"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   3606
            Style           =   1  'Graphical
            TabIndex        =   53
            Tag             =   "All"
            Top             =   60
            Width           =   1140
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   6675
         Left            =   8610
         ScaleHeight     =   6645
         ScaleWidth      =   2715
         TabIndex        =   34
         Top             =   630
         Width           =   2745
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   3195
            Left            =   0
            ScaleHeight     =   3195
            ScaleWidth      =   4050
            TabIndex        =   95
            Top             =   3900
            Width           =   4050
            Begin VB.TextBox txtCusName 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   465
               HideSelection   =   0   'False
               Left            =   30
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   100
               TabStop         =   0   'False
               Top             =   780
               Width           =   2655
            End
            Begin VB.TextBox txtCusEmail 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               HideSelection   =   0   'False
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   99
               TabStop         =   0   'False
               Top             =   1260
               Width           =   1995
            End
            Begin VB.TextBox txtCusAdd 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   645
               HideSelection   =   0   'False
               Left            =   30
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   98
               TabStop         =   0   'False
               Top             =   2070
               Width           =   2655
            End
            Begin VB.TextBox txtCusPhone 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               HideSelection   =   0   'False
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   1530
               Width           =   1995
            End
            Begin VB.TextBox lblNotes 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   30
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   96
               Top             =   30
               Width           =   2655
            End
            Begin VB.Label lblCustAddress 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Address"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   30
               TabIndex        =   103
               Top             =   1800
               Width           =   2655
            End
            Begin VB.Label lblCustPhone 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Phone:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   30
               TabIndex        =   102
               Top             =   1530
               Width           =   675
            End
            Begin VB.Label lblCustEmail 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Email :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   30
               TabIndex        =   101
               Top             =   1260
               Width           =   675
            End
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   30
            TabIndex        =   67
            Top             =   3330
            Width           =   2640
         End
         Begin VB.Label lblAgeing 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   270
            Left            =   30
            TabIndex        =   104
            Top             =   3630
            Width           =   2640
         End
         Begin VB.Label lblLogQuote 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   94
            ToolTipText     =   " Last Quotation Send "
            Top             =   330
            Width           =   1695
         End
         Begin VB.Label lblQ 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Quotation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   93
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblInvoiceNo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   92
            Top             =   3030
            Width           =   1695
         End
         Begin VB.Label lblLogLetters 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   91
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Label lblLogCalls 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   90
            Top             =   1830
            Width           =   1695
         End
         Begin VB.Label lblLogTestDrive 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   89
            ToolTipText     =   " Test Drive Schedules On and Day Elasped"
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label lblLogAppointment 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   88
            ToolTipText     =   "Last sales appointment made on and days elasped"
            Top             =   932
            Width           =   1695
         End
         Begin VB.Label lblLogVisits 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   87
            Top             =   1530
            Width           =   1695
         End
         Begin VB.Label lblLogLoan 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   86
            Top             =   2730
            Width           =   1695
         End
         Begin VB.Label lblLogEmail 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   85
            Top             =   2430
            Width           =   1695
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Invoice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   84
            Top             =   3030
            Width           =   945
         End
         Begin XtremeShortcutBar.ShortcutCaption captionInformation 
            Height          =   315
            Left            =   0
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Width           =   2715
            _Version        =   655364
            _ExtentX        =   4789
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Profile @"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
         Begin VB.Label lblEmails 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   43
            Top             =   2430
            Width           =   945
         End
         Begin VB.Label lblLogSalesOrder 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   42
            Top             =   1230
            Width           =   1695
         End
         Begin VB.Label lblSalesOrder 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Sales order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   41
            Top             =   1230
            Width           =   945
         End
         Begin VB.Label lblLoan 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Loan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   40
            Top             =   2730
            Width           =   945
         End
         Begin VB.Label lblVisits 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Visits"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   39
            Top             =   1530
            Width           =   945
         End
         Begin VB.Label lblTest 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Test Drive"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   38
            Top             =   630
            Width           =   945
         End
         Begin VB.Label lblLetters 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Letters"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   37
            Top             =   2130
            Width           =   945
         End
         Begin VB.Label lblCalls 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Calls"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   36
            Top             =   1830
            Width           =   945
         End
         Begin VB.Label lblAppt 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Appointment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   35
            Top             =   930
            Width           =   945
         End
      End
      Begin VB.PictureBox picOptionProspects 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   2730
         ScaleHeight     =   465
         ScaleWidth      =   5925
         TabIndex        =   14
         Top             =   630
         Width           =   5925
         Begin VB.OptionButton optProspects 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&All"
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
            Height          =   360
            Index           =   4
            Left            =   4740
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   60
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H00C00000&
            Caption         =   "&Follow Ups"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   3
            Left            =   3585
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   60
            Width           =   1140
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H0000C000&
            Caption         =   "&Open"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   60
            Width           =   1140
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Inactive"
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
            Height          =   360
            Index           =   2
            Left            =   2430
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   60
            Width           =   1140
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H000040C0&
            Caption         =   "&Closed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   1
            Left            =   1275
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   60
            Width           =   1140
         End
      End
      Begin VB.TextBox txtSearch_Vehicles 
         Height          =   375
         Left            =   -69910
         TabIndex        =   23
         Top             =   660
         Visible         =   0   'False
         Width           =   5205
      End
      Begin VB.TextBox txtSearch_Prospect 
         Height          =   375
         Left            =   90
         TabIndex        =   21
         Top             =   690
         Width           =   2610
      End
      Begin VB.TextBox txtSearch_SalesOrder 
         Height          =   375
         Left            =   -69910
         TabIndex        =   20
         Top             =   645
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtSearch_LoanApplication 
         Height          =   375
         Left            =   -69910
         TabIndex        =   19
         Top             =   690
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.PictureBox picOptLoan 
         Height          =   855
         Index           =   0
         Left            =   -67210
         ScaleHeight     =   795
         ScaleWidth      =   5715
         TabIndex        =   13
         Top             =   3090
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.PictureBox picOptLoan 
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   1
         Left            =   -66550
         ScaleHeight     =   435
         ScaleWidth      =   7545
         TabIndex        =   8
         Top             =   660
         Visible         =   0   'False
         Width           =   7545
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00404040&
            Caption         =   "A&ll"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   5
            Left            =   5850
            Style           =   1  'Graphical
            TabIndex        =   81
            Tag             =   "C"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00404040&
            Caption         =   "&Cancelled"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   4
            Left            =   4701
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "C"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H000000C0&
            Caption         =   "&Disapproved"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   2
            Left            =   2328
            Style           =   1  'Graphical
            TabIndex        =   12
            Tag             =   "D"
            Top             =   30
            Width           =   1215
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00000080&
            Caption         =   "&Pending"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   1
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   11
            Tag             =   "P"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00800000&
            Caption         =   "&On Process"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   0
            Left            =   1179
            Style           =   1  'Graphical
            TabIndex        =   10
            Tag             =   "O"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00008000&
            Caption         =   "&Approved"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   3
            Left            =   3552
            Style           =   1  'Graphical
            TabIndex        =   9
            Tag             =   "A"
            Top             =   30
            Value           =   -1  'True
            Width           =   1140
         End
      End
      Begin VB.PictureBox picOptLoan 
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   2
         Left            =   -64510
         ScaleHeight     =   435
         ScaleWidth      =   5865
         TabIndex        =   26
         Top             =   645
         Visible         =   0   'False
         Width           =   5865
         Begin VB.OptionButton optInventory 
            BackColor       =   &H000000C0&
            Caption         =   "&Invoiced"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   1
            Left            =   1188
            Style           =   1  'Graphical
            TabIndex        =   33
            Tag             =   "S"
            ToolTipText     =   "View Invoiced Vehicles"
            Top             =   30
            Width           =   1140
         End
         Begin VB.CommandButton cmdPrintVehicles 
            Caption         =   "Print"
            Height          =   360
            Left            =   4710
            TabIndex        =   30
            ToolTipText     =   "Print Details"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optInventory 
            BackColor       =   &H00404040&
            Caption         =   "&Released"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   2
            Left            =   2361
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "R"
            ToolTipText     =   "View Released Vehicles"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optInventory 
            BackColor       =   &H00008000&
            Caption         =   "&On Stock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   3
            Left            =   3534
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "O"
            ToolTipText     =   "View On Stock Vehicles"
            Top             =   30
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optInventory 
            BackColor       =   &H00000080&
            Caption         =   "&Allocated"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   0
            Left            =   15
            Style           =   1  'Graphical
            TabIndex        =   27
            Tag             =   "A"
            ToolTipText     =   "View Allocated Vehicles"
            Top             =   30
            Width           =   1140
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Filter"
         Height          =   285
         Left            =   -65800
         TabIndex        =   83
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Inquiry Stage"
         Height          =   285
         Left            =   -62860
         TabIndex        =   79
         Top             =   750
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "My Performance "
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
         Height          =   300
         Left            =   -69040
         TabIndex        =   66
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4035
      Left            =   30
      TabIndex        =   0
      Top             =   3690
      Width           =   2205
      Begin VB.CommandButton Command10 
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":0DCD
         MousePointer    =   99  'Custom
         TabIndex        =   106
         ToolTipText     =   "View Customer"
         Top             =   3660
         Width           =   1995
      End
      Begin VB.CommandButton Command9 
         Caption         =   "F7 LOAN APP (CORP)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":0F1F
         MousePointer    =   99  'Custom
         TabIndex        =   64
         ToolTipText     =   "View Sales Calculator"
         Top             =   2598
         Width           =   1995
      End
      Begin VB.CommandButton Command8 
         Caption         =   "F6 LOAN APP (INDIV)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":1071
         MousePointer    =   99  'Custom
         TabIndex        =   63
         ToolTipText     =   "View Sales Calculator"
         Top             =   2245
         Width           =   1995
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F8 REMINDERS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":11C3
         MousePointer    =   99  'Custom
         TabIndex        =   62
         ToolTipText     =   "View Sales Calculator"
         Top             =   2951
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F5 PROSPECT LOG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":1315
         MousePointer    =   99  'Custom
         TabIndex        =   61
         ToolTipText     =   "View Sales Calculator"
         Top             =   1892
         Width           =   1995
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F4 CUSTOMER LOG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":1467
         MousePointer    =   99  'Custom
         TabIndex        =   60
         ToolTipText     =   "View Sales Calculator"
         Top             =   1539
         Width           =   1995
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Log Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":15B9
         MousePointer    =   99  'Custom
         TabIndex        =   45
         ToolTipText     =   "View Customer"
         Top             =   3304
         Width           =   1995
      End
      Begin VB.CommandButton Command7 
         Caption         =   "F3 CALCULATOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":170B
         MousePointer    =   99  'Custom
         TabIndex        =   1
         ToolTipText     =   "View Sales Calculator"
         Top             =   1186
         Width           =   1995
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F2 SALES ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":185D
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "View Sales Order"
         Top             =   833
         Width           =   1995
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F1  PROSPECT "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MouseIcon       =   "MainSAE.frx":19AF
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Add NewProspects or Edit Your Prospect Only.."
         Top             =   480
         Width           =   1995
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   3
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   2100
         _Version        =   655364
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   ":: TASK MENU  :::"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   64
      End
   End
   Begin VB.Label labVCount 
      Caption         =   "Label5"
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
      Height          =   375
      Left            =   2400
      TabIndex        =   105
      Top             =   7380
      Width           =   11265
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prospect Logs"
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
      Height          =   345
      Left            =   3420
      TabIndex        =   50
      Top             =   2010
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label66 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quotation"
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
      Height          =   345
      Left            =   3510
      TabIndex        =   49
      Top             =   4860
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reminders"
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
      Height          =   345
      Left            =   3480
      TabIndex        =   48
      Top             =   3420
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label59 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Logs"
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
      Height          =   345
      Left            =   3450
      TabIndex        =   47
      Top             =   2730
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Calculator"
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
      Height          =   345
      Left            =   3480
      TabIndex        =   46
      Top             =   4050
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Menu mnuContextProspect 
      Caption         =   "CONTEXT PROSPECT"
      Begin VB.Menu mnuTestDrive 
         Caption         =   "Schedule Test Drive"
      End
      Begin VB.Menu mnuSalesAppointment 
         Caption         =   "Add Sales Appoinment"
      End
      Begin VB.Menu mnuSendQuotation 
         Caption         =   "Add Quotation"
      End
      Begin VB.Menu mnuLoanApplication 
         Caption         =   "Add Loan Application"
      End
      Begin VB.Menu mnuSalesOrder 
         Caption         =   "Make Sales Order"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVisits 
         Caption         =   "Log Visits"
      End
      Begin VB.Menu mnuCalls 
         Caption         =   "Log Calls"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "Log Email"
      End
      Begin VB.Menu mnuLetter 
         Caption         =   "Log A Letter"
      End
      Begin VB.Menu mnuprospect 
         Caption         =   "Update Prospect"
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "View Prospect Inquiry"
      End
   End
   Begin VB.Menu mnuContextReminder 
      Caption         =   "ContextReminder"
      Begin VB.Menu mnuRemind_Cust 
         Caption         =   "New Customer Reminder"
      End
      Begin VB.Menu mnuRemind_Prospect 
         Caption         =   "New Prospect Reminder"
      End
      Begin VB.Menu mnuRemind_Edit 
         Caption         =   "Edit Reminder"
      End
      Begin VB.Menu mnuUpdateReminder 
         Caption         =   "Update Reminder"
      End
   End
End
Attribute VB_Name = "MainSAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PSTATUS                                                           As String
Dim SSTATUS                                                           As String
Dim lStatus                                                           As String
Dim VSTATUS                                                           As String
Dim RSTATUS                                                           As String

Dim Exist_SO                                                          As Boolean
Dim EXIST_LOAN                                                        As Boolean
Dim EXIST_INVOICE                                                     As Boolean
Dim CustomerCode                                                      As String
Dim WithEvents CustomerInformation                                    As frmAllCustomer
Attribute CustomerInformation.VB_VarHelpID = -1
Dim ProspID                                                           As Long
Dim ProspType                                                         As String
Dim LABDESC                                                           As String

Sub FillAndConfigGrid()
    With cboPriority
        .AddItem "Normal"
        .AddItem "High"
        .AddItem "Low"
        .AddItem "(ANY)"
    End With
    With cbostatus
        .AddItem "Not Started"
        .AddItem "In Progress"
        .AddItem "Completed"
        .AddItem "Waiting"
        .AddItem "Deferred"
        .AddItem "(ANY)"
    End With
    With lstMonth
        .AddItem ("January")
        .AddItem ("February")
        .AddItem ("March")
        .AddItem ("April")
        .AddItem ("May")
        .AddItem ("June")
        .AddItem ("July")
        .AddItem ("August")
        .AddItem ("September")
        .AddItem ("October")
        .AddItem ("November")
        .AddItem ("December")
        .AddItem ("All Months")
    End With
    With Combo1
        .AddItem "(ANY)"
        .AddItem "APPOINTMENT"
        .AddItem "CALLS"
        .AddItem "EMAIL"
        .AddItem "INITIAL INQUIRY"
        .AddItem "LETTER"
        .AddItem "LOAN APPLICATION"
        .AddItem "QUOTATION"
        .AddItem "SALES INVOICE CANCELLED"
        .AddItem "SALES INVOICE"
        .AddItem "SALES ORDER CANCELLED"
        .AddItem "SALES ORDER"
        .AddItem "TEST DRIVE"
        .AddItem "VISITS"
    End With
    fillcbomoreyear cboYear

    ReportControlPaintManager lstProspects
    ReportControlPaintManager lstLoanApplication
    ReportControlPaintManager lstSalesOrder
    ReportControlPaintManager lstVehicles
    ReportControlPaintManager lstReminders
    ReportControlPaintManager lstActivity
    ReportControlAddColumnHeader lstProspects, "Date,Prospect Name,Model,Description,[C],SAE,LeadSource"
    ReportControlAddColumnHeader lstSalesOrder, "Date,Customer Name,Model,Model Description,SAE, Status"
    ReportControlAddColumnHeader lstLoanApplication, "Date,Account Name,Model,Type, Date Of Status , Status"
    ReportControlAddColumnHeader lstReminders, "Type, Date , Due By , Name,Reminder Type ,Subject,Priority, Status"

    ResizeColumnHeader lstProspects, "10,20,12,20,8,12,8"
    ResizeColumnHeader lstSalesOrder, "8,18,8,20,20,10"
    ResizeColumnHeader lstLoanApplication, "10,20,20,8,20,8"
    ResizeColumnHeader lstReminders, "0,10,5,25,20,20,10,10"
    lstReminders.GroupsOrder.Add lstReminders.Columns(0)
    lstReminders.Columns(0).Visible = False

    ReportControlAddColumnHeader lstActivity, "DATE,PROSPECTNAME ,DETAILS, PARTICULAR,ISCLIENT"
    ResizeColumnHeader lstActivity, "10,30,20,40 ,10 "



    lstMonth.Selected(Month(LOGDATE) - 1) = True

End Sub

Sub FillLoanApplication(xDate As String)
    Dim RsUploadData                                                  As ADODB.Recordset
    Dim SQL                                                           As String
    Set RsUploadData = New ADODB.Recordset
    SQL = "Select "
    SQL = SQL & "convert(varchar, DateApplied ,101),  " & vbCrLf
    SQL = SQL & "isnull(Ind_Apl_LastName,'')  + ' . ' + isnull(Ind_Apl_FirstName, ''), " & vbCrLf
    SQL = SQL & "Ind_LoanApl_UnitModel , " & vbCrLf
    SQL = SQL & "'Individual' ," & vbCrLf
    SQL = SQL & "LASTUPDATED , " & vbCrLf
    SQL = SQL & "Case LStatus  " & vbCrLf
    SQL = SQL & "WHEN 'O' THEN 'On Process' " & vbCrLf
    SQL = SQL & "WHEN 'P' THEN 'Pending' " & vbCrLf
    SQL = SQL & "WHEN 'D' THEN 'Disapproved' " & vbCrLf
    SQL = SQL & "WHEN 'C' THEN 'Cancelled' " & vbCrLf
    SQL = SQL & "WHEN 'A' THEN 'Approved' END as Status ," & vbCrLf
    SQL = SQL & "Apl_no  , ID " & vbCrLf
    SQL = SQL & "from SMIS_LoanIndiv " & vbCrLf
    SQL = SQL & " WHERE USERCODE='" & LOGSAE & "' " & xDate & vbCrLf
    SQL = SQL & "UNION " & vbCrLf
    SQL = SQL & "Select " & vbCrLf
    SQL = SQL & "convert(varchar, DateApplied ,101),  " & vbCrLf
    SQL = SQL & "Busname, " & vbCrLf
    SQL = SQL & "unitmodel, " & vbCrLf
    SQL = SQL & "'Corporate' ," & vbCrLf
    SQL = SQL & "LASTUPDATED , " & vbCrLf
    SQL = SQL & "Case LStatus  " & vbCrLf
    SQL = SQL & "WHEN 'O' THEN 'On Process' " & vbCrLf
    SQL = SQL & "WHEN 'P' THEN 'Pending' " & vbCrLf
    SQL = SQL & "WHEN 'D' THEN 'Disapproved' " & vbCrLf
    SQL = SQL & "WHEN 'C' THEN 'Cancelled' " & vbCrLf
    SQL = SQL & "WHEN 'A' THEN 'Approved' END , " & vbCrLf
    SQL = SQL & "Aplcode , ID" & vbCrLf
    SQL = SQL & "from SMIS_LoanCORP " & vbCrLf
    SQL = SQL & " WHERE USERCODE='" & LOGSAE & "'" & xDate & vbCrLf

    SQL = SQL & "Order By 1 DESC" & vbCrLf

    Set RsUploadData = gconDMIS.Execute(SQL)
    flex_FillReportView RsUploadData, lstLoanApplication


End Sub

Sub FillLogInquiry(xquery)
    flex_FillReportView gconDMIS.Execute("Select convert(varchar, deyt, 101) , ACCTNAME , LogName, Particular , case ISNULL(CSCDE,'')  when   '' then 'NO' ELSE 'YES' END , LOGID  from CRIS_VIEWLOG INNER JOIN CRIS_PROSPECTS ON CRIS_PROSPECTS.PROSPECTID= CRIS_VIEWLOG.PROSPECTID  WHERE USERCODE='" & LOGSAE & "' and " & xquery & "   Order By CRIS_VIEWLOG.PROSPECTID ,[deyt] , SN "), lstActivity
End Sub

Sub FillProspect(xDate As String)
    Dim RsUploadData                                                  As ADODB.Recordset
    Set RsUploadData = New ADODB.Recordset
    lstProspects.FilterText = vbNullString
    Set RsUploadData = gconDMIS.Execute("SELECT  LogInitialInquiry , AcctName , Model,Variant, Classification , SAE , LeadSource  , ProspectType, ProspectID ,CUSCDE,LOGQUOTE,LOGSO  FROM CRIS_Prospects Where usercode='" & LOGSAE & "' and " & xDate & PSTATUS & " order by LOGINITIALINQUIRY DESC")
    flex_FillReportView RsUploadData, lstProspects
End Sub

Sub FillReminder(xDate As String)
    Dim SQL                                                           As String
    SQL = " SELECT "
    SQL = SQL & " Case ENTITYTYPE"
    SQL = SQL & " WHEN 'P' THEN 'PROSPECT'"
    SQL = SQL & " WHEN 'C' THEN 'CUSTOMER'"
    SQL = SQL & " WHEN 'S' THEN 'INTERNAL'"
    SQL = SQL & " END ,"
    SQL = SQL & " CONVERT(VARCHAR,DATETIMEREMIND,101) AS DATE,"
    SQL = SQL & " DATEDIFF(DAY, nexttime,GETDATE()) , "
    SQL = SQL & " Case"
    SQL = SQL & " WHEN ENTITYTYPE ='P' THEN (SELECT ACCTNAME FROM CRIS_PROSPECTS WHERE CAST(PROSPECTID AS VARCHAR) =CAST(CSCDE AS VARCHAR))"
    SQL = SQL & " WHEN ENTITYTYPE ='C' THEN (SELECT ACCTNAME FROM ALL_CUSTOMER WHERE CUSCDE  =CSCDE)"
    SQL = SQL & " WHEN ENTITYTYPE ='S' AND LEFT(CSCDE,2)='PR' THEN (SELECT ACCTNAME FROM CRIS_PROSPECTS WHERE CAST(PROSPECTID AS VARCHAR) =CAST(RIGHT(CSCDE ,LEN(CSCDE)-3) AS VARCHAR))"
    SQL = SQL & " WHEN ENTITYTYPE ='S' AND LEFT(CSCDE,2)='CS' THEN (SELECT ACCTNAME FROM ALL_CUSTOMER WHERE CUSCDE  = RIGHT(CSCDE ,LEN(CSCDE)-3)) END ,"
    SQL = SQL & " REMINDERTYPE,SUBJECT,"
    SQL = SQL & " Case Priority"
    SQL = SQL & " WHEN 'H' THEN 'HIGH'"
    SQL = SQL & " WHEN 'L' THEN 'LOW'"
    SQL = SQL & " WHEN 'N' THEN 'NORMAL'"
    SQL = SQL & " END AS PRIORITY,"
    SQL = SQL & " Case status"
    SQL = SQL & " WHEN 'N' THEN 'NOT STARTED'"
    SQL = SQL & " WHEN 'I' THEN 'IN PROGRESS'"
    SQL = SQL & " WHEN 'C' THEN 'COMPLETED'"
    SQL = SQL & " WHEN 'W' THEN 'WAITING'"
    SQL = SQL & " WHEN 'D' THEN 'DEFERRED'"
    SQL = SQL & " Else 'N/A'"
    SQL = SQL & " END As status, ID"
    SQL = SQL & " From CRIS_REMINDERS "
    SQL = SQL & " WHERE ENTITYTYPE <>('E') AND USERCODE='" & LOGSAE & "' " & xDate & RSTATUS
    '
    Dim RsUploadData                                                  As ADODB.Recordset
    Set RsUploadData = New ADODB.Recordset
    lstProspects.FilterText = vbNullString
    Set RsUploadData = gconDMIS.Execute(SQL)
    flex_FillReportView RsUploadData, lstReminders


End Sub

Sub FillSalesOrder(xDate As String)
    Dim RsUploadData                                                  As ADODB.Recordset
    Set RsUploadData = New ADODB.Recordset
    Set RsUploadData = gconDMIS.Execute("SELECT * FROM SMIS_VW_INQSALESORDER  Where usercode='" & LOGSAE & "' and " & xDate & SSTATUS & " order by Deyt DESC")
    flex_FillReportView RsUploadData, lstSalesOrder
End Sub

Sub FillVehicles(xCondition As String, Optional ByVal xDateFilter As String)
    Dim RsUploadData                                                  As ADODB.Recordset
    Dim SQL                                                           As String
    DoEvents


    If xCondition = "A" Then
        ReportControlAddColumnHeader lstVehicles, "MODEL,SO#,DATE#A , CUSTOMERNAME, DESCRIPTION, CSNO,COLOR,AGING,SAE,TERM"
        Call ResizeColumnHeader(lstVehicles, "0,8,10,30,30,10,10,8,15,8")
        SQL = "SELECT upper(SMIS_MrrInv.MODEL)MODEL , SMIS_SalesOrder.so_no, "
        SQL = SQL & " SMIS_SalesOrder.deyt,"
        SQL = SQL & " SMIS_SalesOrder.CustName,"
        SQL = SQL & " SMIS_MrrInv.DESCRIPT,"
        SQL = SQL & " SMIS_MrrInv.ignkey,"
        SQL = SQL & " SMIS_MrrInv.color,"
        SQL = SQL & " DATEDIFF(Day, SMIS_SalesOrder.Deyt, GETDATE()) As Aging,"
        SQL = SQL & " SMIS_SalesOrder.salesae ,"
        SQL = SQL & " SMIS_SalesOrder.TERM "
        SQL = SQL & " FROM SMIS_MrrInv INNER JOIN "
        SQL = SQL & " SMIS_SalesOrder ON SMIS_MrrInv.ignkey = SMIS_SalesOrder.IGNKEY_NO"
        SQL = SQL & " WHERE " & xDateFilter & " (isnull(SMIS_SalesOrder.STATUS,'')  <>'C') and  SMIS_SalesOrder.SOSTATUS ='P' and SMIS_SalesOrder.VI_NO is null  AND (SMIS_MrrInv.iSTATUS='A') and SMIS_SalesOrder.usercode='" & LOGSAE & "'"
    ElseIf xCondition = "S" Then

        ReportControlAddColumnHeader lstVehicles, "MODEL,DATE#I, CUSTOMER NAME ,DESCRIPTION, CSNO, COLOR,SAE, TERM"
        Call ResizeColumnHeader(lstVehicles, "0,8,30,30,10,10,15,8")
        SQL = "SELECT upper(SMIS_PURCHAGREE.MODEL)MODEL, SMIS_PURCHAGREE.InvoicedDate,"
        SQL = SQL & " SMIS_PURCHAGREE.CustName,"
        SQL = SQL & " SMIS_PURCHAGREE.modeldescription,"
        SQL = SQL & " SMIS_PURCHAGREE.ignkey_no,"
        SQL = SQL & " SMIS_PURCHAGREE.color ,"
        SQL = SQL & " SMIS_PURCHAGREE.salesae ,"
        SQL = SQL & " SMIS_PURCHAGREE.TERM"
        SQL = SQL & " FROM SMIS_PURCHAGREE "
        SQL = SQL & " "
        SQL = SQL & " WHERE  " & xDateFilter & "  SMIS_PURCHAGREE.VI_NO is Not Null  order by 1 "

    ElseIf xCondition = "O" Then
        ReportControlAddColumnHeader lstVehicles, "MODEL, CSNO,DESCRIPTION, COLOR, AGE(REC), AGE(P/O)"
        Call ResizeColumnHeader(lstVehicles, "0,10,35,25,10,10")
        SQL = "SELECT  upper(SMIS_MrrInv.MODEL)MODEL, " _
            & " ignkey as CSNO,  " _
            & " DESCRIPT + ' ' + isnull(YEER ,'') as Descriptions," _
            & " Color, " _
            & " Datediff(day,DateReceived,getdate()) as Aging1, " _
            & " Datediff(day,PullOutDate,getdate()) as Aging2, " _
            & " ID FROM SMIS_MRRINV  WHERE RELEASED= 0 and status='P' order by 2 "
    ElseIf xCondition = "R" Then
        ReportControlAddColumnHeader lstVehicles, "MODEL,DATE REL, CUSTOMER NAME ,DESCRIPTION, CSNO, COLOR, TERM"
        Call ResizeColumnHeader(lstVehicles, "0,8,30,30,10,10,8")
        SQL = "SELECT upper(SMIS_PURCHAGREE.MODEL)MODEL, SMIS_PURCHAGREE.datereleased,"
        SQL = SQL & " SMIS_PURCHAGREE.CustName,"
        SQL = SQL & " SMIS_PURCHAGREE.modeldescription,"
        SQL = SQL & " SMIS_PURCHAGREE.ignkey_no,"
        SQL = SQL & " SMIS_PURCHAGREE.color ,"
        SQL = SQL & " SMIS_PURCHAGREE.TERM"
        SQL = SQL & " FROM SMIS_PURCHAGREE "
        SQL = SQL & " "
        SQL = SQL & " WHERE     " & xDateFilter & "   SMIS_PURCHAGREE.STATUS<>'C' order by 1 "
    End If
    lstVehicles.GroupsOrder.Add lstVehicles.Columns(0): lstVehicles.Columns(0).Visible = False
    Set RsUploadData = gconDMIS.Execute(SQL)
    flex_FillReportView RsUploadData, lstVehicles


End Sub

Sub ShowStatus(PROSPECTID)
    Dim temprs                                                        As ADODB.Recordset
    Dim xstatus                                                       As String

    Set temprs = gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID)

    EXIST_INVOICE = False: EXIST_LOAN = False: Exist_SO = False: lblLogQuote = "": lblLogEmail = "": lblLogAppointment = "": lblLogTestDrive = "": lblLogCalls = ""
    lblLogLetters = "": lblLetters = "": lblLogLoan = "": lblLogVisits = "": lblLogSalesOrder = "": lblSTATUS = ""
    txtCusAdd = "": txtCusEmail = "": txtCusName = "": txtCusPhone = "": lblNotes = "": captionInformation.Caption = ""
    CustomerCode = "": captionInformation.Caption = ""
    If Not (temprs.EOF Or temprs.BOF) Then
        lblLogQuote = "x": lblLogQuote.ForeColor = &HC0&
        lblLogEmail = "x": lblLogEmail.ForeColor = &HC0&
        lblLogAppointment = "x": lblLogAppointment.ForeColor = &HC0&
        lblLogTestDrive = "x": lblLogTestDrive.ForeColor = &HC0&
        lblLogCalls = "x": lblLogCalls.ForeColor = &HC0&
        lblLogLetters = "x": lblLogLetters.ForeColor = &HC0&
        lblLogLoan = "x": lblLogLoan.ForeColor = &HC0&
        lblSTATUS = "x": lblSTATUS.BackColor = vbWhite
        lblLogVisits = "x": lblLogVisits.ForeColor = &HC0&

        lblInvoiceNo = "x": lblInvoiceNo.ForeColor = &HC0&
        lblLogSalesOrder.Caption = "x": lblLogSalesOrder.ForeColor = &HC0&
        lblNotes.Text = ""
        ProspType = Null2String(temprs!ProspectType)
        lblNotes.Text = Null2String(temprs!Notes)

        If IsNull(temprs!LogQuote) = False Then
            lblLogQuote = Chr(187) & FormatDateTime(temprs!LogQuote, vbShortDate): lblLogQuote.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogEmail) = False Then
            lblLogEmail = Chr(187) & FormatDateTime(temprs!LogEmail, vbShortDate): lblLogEmail.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogAppointment) = False Then
            lblLogAppointment = Chr(187) & FormatDateTime(temprs!LogAppointment, vbShortDate): lblLogAppointment.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogTestDrive) = False Then
            lblLogTestDrive = Chr(187) & FormatDateTime(temprs!LogTestDrive, vbShortDate) & "(" & DateDiff("d", temprs!LogTestDrive, LOGDATE) & ")": lblLogTestDrive.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogCall) = False Then
            lblLogCalls = Chr(187) & FormatDateTime(temprs!LogCall, vbShortDate): lblLogCalls.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogLetter) = False Then
            lblLogLetters = Chr(187) & FormatDateTime(temprs!LogLetter, vbShortDate): lblLogLetters.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogApplication) = False Then
            EXIST_LOAN = True
            lblLogLoan = Chr(187) & FormatDateTime(temprs!LogApplication, vbShortDate): lblLogLoan.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogVisit) = False Then
            lblLogVisits = Chr(187) & FormatDateTime(temprs!LogVisit, vbShortDate): lblLogVisits.ForeColor = &H8000&
        End If

        If IsNull(temprs!INVOICENO) = False Then
            EXIST_INVOICE = True
            lblInvoiceNo = Chr(187) & temprs!INVOICENO: lblInvoiceNo.ForeColor = &H8000&
        End If



        If IsNull(temprs!LOGSO) = False Then
            Exist_SO = True
            lblLogSalesOrder = Chr(187) & FormatDateTime(temprs!LOGSO, vbShortDate) & "(" & Null2String(temprs!SO_NO) & ")": lblLogSalesOrder.ForeColor = &H8000&
            If IsDate(temprs!LOGCLOSINGDATE) = True Then
                lblAgeing = "Closed In: " & DateDiff("d", temprs!LOGCLOSINGDATE, temprs!loginitialinquiry) & " Days"
            End If

        Else
            lblAgeing = "Prospect Age: " & DateDiff("d", temprs!loginitialinquiry, LOGDATE) & " Days"
        End If
        xstatus = Null2String(temprs!STATUS)

        If xstatus = "O" Then
            lblSTATUS = "OPEN": lblSTATUS.BackColor = &HC000&
        ElseIf xstatus = "C" And EXIST_INVOICE Then
            lblSTATUS = "CLOSED": lblSTATUS.BackColor = &H40C0&
        ElseIf xstatus = "C" And EXIST_INVOICE = False Then
            lblSTATUS = "INVOICING": lblSTATUS.BackColor = &H40C0&
        ElseIf xstatus = "I" Then
            lblSTATUS = "INACTIVE": lblSTATUS.BackColor = &HC0C0C0
        ElseIf xstatus = "L" Then
            lblSTATUS = "LOST SALES": lblSTATUS.BackColor = &HC0C0C0
        Else
            lblSTATUS = "OPEN": lblSTATUS.BackColor = &HC000&
        End If
        If IsNull(temprs!CUSCDE) = False Then
            CustomerCode = temprs!CUSCDE: ShowCustomerInfo CustomerCode
        Else
            If IsNull(temprs!EMAIL) = False Then
                txtCusEmail = temprs!EMAIL
            End If

            If IsNull(temprs!Telephone) = False Then
                txtCusPhone = temprs!Telephone
            End If

            If IsNull(temprs!Address) = False Then
                txtCusAdd = temprs!Address
            End If

        End If
    End If
    Set temprs = Nothing
End Sub

Private Sub cboPriority_Click()
    ShowData
End Sub

Private Sub cboStatus_Click()
    ShowData
End Sub

Private Sub cboYear_Click()
    ShowData
End Sub

Private Sub cmdOther_SAEPerformance_1_Click()
    frmSMIS_Report_SAEPersonal.Show
End Sub

Private Sub cmdPrintSO_Click()

    On Error GoTo Errorcode:

    If lstSalesOrder.Records.Count = 0 Then
        MsgSpeechBox "No Records To print"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Dim MANTH, YEER, REPORTSQL
    MANTH = Month(DateSerial(cboYear, (lstMonth.ListIndex + 1), 1))
    YEER = cboYear

    If optSO(0).Value = True Then
        REPORTSQL = " {SO.STATUS}='ON ORDER'"
        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " AND MONTH({SO.DEYT})=" & MANTH
        End If
        REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
    End If

    If optSO(1).Value = True Then
        REPORTSQL = " {SO.STATUS}='CANCELLED INVOICE' OR {SO.STATUS}='CANCELLED SO'"

        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " AND MONTH({SO.DEYT})=" & MANTH
        End If
        REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
    End If

    If optSO(2).Value = True Then
        REPORTSQL = " {SO.STATUS}='INVOICED'"

        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " AND MONTH({SO.DEYT})=" & MANTH
        End If
        REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
    End If

    If optSO(3).Value = True Then
        REPORTSQL = " {SO.STATUS}='RELEASED'"

        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " AND MONTH({SO.DEYT})=" & MANTH
        End If
        REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
    End If

    If optSO(4).Value = True Then
        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " MONTH({SO.DEYT})=" & MANTH
            REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
        Else
            REPORTSQL = REPORTSQL & " YEAR({SO.DEYT})=" & YEER
        End If
    End If

    rptMain.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptMain.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptMain, SMIS_REPORT_PATH & "SOListing.rpt", REPORTSQL & " AND {SO.USERCODE}='" & LOGSAE & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = vbDefault
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrintVehicles_Click()
    On Error GoTo Errorcode:

    If lstVehicles.Records.Count = 0 Then
        MsgSpeechBox "No Records To print"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lstVehicles.PrintOptions.Header.Font.Size = 8
    lstVehicles.PrintOptions.Header.Font.Bold = True
    If optInventory(0).Value = True Then
        lstVehicles.PrintOptions.Header.TextCenter = UCase(COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & "ALLOCATED VEHICLES")
    ElseIf optInventory(1).Value = True Then
        lstVehicles.PrintOptions.Header.TextCenter = UCase(COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & "INVOICED VEHICLES")
    ElseIf optInventory(2).Value = True Then
        lstVehicles.PrintOptions.Header.TextCenter = UCase(COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & "RELEASED VEHICLES")
    ElseIf optInventory(3).Value = True Then
        lstVehicles.PrintOptions.Header.TextCenter = UCase(COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & "ON STOCK VEHICLES")
    End If
    lstVehicles.PrintPreview True



    Screen.MousePointer = vbNormal
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Combo1_Click()
    ShowData
End Sub

Private Sub Command1_Click()
    frmSMIS_Files_Prospects.Show
End Sub

Private Sub Command10_Click()
    frmSMIS_FilesAccMaintenance.Show
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    frmSMIS_Trans_SalesOrder.Show


End Sub

Private Sub Command3_Click()
    If MsgBox(" Logging Off?", vbYesNo + vbExclamation) = vbYes Then
        LOGSAE = ""
        UnloadForm Me
    End If
End Sub

Private Sub Command4_Click()
    frmSMIS_Log_Menu.picLogCustomer.Visible = True
    frmSMIS_Log_Menu.picLogProspect.Visible = False
    frmSMIS_Log_Menu.Show
    frmSMIS_Log_Menu.ZOrder 0
End Sub

Private Sub Command5_Click()
    frmSMIS_Log_Menu.picLogCustomer.Visible = False
    frmSMIS_Log_Menu.picLogProspect.Visible = True
    frmSMIS_Log_Menu.Show
    frmSMIS_Log_Menu.ZOrder 0

End Sub

Private Sub Command6_Click()
    mnuUpdateReminder.Enabled = False
    mnuRemind_Edit.Enabled = False
    PopupMenu mnuContextReminder
End Sub

Private Sub Command7_Click()
    frmSMIS_Mis_AOR.Show
End Sub

Private Sub Command8_Click()
    On Error Resume Next
    frmSMIS_Trans_ApplicationIndividual.Show
    If FormExist("frmSMIS_Trans_ApplicationIndividual") Then
        frmSMIS_Trans_ApplicationIndividual.ZOrder 0
    End If

End Sub

Private Sub Command9_Click()
    On Error Resume Next
    frmSMIS_Trans_ApplicationCorporate.Show
    If FormExist("frmSMIS_Trans_ApplicationCorporate") Then
        frmSMIS_Trans_ApplicationCorporate.ZOrder 0
    End If

End Sub

Private Sub CustomerInformation_ProspectConverted(CustomerCode As String, xGoingWhere As String, PROSPECTID As Long)
    On Error GoTo Errorcode:
    ShowStatus PROSPECTID
    gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=" & N2Str2Null(CustomerCode) & " where  Prospectid=" & PROSPECTID)
    Call frmSMIS_Trans_SalesOrder.AddNewSOFromProspect(gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID))
    Unload CustomerInformation
    On Error Resume Next
    frmSMIS_Trans_SalesOrder.Show
    If FormExist("frmSMIS_Trans_SalesOrder") Then
        On Error Resume Next
        frmSMIS_Trans_SalesOrder.SetFocus
        frmSMIS_Trans_SalesOrder.ZOrder 0
    End If
    Exit Sub
    If xGoingWhere = "Order" Then
        ShowStatus PROSPECTID
        gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=" & N2Str2Null(CustomerCode) & " where  Prospectid=" & PROSPECTID)
        Call frmSMIS_Trans_SalesOrder.AddNewSOFromProspect(gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID))
        Unload CustomerInformation
        frmSMIS_Trans_SalesOrder.Show
    End If
    If xGoingWhere = "Prospects" Then
        ShowStatus PROSPECTID
        gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=" & N2Str2Null(CustomerCode) & " where  Prospectid=" & PROSPECTID)
        Call Unload(CustomerInformation)
        Set CustomerInformation = Nothing
    End If
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            Command1_Click
        Case vbKeyF2
            Command2_Click
        Case vbKeyF3
            Command7_Click
        Case vbKeyF4
            Command4_Click
        Case vbKeyF5
            Command5_Click
        Case vbKeyF6
            Command8_Click
        Case vbKeyF7
            Command9_Click
        Case vbKeyF8

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1

    FillAndConfigGrid
    With Me
        .StatusBar1.Panels(1).Text = "AGENT CODE: " & LOGSAE
        .StatusBar1.Panels(2).Text = "AGENT NAME: " & SAENAME
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
    End With
    cboYear.Text = Format(LOGDATE, "yyyy")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox(" Logging Off?", vbYesNo + vbExclamation) = vbYes Then
        LOGSAE = ""
    Else
        Cancel = 1
    End If
End Sub

Private Sub lstLoanApplication_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If UCase(Row.Record(4).Value) = "CORPORATE" Then
        frmSMIS_Trans_ApplicationCorporate.Show
        frmSMIS_Trans_ApplicationCorporate.SearchID (Row.Record(7).Value)
    Else
        frmSMIS_Trans_ApplicationIndividual.Show
        frmSMIS_Trans_ApplicationIndividual.SearchID (Row.Record(7).Value)
    End If
End Sub

Private Sub lstMonth_Click()
    ShowData
End Sub

Private Sub lstProspects_GotFocus()
    lstProspects_SelectionChanged
End Sub

Private Sub lstProspects_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then: Exit Sub
    frmSMIS_Files_Prospects.EditProspect (Row.Record(8).Value)
    frmSMIS_Files_Prospects.Show
End Sub

Private Sub lstProspects_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    lstProspects_SelectionChanged
    If Exist_SO = True Then
        mnuSalesOrder.Caption = "Open Sales Order"
    Else
        mnuSalesOrder.Caption = "Add New Sales Order"
    End If
    If EXIST_LOAN = True Then
        mnuLoanApplication.Caption = "Open Loan Application"
    Else
        mnuLoanApplication.Caption = "Add Loan Application"
    End If
    PopupMenu mnuContextProspect
End Sub

Private Sub lstProspects_SelectionChanged()
    If lstProspects.Records.Count = 0 Then
        ShowStatus 0
        Exit Sub
    End If
    ProspID = lstProspects.SelectedRows(0).Record(8).Value
    ShowStatus ProspID
End Sub

Private Sub lstReminders_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then Exit Sub
    frmSMIS_Log_ReminderStatus.REMINDERID = Row.Record(8).Value
    frmSMIS_Log_ReminderStatus.Show
End Sub

Private Sub lstReminders_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then: Exit Sub
    mnuUpdateReminder.Enabled = True
    mnuRemind_Edit.Enabled = True
    PopupMenu mnuContextReminder
End Sub

Private Sub lstSalesOrder_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then Exit Sub
    Load frmSMIS_Trans_SalesOrder
    Call frmSMIS_Trans_SalesOrder.SearchID(Row.Record(7).Value)
End Sub

Private Sub mnuCalls_Click()
    Call frmSMIS_Log_Call.AddCall(ProspID, vbNullString)
    frmSMIS_Log_Call.Show
    frmSMIS_Log_Call.ZOrder 0
End Sub

Private Sub mnuEmail_Click()
    Call frmSMIS_Log_Email.AddEmail(ProspID, vbNullString)
    frmSMIS_Log_Email.Show
    frmSMIS_Log_Email.ZOrder 0
End Sub

Private Sub mnuLetter_Click()
    Call frmSMIS_Log_Letter.AddLetter(ProspID, vbNullString)
    frmSMIS_Log_Letter.Show
    frmSMIS_Log_Letter.ZOrder 0
End Sub

Private Sub mnuLoanApplication_Click()


    Dim temprs                                                        As ADODB.Recordset
    Dim rsExists                                                      As ADODB.Recordset

    On Error Resume Next

    Set temprs = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE PROSPECTID= " & ProspID)

    If ProspType = "P" Then

        If EXIST_LOAN = False Then
LINDV:             frmSMIS_Trans_ApplicationIndividual.AddFromProspects ProspID
            frmSMIS_Trans_ApplicationIndividual.Show
            frmSMIS_Trans_ApplicationIndividual.ZOrder 0
        Else
            Set rsExists = gconDMIS.Execute("select ID from SMIS_LoanINDIV where ProspectID=" & ProspID)
            If Not rsExists.EOF Or Not rsExists.BOF Then
                frmSMIS_Trans_ApplicationIndividual.Show
                frmSMIS_Trans_ApplicationIndividual.SearchID rsExists("ID").Value
                frmSMIS_Trans_ApplicationIndividual.ZOrder 0
            Else
                GoTo LINDV:
            End If
        End If

    Else
        If EXIST_LOAN = False Then
CORPDIV:             frmSMIS_Trans_ApplicationCorporate.AddFromProspects ProspID
            frmSMIS_Trans_ApplicationCorporate.Show
            frmSMIS_Trans_ApplicationCorporate.ZOrder 0
        Else
            Set rsExists = gconDMIS.Execute("select ProspectID from SMIS_LoanINDIV where ProspectID=" & ProspID)
            If Not rsExists.EOF Or Not rsExists.BOF Then
                frmSMIS_Trans_ApplicationCorporate.Show
                frmSMIS_Trans_ApplicationCorporate.SearchID rsExists("ID").Value
                frmSMIS_Trans_ApplicationCorporate.ZOrder 0
            Else
                GoTo CORPDIV:
            End If
        End If

    End If

End Sub

Private Sub mnuprospect_Click()
    frmSMIS_Files_ProspectStatus.PROSPECTID = lstProspects.SelectedRows(0).Record(8).Value
    frmSMIS_Files_ProspectStatus.Show
End Sub

Private Sub mnuRemind_Cust_Click()
    frmSMIS_Log_CustomerReminder.Show
    frmSMIS_Log_CustomerReminder.cmdAdd.Value = True
End Sub

Private Sub mnuRemind_Edit_Click()
    If lstReminders.SelectedRows(0).GroupRow = True Then Exit Sub
    If lstReminders.SelectedRows(0).Record(0).Value = "CUSTOMER" Then
        frmSMIS_Log_CustomerReminder.EditReminder "C", (lstReminders.SelectedRows(0).Record(8).Value)
        frmSMIS_Log_CustomerReminder.Show
    ElseIf lstReminders.SelectedRows(0).Record(0).Value = "PROSPECT" Then
        frmSMIS_Log_ProspectReminder.EditReminder (lstReminders.SelectedRows(0).Record(8).Value)
        frmSMIS_Log_ProspectReminder.Show

    End If
End Sub

Private Sub mnuRemind_Prospect_Click()
    frmSMIS_Log_ProspectReminder.Show
    frmSMIS_Log_ProspectReminder.cmdAdd.Value = True
End Sub

Private Sub mnuSalesAppointment_Click()
    frmSMIS_Log_SalesAppointment.AddSalesAppointment (ProspID)
    frmSMIS_Log_SalesAppointment.Show
    frmSMIS_Log_SalesAppointment.ZOrder 0
End Sub

Private Sub mnuSalesOrder_Click()
    Dim temprs                                                        As ADODB.Recordset
    Dim rsExists                                                      As ADODB.Recordset

    Set temprs = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE PROSPECTID= " & ProspID)
    If Null2String(temprs!CUSCDE) = "" Then
        Call MsgBox("Current Prospect Has Not been Converted to Customer !" & vbCrLf & " Please Convert It to Customer First ! Contact Sales Admin", vbExclamation Or vbDefaultButton1, App.TITLE)
        Exit Sub
    End If


    Set rsExists = gconDMIS.Execute("select ID from SMIS_SALESORDER WHERE sostatus<>'C' and CODE='" & CustomerCode & "'")
    If Not (rsExists.EOF Or rsExists.BOF) And Exist_SO = True Then

        frmSMIS_Trans_SalesOrder.Show
        frmSMIS_Trans_SalesOrder.SearchID rsExists!ID
        frmSMIS_Trans_SalesOrder.ZOrder 0
    Else
        If frmSMIS_Trans_SalesOrder.AddNewSOFromProspect(temprs) = True Then
            frmSMIS_Trans_SalesOrder.Show
            frmSMIS_Trans_SalesOrder.ZOrder 0
        End If
    End If
End Sub

Private Sub mnuSendQuotation_Click()
    On Error Resume Next
    Call frmSMIS_Trans_Quotation.AddNewQuotation(ProspID)
    frmSMIS_Trans_Quotation.Show
    frmSMIS_Trans_Quotation.ZOrder 0
    frmSMIS_Trans_Quotation.cboModel_Change
End Sub

Private Sub mnuTestDrive_Click()
    On Error Resume Next

    frmSMIS_Log_TestDriveAppointment.AddTestDriveAppointment (ProspID)
    frmSMIS_Log_TestDriveAppointment.Show
    frmSMIS_Log_TestDriveAppointment.ZOrder 0
End Sub

Private Sub mnuUpdateReminder_Click()
    If lstReminders.SelectedRows(0).GroupRow = True Then Exit Sub
    frmSMIS_Log_ReminderStatus.REMINDERID = lstReminders.SelectedRows(0).Record(8).Value
    frmSMIS_Log_ReminderStatus.Show
End Sub

Private Sub mnuViewLog_Click()
    frmSMIS_Inquiry_ViewLog.SHOWPROSPECTLOG lstProspects.SelectedRows.Row(0).Record(8).Value, lstProspects.SelectedRows.Row(0).Record(1).Value
End Sub

Private Sub mnuVisits_Click()
    Call frmSMIS_Log_Visits.AddVisit(ProspID, vbNullString)
    frmSMIS_Log_Visits.Show
    frmSMIS_Log_Visits.ZOrder 0
End Sub

Private Sub optInventory_Click(Index As Integer)
    ShowData
End Sub

Private Sub optLoan_Click(Index As Integer)
    ShowData
End Sub

Private Sub optProspects_Click(Index As Integer)
    ShowData
End Sub

Private Sub optReminder_Click(Index As Integer)
    ShowData
End Sub

Private Sub optSO_Click(Index As Integer)
    ShowData
End Sub

Private Sub ShowCustomerInfo(xxxcode)
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute("Select * from ALL_CUSTOMER WHERE CUSCDE=" & N2Str2Null(xxxcode))
    If Not (temprs.EOF Or temprs.BOF) Then
        If temprs.Fields("CUSTYPE") = "P" Then
            txtCusAdd = Replace(Null2String(temprs("CUSTOMERADD")), Chr(10), "")
            txtCusEmail = Null2String(temprs("EMAIL"))
            txtCusName = Null2String(temprs("LASTNAME")) & IIf(IsNull(temprs("LASTNAME")), "", ",") & Null2String(temprs("FirstName")) & IIf(IsNull(temprs("MIDDLEINITIAL")), "", ".") & Null2String(temprs("MIDDLEINITIAL"))
            txtCusPhone = Null2String(temprs("HOMEPHONE"))
            captionInformation.Caption = Null2String(temprs("ACCTNAME"))

        Else
            txtCusAdd = Replace(Null2String(temprs("COMPANYADD")), Chr(10), "")
            txtCusEmail = Null2String(temprs("EMAIL"))
            txtCusName = Null2String(temprs("CUSCOMP"))
            txtCusPhone = Null2String(temprs("TELEPHONENO"))
            captionInformation.Caption = Null2String(temprs("ACCTNAME"))

        End If
    End If
    Set temprs = Nothing
End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Index = 0 Then                                   'PROSPECT
        txtSearch_Prospect.SetFocus
    ElseIf Item.Index = 1 Then                               'SALES ORDER
        txtSearch_SalesOrder.SetFocus
    ElseIf Item.Index = 2 Then                               'LOAN
        txtSearch_LoanApplication.SetFocus
    ElseIf Item.Index = 3 Then                               'VEHICLE
        txtSearch_Vehicles.SetFocus

    ElseIf Item.Index = 4 Then                               'REPORTS
        txtSearch_Activity.SetFocus
        'ElseIf Item.Index = 5 Then                              ' REMINDERS
    End If
    DoEvents
    ShowData
End Sub

Private Sub txtSearch_Activity_Change()
    lstActivity.FilterText = txtSearch_Activity
    lstActivity.Populate
End Sub

Private Sub txtSearch_Activity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstActivity.Records.Count > 0 Then
            lstActivity.SelectedRows(0).Selected = True
            lstActivity.SetFocus
        End If
    End If
End Sub

Private Sub txtSearch_LoanApplication_Change()
    lstLoanApplication.FilterText = txtSearch_LoanApplication
    lstLoanApplication.Populate
End Sub

Private Sub txtSearch_LoanApplication_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstLoanApplication.Records.Count > 0 Then
            lstLoanApplication.SelectedRows(0).Selected = True
            lstLoanApplication.SetFocus
        End If
    End If
End Sub

Private Sub txtSearch_Prospect_Change()
    lstProspects.FilterText = txtSearch_Prospect.Text
    lstProspects.Populate
    ShowStatus 0
End Sub

Private Sub txtSearch_Prospect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstProspects.Records.Count > 0 Then
            lstProspects.SelectedRows(0).Selected = True
            lstProspects.SetFocus
        End If
    End If
End Sub

Private Sub txtSearch_SalesOrder_Change()
    lstSalesOrder.FilterText = txtSearch_SalesOrder
    lstSalesOrder.Populate
End Sub

Private Sub txtSearch_SalesOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstSalesOrder.Records.Count > 0 Then
            lstSalesOrder.SelectedRows(0).Selected = True
            lstSalesOrder.SetFocus
        End If
    End If
End Sub

Private Sub txtSearch_Vehicles_Change()
    lstVehicles.FilterText = txtSearch_Vehicles.Text
    lstVehicles.Populate

End Sub

Private Sub txtSearch_Vehicles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstVehicles.Records.Count > 0 Then
            lstVehicles.SelectedRows(0).Selected = True
            lstVehicles.SetFocus
        End If
    End If

End Sub

Public Sub ShowData()

    Dim Indx                                                          As Integer
    Dim Priority                                                      As String
    Dim xstatus                                                       As String
    Dim XLOGName                                                      As String
    Select Case lstMonth.Text
        Case "January": Indx = 1
        Case "February": Indx = 2
        Case "March": Indx = 3
        Case "April": Indx = 4
        Case "May": Indx = 5
        Case "June": Indx = 6
        Case "July": Indx = 7
        Case "August": Indx = 8
        Case "September": Indx = 9
        Case "October": Indx = 10
        Case "November": Indx = 11
        Case "December": Indx = 12
        Case Else: Indx = -1
    End Select



    Select Case TabControl.SelectedItem
        Case 0                                               'PROSPECT
            mnuCalls.Enabled = True
            mnuEmail.Enabled = True
            mnuLetter.Enabled = True
            mnuLoanApplication.Enabled = True
            mnuSalesAppointment.Enabled = True
            mnuSalesOrder.Enabled = True
            mnuSendQuotation.Enabled = True
            mnuTestDrive.Enabled = True
            mnuVisits.Enabled = True

            If optProspects(0).Value = True Then
                PSTATUS = " AND STATUS='O' "
                LABDESC = " SHOWING OPEN PROSPECT "
            ElseIf optProspects(1).Value = True Then
                PSTATUS = " AND STATUS='C' "
                mnuLoanApplication.Enabled = False
                mnuSalesAppointment.Enabled = False
                mnuSalesOrder.Enabled = False
                mnuSendQuotation.Enabled = False
                mnuTestDrive.Enabled = False
                LABDESC = "SHOWING CLOSED PROSPECT "
            ElseIf optProspects(2).Value = True Then
                PSTATUS = " AND STATUS='I' "
                mnuLoanApplication.Enabled = False
                mnuSalesAppointment.Enabled = False
                mnuSalesOrder.Enabled = False
                mnuSendQuotation.Enabled = False
                mnuTestDrive.Enabled = False
                LABDESC = "SHOWING INACTIVE PROSPECT "
            ElseIf optProspects(3).Value = True Then
                PSTATUS = " AND LOGFOLLOWUPDATE IS NOT NULL AND LOGFOLLOWUPDATE >= CONVERT(varchar,GETDATE(),101) "
                LABDESC = "SHOWING PROSPECTS FOR FOLLOW UP "
            Else
                PSTATUS = " "
                LABDESC = " ALL PROSPECTS "
            End If




            If Indx = -1 Then
                FillProspect "( YEAR(LOGCLOSINGDATE)=" & cboYear.Text & " OR  YEAR(LOGINITIALINQUIRY)=" & cboYear.Text & ")"
                labVCount = LABDESC & "(" & lstProspects.Records.Count & ") For the " & cboYear
            Else
                FillProspect " (YEAR(LOGCLOSINGDATE)=" & cboYear.Text & "  AND MONTH(LOGCLOSINGDATE)  =" & Indx & " OR  YEAR(LOGINITIALINQUIRY)=" & cboYear.Text & "  AND MONTH(LOGINITIALINQUIRY)  =" & Indx & ")"
                labVCount = LABDESC & "(" & lstProspects.Records.Count & ") For the  " & lstMonth & " " & cboYear
            End If
        Case 1                                               'SALES ORDER
            If optSO(0).Value = True Then
                SSTATUS = " AND (Status='ON ORDER') "
                LABDESC = " SALES ORDER ON ORDER "
            ElseIf optSO(1).Value = True Then
                SSTATUS = "  AND  (STATUS='CANCELLED INVOICE' OR Status='CANCELLED SO')"
                LABDESC = " CANCELLED SALES ORDER/SALES INVOICE "
            ElseIf optSO(2).Value = True Then
                SSTATUS = " AND   STATUS='INVOICED' "
                LABDESC = " INVOICED SALES ORDER "
            ElseIf optSO(3).Value = True Then
                SSTATUS = " AND   STATUS='RELEASED' "
                LABDESC = " RELEASED SALES ORDER "
            Else
                SSTATUS = vbNullString
                LABDESC = " "
            End If

            If Indx = -1 Then
                FillSalesOrder " YEAR(DEYT)=" & cboYear.Text
                labVCount = LABDESC & "(" & lstSalesOrder.Records.Count & ") Sales Order For the " & cboYear
            Else
                FillSalesOrder " YEAR(DEYT)=" & cboYear.Text & "  and MOnth(deyt)=" & (Indx)
                labVCount = LABDESC & "(" & lstSalesOrder.Records.Count & ") Sales Order For the " & lstMonth & " " & cboYear
            End If
        Case 2                                               'LOAN APPLICATION
            If optLoan(0).Value = True Then
                lStatus = " AND LSTATUS='" & optLoan(0).Tag & "' "
                LABDESC = " PENDING LOAN APPLCATION "
            ElseIf optLoan(1).Value = True Then
                lStatus = " AND LSTATUS='" & optLoan(1).Tag & "' "
                LABDESC = " LOAN APPLCATION ON PROCESS"
            ElseIf optLoan(2).Value = True Then
                lStatus = " AND LSTATUS='" & optLoan(2).Tag & "' "
                LABDESC = " DISAPPROVED LOAN APPLCATION "
            ElseIf optLoan(3).Value = True Then
                lStatus = " AND LSTATUS='" & optLoan(3).Tag & "' "
                LABDESC = " APPROVED LOAN APPLCATION "
            ElseIf optLoan(4).Value = True Then
                lStatus = " AND LSTATUS='" & optLoan(4).Tag & "' "
                LABDESC = " CANCELLED LOAN APPLCATION "
            Else
                lStatus = " "
                LABDESC = " ALL LOAN APPLCATION "
            End If
            If Indx = -1 Then
                FillLoanApplication " AND YEAR(DateApplied)=" & cboYear.Text & lStatus
                labVCount = LABDESC & "(" & lstLoanApplication.Records.Count & ") Loan Applicaion  For the  " & cboYear
            Else
                FillLoanApplication " AND  YEAR(DateApplied)=" & cboYear.Text & " and MOnth(DateApplied)=" & (Indx) & lStatus
                labVCount = LABDESC & "(" & lstLoanApplication.Records.Count & ") Loan Applicaion  For the  " & lstMonth & " " & cboYear
            End If

        Case 3                                               'VEHICLES INQUIRY
            If optInventory(0).Value = True Then
                VSTATUS = "A"
                LABDESC = " ALLOCATED VEHICLES "
                If Indx = -1 Then
                    FillVehicles VSTATUS, " YEAR(SMIS_SALESORDER.Deyt)=" & cboYear & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For  the " & cboYear
                Else
                    FillVehicles VSTATUS, "YEAR(SMIS_SALESORDER.Deyt)=" & cboYear & "  and MONTH(SMIS_SALESORDER.Deyt)=" & Indx & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For the  " & lstMonth & " " & cboYear
                End If

            ElseIf optInventory(1).Value = True Then
                VSTATUS = "S"
                LABDESC = " VEHICLES INVOICED"
                If Indx = -1 Then
                    FillVehicles VSTATUS, " YEAR(SMIS_PURCHAGREE.INVOICEDDATE)=" & cboYear & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") FOR  THE " & cboYear
                Else
                    FillVehicles VSTATUS, "YEAR(SMIS_PURCHAGREE.INVOICEDDATE)=" & cboYear & "  AND MONTH(SMIS_PURCHAGREE.INVOICEDDATE)=" & Indx & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") FOR THE  " & lstMonth & " " & cboYear
                End If

            ElseIf optInventory(2).Value = True Then
                VSTATUS = "R"
                LABDESC = " TOTAL VEHICLES ON RELEASED "
                If Indx = -1 Then
                    FillVehicles VSTATUS, " YEAR(SMIS_PURCHAGREE.DateReleased)=" & cboYear & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For  the " & cboYear
                Else
                    FillVehicles VSTATUS, "YEAR(SMIS_PURCHAGREE.DateReleased)=" & cboYear & "  and MONTH(SMIS_PURCHAGREE.DateReleased)=" & Indx & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For the  " & lstMonth & " " & cboYear
                End If
            Else
                VSTATUS = "O"
                LABDESC = " TOTAL VEHICLES ON STOCK "
                If Indx = -1 Then
                    FillVehicles VSTATUS, ""
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For  the " & cboYear
                Else
                    FillVehicles VSTATUS, ""
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For the  " & lstMonth & " " & cboYear
                End If
            End If










        Case 4
            If Left(Combo1, 1) <> "" And Left(Combo1, 1) <> "(" Then
                XLOGName = " AND LOGNAME='" & Combo1 & "'"
            End If

            If Indx = -1 Then
                FillLogInquiry " YEAR(DEYT)=" & cboYear & XLOGName
                labVCount = Combo1 & "(" & lstActivity.Records.Count & ") Prospect For the  " & cboYear
            Else
                FillLogInquiry "YEAR(DEYT)=" & cboYear & "  and MONTH(DEYT)=" & Indx & XLOGName
                labVCount = Combo1 & "(" & lstActivity.Records.Count & ") Prospect For the  " & lstMonth & " " & cboYear
            End If


        Case 5
            Priority = ""
            If Left(cboPriority, 1) <> "" And Left(cboPriority, 1) <> "(" Then
                Priority = " AND Priority='" & Left(cboPriority, 1) & "'"
            End If

            If Left(cbostatus, 1) <> "" And Left(cbostatus, 1) <> "(" Then
                xstatus = " AND STATUS='" & Left(cbostatus, 1) & "'"
            End If

            If optReminder(0).Value = True Then
                RSTATUS = " AND ENTITYTYPE='P'" & Priority & xstatus
                LABDESC = "PROSPECT REMINDER(s)"
            ElseIf optReminder(1).Value = True Then
                RSTATUS = " AND ENTITYTYPE='C'" & Priority & xstatus
                LABDESC = "CUSTOMER REMINDER(s)"
            ElseIf optReminder(3).Value = True Then
                RSTATUS = " AND ENTITYTYPE='S'" & Priority & xstatus
                LABDESC = "INTERNAL REMINDER(s)"
            Else
                RSTATUS = "" & Priority & xstatus
                LABDESC = "ALL REMINDER(s)"
            End If

            If Indx = -1 Then
                FillReminder " and  YEAR(DATETIMEREMIND)>=" & cboYear
                labVCount = LABDESC & "(" & lstReminders.Records.Count & ") For the  " & cboYear
            Else
                FillReminder " and  YEAR(DATETIMEREMIND)=" & cboYear & "  and MONTH(DATETIMEREMIND)>=" & Indx
                labVCount = LABDESC & "(" & lstReminders.Records.Count & ") For the  " & lstMonth & " " & cboYear
            End If

    End Select
End Sub


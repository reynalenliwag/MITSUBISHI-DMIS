VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMSCustomerHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Customer and Vehicle Service History Inquiry"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSCustomerServiceHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   12225
   Begin XtremeReportControl.ReportControl rptHIST 
      Height          =   2655
      Left            =   60
      TabIndex        =   24
      Top             =   870
      Width           =   12075
      _Version        =   655364
      _ExtentX        =   21299
      _ExtentY        =   4683
      _StockProps     =   64
      BorderStyle     =   4
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   10455
      TabIndex        =   25
      Top             =   30
      Width           =   10485
      Begin VB.OptionButton Option2 
         Caption         =   "By Vehicle"
         Height          =   255
         Left            =   8940
         TabIndex        =   29
         Top             =   510
         Width           =   1635
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Customer"
         Height          =   255
         Left            =   8940
         TabIndex        =   28
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   30
         TabIndex        =   27
         Top             =   330
         Width           =   8745
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Left            =   -60
         TabIndex        =   26
         Top             =   0
         Width           =   10695
         _Version        =   655364
         _ExtentX        =   18865
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   $"frmCSMSCustomerServiceHistory.frx":1082
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
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "With Transaction"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9570
      TabIndex        =   5
      ToolTipText     =   "View Service History Inquiry with Transaction"
      Top             =   1350
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All Customer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   8
      ToolTipText     =   "View All Customers Service History Inquiry"
      Top             =   1890
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT CUSTOMER HISTORY"
      Height          =   705
      Left            =   10590
      TabIndex        =   23
      ToolTipText     =   "View All Customers Service History Inquiry"
      Top             =   90
      Width           =   1575
   End
   Begin VB.Frame fraAllCustomer 
      Caption         =   "ALL CUSTOMER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   60
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   8925
      Begin VB.OptionButton Otp 
         Caption         =   "By Customer Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   300
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton Otp 
         Caption         =   "By FirstName"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton Otp 
         Caption         =   "By Plate No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3510
         TabIndex        =   11
         Top             =   300
         Width           =   1305
      End
      Begin VB.TextBox txtSearchKey_All 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4830
         TabIndex        =   10
         Top             =   240
         Width           =   4005
      End
   End
   Begin MSComctlLib.ListView lstActiveCustomer 
      Height          =   1965
      Left            =   60
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   3466
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
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":112F
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Acct No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Model"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Home Phone "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Last Update "
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Time Update "
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "mobile "
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   555
      Left            =   12450
      MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":1291
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSCustomerServiceHistory.frx":13E3
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit Window"
      Top             =   7260
      Visible         =   0   'False
      Width           =   975
   End
   Begin XtremeSuiteControls.TabControl OtherTab 
      Height          =   2295
      Left            =   30
      TabIndex        =   4
      Top             =   6150
      Width           =   12105
      _Version        =   655364
      _ExtentX        =   21352
      _ExtentY        =   4048
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   4
      Item(0).Caption =   "Job Detail"
      Item(0).Tooltip =   "Job Detail"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "Listjob"
      Item(1).Caption =   "Parts"
      Item(1).Tooltip =   "Parts"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "ListParts"
      Item(2).Caption =   "Material"
      Item(2).Tooltip =   "Material"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "ListMat"
      Item(3).Caption =   "Accessories"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lsvACC"
      Begin Crystal.CrystalReport rptINQUIRY 
         Left            =   11580
         Top             =   1860
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComctlLib.ListView ListMat 
         Height          =   1905
         Left            =   -69970
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   12045
         _ExtentX        =   21246
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":1749
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Material Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "SRP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Dsc Amount"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView ListParts 
         Height          =   1905
         Left            =   -69970
         TabIndex        =   19
         Top             =   30
         Visible         =   0   'False
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":18AB
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Parts Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "SRP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Dsc Amount"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView Listjob 
         Height          =   1905
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":1A0D
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "OR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "EstimateNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Line No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Job Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Detdsc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Detail"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Technician"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Flat Rate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "STD Hrs"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Detunt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Det Prc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "AMOUNT"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Tax Rate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Tax Val"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Ref Riv Adb"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Save date "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Save Time"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "QC Inspection"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lsvACC 
         Height          =   1905
         Left            =   -69970
         TabIndex        =   22
         Top             =   30
         Visible         =   0   'False
         Width           =   12045
         _ExtentX        =   21246
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":1B6F
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CODE"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Accessories Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "SRP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Dsc Amount"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin XtremeSuiteControls.TabControl MyTab 
      Height          =   2565
      Left            =   30
      TabIndex        =   1
      Top             =   3570
      Width           =   12105
      _Version        =   655364
      _ExtentX        =   21352
      _ExtentY        =   4524
      _StockProps     =   64
      Appearance      =   6
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.FixedTabWidth=   100
      PaintManager.MinTabWidth=   100
      ItemCount       =   2
      Item(0).Caption =   "Repair Order"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "ListRo"
      Item(1).Caption =   "Vehicle Info"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "listCusveh"
      Begin MSComctlLib.ListView listCusveh 
         Height          =   2085
         Left            =   -69940
         TabIndex        =   3
         Top             =   390
         Visible         =   0   'False
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   3678
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
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":1CD1
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Vin"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Plate No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cond No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Clrcde"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Year"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Make"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Model"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Engine"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "KReading"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Product No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Serial No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Tin No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "D Sold"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "X var Cert"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Del. Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Selling Dealer"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListRo 
         Height          =   2115
         Left            =   60
         TabIndex        =   2
         Top             =   390
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   3731
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
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":1E33
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Appointmentdate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "R.O No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Estimate No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Appt No "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cus. Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Plate no"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Hours Work"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Percentage"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Promise Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Writer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Technician 1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Technician 2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Technician 3"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Save Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Save Time "
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraActiveCustomer 
      Caption         =   "WITH TRANSACTION"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   60
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   8475
      Begin VB.TextBox txtSearchKey_Active 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4830
         TabIndex        =   17
         Top             =   210
         Width           =   3525
      End
      Begin VB.OptionButton Wotp 
         Caption         =   "By Plate no."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3510
         TabIndex        =   16
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton Wotp 
         Caption         =   "By First Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   2010
         TabIndex        =   15
         Top             =   300
         Width           =   1425
      End
      Begin VB.OptionButton Wotp 
         Caption         =   "By Customer Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   1965
      End
   End
   Begin MSComctlLib.ListView lstAllCustomer 
      Height          =   1965
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   3466
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
      MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":1F95
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CUSCDE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Acount No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Customer"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "LastName"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "First Name"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "MI"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Sex"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Address"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Provicial Add"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Home Phone"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Telophone No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "PlateNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Email"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Mobile"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Fax"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "City"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Birthdate"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmCSMSCustomerHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheCuscde                                          As String
Dim TheEstimate                                        As String
Dim TheOR                                              As String
Dim xCUSTNAME                                          As String
Dim vPLATENO                                           As String

Function FindTechName(SACODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset
    Dim RSCON                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT TECH_NAME FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & SACODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindTechName = Null2String(RSTMP!TECH_NAME)
    Else
        Set RSCON = gconDMIS.Execute("SELECT COMPANYNAME FROM CSMS_CONTRACTOR WHERE CODE = '" & SACODE & "'")
        If Not (RSCON.BOF And RSCON.EOF) Then
            FindTechName = Null2String(RSCON!CompanyName)
        Else
            FindTechName = ""
        End If
        Set RSCON = Nothing
    End If

    Set RSTMP = Nothing
End Function

Sub Fill_ActiveCustomer()
    Dim SQL                                            As String
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim RS                                             As New ADODB.Recordset

    lstActiveCustomer.Enabled = False

    SQL = "SELECT  DISTINCT ACCT_NO,CUSTOMER,MODEL,CUSTOMERADD,HOMEPHONE, MOBILE  FROM CSMS_vw_ActiveCust"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    lstActiveCustomer.ListItems.Clear
    cnt = 0

    If Not RS.EOF And Not RS.BOF Then
        lstActiveCustomer.Enabled = True
    End If

    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = lstActiveCustomer.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!ACCT_NO)
            ITEM.SubItems(2) = Null2String(!Customer)
            ITEM.SubItems(3) = Null2String(!MODEL)
            ITEM.SubItems(4) = Null2String(!CUSTOMERADD)
            ITEM.SubItems(5) = Null2String(!HomePhone)
            ITEM.SubItems(6) = Null2String(!Mobile)

            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub Fill_ActiveCustomerSearch()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim cnt                                            As String
    Dim ITEM                                           As ListItem
    Dim Keyword                                        As String

    lstActiveCustomer.Enabled = False

    SQL = "SELECT * FROM CSMS_vw_Activecust WHERE"

    Keyword = Trim(txtSearchKey_Active.Text)


    If Len(Keyword) = 0 Then Exit Sub

    If Wotp(0).Value = True Then
        SQL = SQL & " CUSTOMER LIKE '%" & ReplaceQuote(Keyword) & "%'"
    End If

    If Wotp(1).Value = True Then
        SQL = SQL & " FIRSTNAME LIKE '%" & ReplaceQuote(Keyword) & "%'"
    End If

    If Wotp(2).Value = True Then
        SQL = SQL & " Plate_no LIKE '%" & ReplaceQuote(Keyword) & "%'"
    End If

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    lstActiveCustomer.ListItems.Clear

    If Not RS.EOF And Not RS.BOF Then
        lstActiveCustomer.Enabled = True
    End If

    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = lstActiveCustomer.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!ACCT_NO)
            ITEM.SubItems(2) = Null2String(!Customer)
            ITEM.SubItems(3) = Null2String(!MODEL)
            ITEM.SubItems(4) = Null2String(!CUSTOMERADD)
            ITEM.SubItems(5) = Null2String(!HomePhone)
            ITEM.SubItems(6) = Null2String(!lastupdate)
            ITEM.SubItems(7) = Null2String(!TIMEUPDATE)
            ITEM.SubItems(8) = Null2String(!MODEL)
            .MoveNext
        Loop
    End With



    Set RS = Nothing
End Sub

Sub Fill_AllCustomer()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer

    lstAllCustomer.Enabled = False

    SQL = "SELECT TOP 100 cuscde,accountno,acctname,lastname,firstname,middleinitial,sex,customeradd,Provincialadd,telephoneno,plateno,department,email,mobile,fax,city,birthdate,description FROM all_customer"
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    lstAllCustomer.ListItems.Clear
    cnt = 0

    If Not RS.EOF And Not RS.BOF Then
        lstAllCustomer.Enabled = True
    End If

    With RS

        Do While Not .EOF
            cnt = cnt + 1

            Set ITEM = lstAllCustomer.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!CUSCDE)
            ITEM.SubItems(2) = Null2String(!ACCOUNTNO)
            ITEM.SubItems(3) = Null2String(!ACCTNAME)
            ITEM.SubItems(4) = Null2String(!lastname)
            ITEM.SubItems(5) = Null2String(!Firstname)
            ITEM.SubItems(6) = Null2String(!MiddleInitial)
            ITEM.SubItems(7) = Null2String(!Sex)
            ITEM.SubItems(8) = Null2String(!CUSTOMERADD)
            ITEM.SubItems(9) = Null2String(!provincialadd)
            ITEM.SubItems(10) = Null2String(!TelephoneNo)
            ITEM.SubItems(11) = Null2String(!PlateNo)
            ITEM.SubItems(12) = Null2String(!Department)
            ITEM.SubItems(13) = Null2String(!EMAIL)
            ITEM.SubItems(14) = Null2String(!Mobile)
            ITEM.SubItems(15) = Null2String(!Fax)
            ITEM.SubItems(16) = Null2String(!City)
            ITEM.SubItems(17) = Null2String(!BirthDate)
            .MoveNext

        Loop
    End With

    Set RS = Nothing

    Exit Sub



End Sub

Sub Fill_AllCustomerSearch()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Keyword                                        As String

    lstAllCustomer.Enabled = False

    SQL = "SELECT cuscde,accountno,acctname,lastname,firstname,middleinitial,sex,customeradd,Provincialadd,telephoneno,plateno,department,email,mobile,fax,city,birthdate,description FROM all_customer WHERE"
    Keyword = Trim(txtSearchKey_All.Text)

    If Len(Keyword) = 0 Then Exit Sub

    If Otp(0).Value = True Then
        SQL = SQL & " Acctname LIKE '%" & ReplaceQuote(Keyword) & "%' ORDER BY ACCTNAME"
    End If

    If Otp(1).Value = True Then
        SQL = SQL & " Firstname LIKE '%" & ReplaceQuote(Keyword) & "%' ORDER BY FIRSTNAME"
    End If

    If Otp(2).Value = True Then
        SQL = SQL & " Plateno like '%" & ReplaceQuote(Keyword) & "%'"
    End If

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    lstAllCustomer.ListItems.Clear
    cnt = 0

    If Not RS.EOF And Not RS.BOF Then
        lstAllCustomer.Enabled = True
    End If

    With RS
        Do While Not .EOF
            cnt = cnt + 1

            Set ITEM = lstAllCustomer.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!CUSCDE)
            ITEM.SubItems(2) = Null2String(!ACCOUNTNO)
            ITEM.SubItems(3) = Null2String(!ACCTNAME)
            ITEM.SubItems(4) = Null2String(!lastname)
            ITEM.SubItems(5) = Null2String(!Firstname)
            ITEM.SubItems(6) = Null2String(!MiddleInitial)
            ITEM.SubItems(7) = Null2String(!Sex)
            ITEM.SubItems(8) = Null2String(!CUSTOMERADD)
            ITEM.SubItems(9) = Null2String(!provincialadd)
            ITEM.SubItems(10) = Null2String(!TelephoneNo)
            ITEM.SubItems(11) = Null2String(!PlateNo)
            ITEM.SubItems(12) = Null2String(!Department)
            ITEM.SubItems(13) = Null2String(!EMAIL)
            ITEM.SubItems(14) = Null2String(!Mobile)
            ITEM.SubItems(15) = Null2String(!Fax)
            ITEM.SubItems(16) = Null2String(!City)
            ITEM.SubItems(17) = Null2String(!BirthDate)
            .MoveNext

        Loop
    End With


    Set RS = Nothing

    Exit Sub

End Sub

Sub FillJob()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Estimate                                       As String
    Dim ro                                             As String

    Estimate = TheEstimate
    ro = TheOR

    Listjob.Enabled = False

    SQL = "SELECT * FROM CSMS_RO_DET where (rep_or ='" & ro & "' or estimateno = '" & Estimate & "') AND LIVIL = '1'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    Dim I                                              As Integer
    Listjob.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = Listjob.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!REP_OR)
            ITEM.SubItems(2) = Null2String(!EstimateNo)
            ITEM.SubItems(3) = Null2String(!LINE_NO)
            ITEM.SubItems(4) = Null2String(!DETCDE)
            ITEM.SubItems(5) = Null2String(!DETDSC)
            ITEM.SubItems(6) = Null2String(!Detail)
            ITEM.SubItems(7) = FindTechName(LTrim(RTrim(Null2String(!Technician))))
            ITEM.SubItems(8) = Null2String(!FLATRATE)
            ITEM.SubItems(9) = Null2String(!DET_HRS)
            ITEM.SubItems(10) = Null2String(!detunt)
            ITEM.SubItems(11) = Null2String(!DetPrc)
            ITEM.SubItems(12) = FormatNumber(NumericVal(!DetCost))
            ITEM.SubItems(13) = FormatNumber(NumericVal(!taxrate))
            ITEM.SubItems(14) = FormatNumber(NumericVal(!TAXVAL))
            ITEM.SubItems(15) = FormatNumber(NumericVal(!Discount_2))
            ITEM.SubItems(16) = FormatNumber(NumericVal(!DET_AMT)) - FormatNumber(NumericVal(!Discount_2))
            ITEM.SubItems(17) = Null2String(!REF_RIV_ADB)
            ITEM.SubItems(18) = Null2String(!savedate)
            ITEM.SubItems(19) = Null2String(!savetime)

            If Null2String(!Approve) = "Rejected" Then
                ITEM.SubItems(20) = "Rejected"
                For I = 1 To Listjob.ColumnHeaders.Count - 1
                    Listjob.ListItems(cnt).ListSubItems(I).ForeColor = vbRed
                    Listjob.ListItems(cnt).ListSubItems(I).Bold = True
                Next
            ElseIf Null2String(!Approve) = "Approved" Then
                ITEM.SubItems(20) = "Approved"
                For I = 1 To Listjob.ColumnHeaders.Count - 1
                    Listjob.ListItems(cnt).ListSubItems(I).ForeColor = &H8000&
                Next
            Else
                ITEM.SubItems(20) = "Waiting For:QC"
                For I = 1 To Listjob.ColumnHeaders.Count - 1
                    Listjob.ListItems(cnt).ListSubItems(I).ForeColor = vbBlue
                Next
            End If

            .MoveNext
        Loop
    End With

    Listjob.Enabled = True

    Set RS = Nothing
End Sub

Sub FillMaterial()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From CSMS_ro_det Where Rep_OR = '" & ListRo.SelectedItem.ListSubItems(2).Text & "' And Livil = '3' Order By Line_NO")
    ListMat.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = ListMat.ListItems.Add(, , Null2String(RSTMP!LINE_NO))
            ITEM.SubItems(1) = Null2String(RSTMP!DETCDE)
            ITEM.SubItems(2) = Null2String(RSTMP!DETDSC)
            ITEM.SubItems(3) = Null2String(RSTMP!detvol)
            ITEM.SubItems(4) = Format(Null2String(RSTMP!DetPrc), "#,###,##0.00")
            ITEM.SubItems(5) = Format(Null2String(RSTMP!DetCost), "#,###,##0.00")
            ITEM.SubItems(6) = Format(NumericVal(ITEM.SubItems(3)) * NumericVal(ITEM.SubItems(4)), "#,###,##0.00")
            ITEM.SubItems(7) = Format((NumericVal(ITEM.SubItems(6)) / 100) * CDbl(N2Str2IntZero(RSTMP!discrate)), "#,###,##0.00")
            ITEM.SubItems(8) = Format(ITEM.SubItems(6) - ITEM.SubItems(7), "#,###,##0.00")

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub FillParts()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From CSMS_ro_det Where Rep_OR = '" & ListRo.SelectedItem.ListSubItems(2).Text & _
                                 "' And Livil = '2' Order By Line_NO")
    ListParts.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = ListParts.ListItems.Add(, , Null2String(RSTMP!LINE_NO))
            ITEM.SubItems(1) = Null2String(RSTMP!DETCDE)
            ITEM.SubItems(2) = Null2String(RSTMP!DETDSC)
            ITEM.SubItems(3) = Null2String(RSTMP!detvol)
            ITEM.SubItems(4) = Format(Null2String(RSTMP!DetPrc), "#,###,##0.00")
            ITEM.SubItems(5) = Format(Null2String(RSTMP!DetCost), "#,###,##0.00")
            ITEM.SubItems(6) = Format(NumericVal(ITEM.SubItems(3)) * NumericVal(ITEM.SubItems(4)), "#,###,##0.00")
            ITEM.SubItems(7) = Format((NumericVal(ITEM.SubItems(6)) / 100) * CDbl(N2Str2IntZero(RSTMP!discrate)), "#,###,##0.00")
            ITEM.SubItems(8) = Format(ITEM.SubItems(6) - ITEM.SubItems(7), "#,###,##0.00")

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub FillRO(Optional XPLATENO As String)
    Dim RS                                             As New ADODB.Recordset
    Dim SQL                                            As String
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim CUSCDE                                         As String

    CUSCDE = TheCuscde

    ListRo.Enabled = False

    If XPLATENO = "" Then
        SQL = "SELECT * from CSMS_RepairOrder WHERE acct_no = '" & CUSCDE & "' AND TRANSTYPE = 'R'"
        Set RS = New ADODB.Recordset
        Set RS = gconDMIS.Execute(SQL)

        ListRo.ListItems.Clear
        cnt = 0

        If Not RS.EOF And Not RS.BOF Then
            ListRo.Enabled = True
        End If

        On Error GoTo loaderror

        With RS
            Do While Not .EOF
                cnt = cnt + 1
                Set ITEM = ListRo.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!AppointmentDate)
                ITEM.SubItems(2) = Null2String(!RO_NO)
                ITEM.SubItems(3) = Null2String(!EstimateNo)
                ITEM.SubItems(4) = Null2String(!APPTNO)
                ITEM.SubItems(5) = Null2String(!ACCT_NO)
                ITEM.SubItems(6) = Null2String(!PLATE_NO)
                ITEM.SubItems(7) = Null2String(!xHrsWork)
                ITEM.SubItems(8) = Null2String(!percentage)
                ITEM.SubItems(9) = Null2String(!PromiseDate)
                ITEM.SubItems(10) = Null2String(!Status)
                ITEM.SubItems(11) = Null2String(!writer)
                ITEM.SubItems(12) = Null2String(!tech1)
                ITEM.SubItems(13) = Null2String(!tech2)
                ITEM.SubItems(14) = Null2String(!tech3)
                ITEM.SubItems(15) = Null2String(!savedate)
                ITEM.SubItems(16) = Null2String(!savetime)
                .MoveNext
            Loop
        End With


        Set RS = Nothing
        ListRo.Refresh
    Else
        SQL = "SELECT * from CSMS_RepairOrder WHERE PLATE_NO = '" & XPLATENO & "' AND TRANSTYPE = 'R'"
        Set RS = New ADODB.Recordset
        Set RS = gconDMIS.Execute(SQL)

        ListRo.ListItems.Clear
        cnt = 0

        If Not RS.EOF And Not RS.BOF Then
            ListRo.Enabled = True
        End If

        On Error GoTo loaderror

        With RS
            Do While Not .EOF
                cnt = cnt + 1
                Set ITEM = ListRo.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!AppointmentDate)
                ITEM.SubItems(2) = Null2String(!RO_NO)
                ITEM.SubItems(3) = Null2String(!EstimateNo)
                ITEM.SubItems(4) = Null2String(!APPTNO)
                ITEM.SubItems(5) = Null2String(!ACCT_NO)
                ITEM.SubItems(6) = Null2String(!PLATE_NO)
                ITEM.SubItems(7) = Null2String(!xHrsWork)
                ITEM.SubItems(8) = Null2String(!percentage)
                ITEM.SubItems(9) = Null2String(!PromiseDate)
                ITEM.SubItems(10) = Null2String(!Status)
                ITEM.SubItems(11) = Null2String(!writer)
                ITEM.SubItems(12) = Null2String(!tech1)
                ITEM.SubItems(13) = Null2String(!tech2)
                ITEM.SubItems(14) = Null2String(!tech3)
                ITEM.SubItems(15) = Null2String(!savedate)
                ITEM.SubItems(16) = Null2String(!savetime)
                .MoveNext
            Loop
        End With


        Set RS = Nothing
        ListRo.Refresh
    End If

    Exit Sub

loaderror:

    MsgBox Err.Description

End Sub

Sub FillVehicle(Optional XPLATENO As String)
    Dim RS                                             As New ADODB.Recordset
    Dim SQL                                            As String
    Dim ITEM                                           As ListItem
    Dim CUSCDE                                         As String
    Dim cnt                                            As String

    CUSCDE = TheCuscde

    If XPLATENO = "" Then
        SQL = "SELECT * FROM CSMS_cusveh WHERE cuscde = '" & CUSCDE & "'"
        Set RS = New ADODB.Recordset
        Set RS = gconDMIS.Execute(SQL)

        listCusveh.ListItems.Clear
        cnt = 0

        If Not RS.EOF And Not RS.BOF Then
            ListRo.Enabled = True
        End If

        With RS
            Do While Not .EOF
                cnt = cnt + 1
                Set ITEM = listCusveh.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!NIYM)
                ITEM.SubItems(2) = Null2String(!VIN)
                ITEM.SubItems(3) = Null2String(!PLATE_NO)
                ITEM.SubItems(4) = Null2String(!VCOND_NO)
                ITEM.SubItems(5) = Null2String(!ClrCde)
                ITEM.SubItems(6) = Null2String(!YER)
                ITEM.SubItems(7) = Null2String(!Make)
                ITEM.SubItems(8) = Null2String(!MODEL)
                ITEM.SubItems(9) = Null2String(!Engine)
                ITEM.SubItems(10) = Null2String(!KMReading)
                ITEM.SubItems(11) = Null2String(!ProdNo)
                ITEM.SubItems(12) = Null2String(!SERIAL)
                ITEM.SubItems(13) = Null2String(!TIN_Number)
                ITEM.SubItems(14) = Null2String(!D_SOLD)
                ITEM.SubItems(16) = Null2String(!DEL_DATE)
                .MoveNext
            Loop
        End With
    Else
        Set RS = New ADODB.Recordset
        Set RS = gconDMIS.Execute("SELECT * FROM CSMS_cusveh WHERE PLATE_NO = '" & XPLATENO & "'")

        listCusveh.ListItems.Clear
        cnt = 0

        If Not RS.EOF And Not RS.BOF Then
            ListRo.Enabled = True
        End If

        With RS
            Do While Not .EOF
                cnt = cnt + 1
                Set ITEM = listCusveh.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!NIYM)
                ITEM.SubItems(2) = Null2String(!VIN)
                ITEM.SubItems(3) = Null2String(!PLATE_NO)
                ITEM.SubItems(4) = Null2String(!VCOND_NO)
                ITEM.SubItems(5) = Null2String(!ClrCde)
                ITEM.SubItems(6) = Null2String(!YER)
                ITEM.SubItems(7) = Null2String(!Make)
                ITEM.SubItems(8) = Null2String(!MODEL)
                ITEM.SubItems(9) = Null2String(!Engine)
                ITEM.SubItems(10) = Null2String(!KMReading)
                ITEM.SubItems(11) = Null2String(!ProdNo)
                ITEM.SubItems(12) = Null2String(!SERIAL)
                ITEM.SubItems(13) = Null2String(!TIN_Number)
                ITEM.SubItems(14) = Null2String(!D_SOLD)
                ITEM.SubItems(16) = Null2String(!DEL_DATE)
                .MoveNext
            Loop
        End With
    End If
    Set RS = Nothing
End Sub

Sub FillAccessories()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From CSMS_ro_det Where Rep_OR = '" & ListRo.SelectedItem.ListSubItems(2).Text & "' And Livil = '4' Order By Line_NO")
    lsvAcc.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvAcc.ListItems.Add(, , Null2String(RSTMP!LINE_NO))
            ITEM.SubItems(1) = Null2String(RSTMP!DETCDE)
            ITEM.SubItems(2) = Null2String(RSTMP!DETDSC)
            ITEM.SubItems(3) = Null2String(RSTMP!detvol)
            ITEM.SubItems(4) = Format(Null2String(RSTMP!DetPrc), "#,###,##0.00")
            ITEM.SubItems(5) = Format(Null2String(RSTMP!DetCost), "#,###,##0.00")
            ITEM.SubItems(6) = Format(ITEM.SubItems(3) * ITEM.SubItems(4), "#,###,##0.00")
            ITEM.SubItems(7) = Format((ITEM.SubItems(6) / 100) * CDbl(Null2String(RSTMP!discrate)), "#,###,##0.00")
            ITEM.SubItems(8) = Format(ITEM.SubItems(6) - ITEM.SubItems(7), "#,###,##0.00")

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub Myrefresh()
    ListRo.ListItems.Clear
    Listjob.ListItems.Clear
    listCusveh.ListItems.Clear
    lstAllCustomer.ListItems.Clear
    lstActiveCustomer.ListItems.Clear
End Sub

Private Sub CmdAll_Click()
    Call Myrefresh
    lstAllCustomer.Visible = True
    fraAllCustomer.Visible = True
    fraActiveCustomer.Visible = False
    lstActiveCustomer.ZOrder 1
    Fill_AllCustomer
    If lstAllCustomer.ListItems.Count > 0 Then
        lstAllCustomer.ListItems(1).Selected = True
        lstAllCustomer.ListItems(1).EnsureVisible
        lstAllCustomer_ItemClick lstActiveCustomer.SelectedItem
    Else
        ListRo.ListItems.Clear
        Listjob.ListItems.Clear
        listCusveh.ListItems.Clear
    End If

    LogAudit "I", "CUSTOMER AND VEHICLE SERVICE HISTORY INQUIRY - ALL CUSTOMER "

    On Error Resume Next
    txtSearchKey_All.SetFocus

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Option1.Value = True Then
        If Not TheCuscde = "" Then
            Screen.MousePointer = 11
            rptINQUIRY.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptINQUIRY.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptINQUIRY.WindowTitle = "Service Vehicle History"
            PrintSQLReport rptINQUIRY, CSMS_REPORT_PATH & "VehicleHistory.rpt", "{ALL_CUSTOMER_TABLE.CUSCDE} = '" & TheCuscde & "'", CSMS_REPORT_CONNECTION, 1

            Call NEW_LogAudit("V", "CUSTOMER VEHICLE INQUIRY", "", "", "", "CUST NAME: " & xCUSTNAME, "", "")
            Screen.MousePointer = 0
        Else
            MsgBox "Choose a Customer to Print", vbInformation, "CSMS"
            Exit Sub
        End If
    Else
        If Not vPLATENO = "" Then
            Screen.MousePointer = 11
            rptINQUIRY.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptINQUIRY.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptINQUIRY.WindowTitle = "Service Vehicle History"
            PrintSQLReport rptINQUIRY, CSMS_REPORT_PATH & "VehicleHistory_vehicle.rpt", "{CSMS_REPOR.PLATE_NO} = '" & vPLATENO & "'", CSMS_REPORT_CONNECTION, 1

            Call NEW_LogAudit("V", "CUSTOMER VEHICLE INQUIRY", "", "", "", "PLATE NO: " & vPLATENO, "", "")
            Screen.MousePointer = 0
        Else
            MsgBox "Choose a Plate no to Print", vbInformation, "CSMS"
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdTrans_Click()
    Myrefresh
    Fill_ActiveCustomer

    lstActiveCustomer.ZOrder 0
    fraAllCustomer.Visible = False
    fraActiveCustomer.Visible = True

    lstAllCustomer.Visible = False

    If lstActiveCustomer.ListItems.Count > 0 Then
        lstActiveCustomer.ListItems(1).Selected = True
        lstActiveCustomer.ListItems(1).EnsureVisible
        lstActiveCustomer_ItemClick lstActiveCustomer.SelectedItem
    Else
        ListRo.ListItems.Clear
        Listjob.ListItems.Clear
        listCusveh.ListItems.Clear
    End If

    'CALL NEW_LogAudit( "I", "USTOMER VEHICLE INQUIRY. - W/TRANSACTION"

    On Error Resume Next
    txtSearchKey_Active.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CUSTOMER VEHICLE INQUIRY)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CUSTOMER VEHICLE INQUIRY", "PRINTING")

        Case vbKeyF3:
            On Error Resume Next
            Text1.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    Option1.Value = True

    'Call CmdAll_Click
End Sub

Private Sub ListRo_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    TheEstimate = ListRo.ListItems.ITEM(ListRo.SelectedItem.Index)
    TheEstimate = ListRo.SelectedItem.SubItems(3)

    TheOR = ListRo.ListItems.ITEM(ListRo.SelectedItem.Index)
    TheOR = ListRo.SelectedItem.SubItems(2)

    Call FillJob
    Call FillParts
    Call FillMaterial
    Call FillAccessories

    If OtherTab(1).Selected = True Then
        OtherTab(0).Selected = True
    End If
End Sub

Private Sub lstActiveCustomer_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next

    TheCuscde = lstActiveCustomer.ListItems.ITEM(lstActiveCustomer.SelectedItem.Index)
    TheCuscde = lstActiveCustomer.SelectedItem.SubItems(1)
    xCUSTNAME = Null2String(lstActiveCustomer.SelectedItem.SubItems(2))

    Call FillRO
    Call FillVehicle

    If MyTab(1).Selected = True Then
        MyTab(0).Selected = True
    End If

    If ListRo.ListItems.Count > 0 Then
        ListRo.ListItems(1).Selected = True
        ListRo.ListItems(1).EnsureVisible
        ListRo_ItemClick ListRo.SelectedItem
    Else
        ListRo.ListItems.Clear
        Listjob.ListItems.Clear
        listCusveh.ListItems.Clear
    End If

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("I", "CUSTOMER VEHICLE INQUIRY", "", "", "", "CUST NAME: " & xCUSTNAME, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
End Sub

Private Sub lstAllCustomer_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    TheCuscde = lstAllCustomer.ListItems.ITEM(lstAllCustomer.SelectedItem.Index)
    TheCuscde = lstAllCustomer.SelectedItem.SubItems(1)
    xCUSTNAME = Null2String(lstAllCustomer.SelectedItem.SubItems(3))

    Call FillVehicle
    Call FillRO

    If ListRo.ListItems.Count > 0 Then
        ListRo.ListItems(1).Selected = True
        ListRo.ListItems(1).EnsureVisible
        ListRo_ItemClick ListRo.SelectedItem
    Else
        ListRo.ListItems.Clear
        Listjob.ListItems.Clear
    End If

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("I", "CUSTOMER VEHICLE INQUIRY", "", "", "", "CUST NAME: " & xCUSTNAME, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        cmdPrint.Caption = "PRINT CUSTOMER HISTORY"
        Call DisplayHistory
    Else
        cmdPrint.Caption = "PRINT VEHICLE HISTORY"
        Call DisplayVehicleHistory
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        cmdPrint.Caption = "PRINT VEHICLE HISTORY"
        Call DisplayVehicleHistory
    Else
        cmdPrint.Caption = "PRINT CUSTOMER HISTORY"
        Call DisplayHistory
    End If
End Sub

Private Sub Otp_Click(Index As Integer)
    On Error Resume Next
    txtSearchKey_All.SetFocus
End Sub

Private Sub rptHIST_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    Dim Index                                          As Long
    Dim vID                                            As Long


    If Row.Record Is Nothing Then: Exit Sub


    If Option1.Value = True Then
        TheCuscde = Null2String(Row.Record(0).Value)  'CUSTOMER CODE
        xCUSTNAME = Null2String(Row.Record(1).Value)  'CUSTOMER NAME

        Call FillRO
        Call FillVehicle

        If MyTab(1).Selected = True Then
            MyTab(0).Selected = True
        End If

        If ListRo.ListItems.Count > 0 Then
            ListRo.ListItems(1).Selected = True
            ListRo.ListItems(1).EnsureVisible
            ListRo_ItemClick ListRo.SelectedItem
        Else
            ListRo.ListItems.Clear
            Listjob.ListItems.Clear
            listCusveh.ListItems.Clear
        End If

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("I", "CUSTOMER VEHICLE INQUIRY", "", "", "", "CUST NAME: " & xCUSTNAME, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        vPLATENO = Null2String(Row.Record(3).Value)   'PLATE NO

        Call FillRO(vPLATENO)
        Call FillVehicle(vPLATENO)

        If MyTab(1).Selected = True Then
            MyTab(0).Selected = True
        End If

        If ListRo.ListItems.Count > 0 Then
            ListRo.ListItems(1).Selected = True
            ListRo.ListItems(1).EnsureVisible
            ListRo_ItemClick ListRo.SelectedItem
        Else
            ListRo.ListItems.Clear
            Listjob.ListItems.Clear
            listCusveh.ListItems.Clear
        End If

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("I", "CUSTOMER VEHICLE INQUIRY", "", "", "", "PLATE NO: " & vPLATENO, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    End If
End Sub

Private Sub Text1_Change()
    rptHIST.FilterText = Text1.Text
    rptHIST.Populate
End Sub

Private Sub Text1_GotFocus()
    Text1.BackColor = &HC0FFFF
End Sub

Private Sub Text1_LostFocus()
    Text1.BackColor = vbWhite
End Sub

Private Sub txtSearchKey_Active_Change()
    Fill_ActiveCustomerSearch
    If lstActiveCustomer.ListItems.Count > 0 Then
        lstActiveCustomer.ListItems(1).Selected = True
        lstActiveCustomer.ListItems(1).EnsureVisible
        'lstActiveCustomer_ItemClick lstActiveCustomer.SelectedItem
    Else
        ListRo.ListItems.Clear
        Listjob.ListItems.Clear
        listCusveh.ListItems.Clear
    End If
End Sub

Private Sub txtSearchKey_All_Change()

    '    Fill_AllCustomerSearch
    '    If lstAllCustomer.ListItems.Count > 0 Then
    '        lstAllCustomer.ListItems(1).Selected = True
    '        lstAllCustomer.ListItems(1).EnsureVisible
    '        'lstAllCustomer_ItemClick lstActiveCustomer.SelectedItem
    '    Else
    '        ListRo.ListItems.Clear
    '        Listjob.ListItems.Clear
    '    End If
End Sub

Private Sub Wotp_Click(Index As Integer)
    On Error Resume Next
    txtSearchKey_Active.SetFocus
End Sub

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                            As ADODB.Field
    Dim j                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Sub ReportControlAddColumnHeader(lst As ReportControl, StringHeaders As String)
    Dim ar()                                           As String
    Dim I                                              As Integer

    ar = Split(StringHeaders, ",")
    lst.Columns.DeleteAll
    For I = LBound(ar) To UBound(ar)
        lst.Columns.Add I, ar(I), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(lst As ReportControl)
    With lst
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                           As String
    Dim cWidth                                         As Long
    Dim I                                              As Integer
    Dim scwidth                                        As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub

Public Sub DisplayHistory()
    On Error Resume Next
    Dim RSTMP                                          As New ADODB.Recordset
    Screen.MousePointer = 11

    TheCuscde = ""
    Text1.Text = ""
    Set RSTMP = gconDMIS.Execute("SELECT CUSCDE,ACCTNAME,CUSTOMERADD,HOMEPHONE,TELEPHONENO,MOBILE,EMAIL FROM ALL_CUSTOMER_TABLE ORDER BY ACCTNAME")
    Call ReportControlAddColumnHeader(rptHIST, "CUSCODE, CUSTOMER NAME, ADDRESS, HOME PHONE, TELEPHONE, MOBILE NO, EMAIL ADD")
    Call ReportControlPaintManager(rptHIST)
    'rptHIST.GroupsOrder.Add rptHIST.Columns(1)
    rptHIST.Columns(0).Visible = False
    'rptHIST.Columns(4).Visible = False
    'rptHIST.Columns(8).Visible = False
    Call ResizeColumnHeader(rptHIST, "0, 25, 25, 10, 10, 10, 10")
    Call flex_FillReportView(RSTMP, rptHIST)

    Text1.SetFocus
    Screen.MousePointer = 0
End Sub

Public Sub DisplayVehicleHistory()
    On Error Resume Next
    Dim RSTMP                                          As New ADODB.Recordset
    Screen.MousePointer = 11

    vPLATENO = ""
    Text1.Text = ""
    Set RSTMP = gconDMIS.Execute("SELECT CUSCDE, NIYM, VIN, PLATE_NO, VCOND_NO, SERIAL, MAKE, DESCRIPTION FROM CSMS_CUSVEH ORDER BY PLATE_NO")
    Call ReportControlAddColumnHeader(rptHIST, "CUSCODE, CUSTOMER NAME, VIN NO, PLATE NO, COND NO., SERIAL NO., MAKE, DESCRIPTION")
    Call ReportControlPaintManager(rptHIST)
    rptHIST.Columns(0).Visible = False
    rptHIST.Columns(5).Visible = False
    Call ResizeColumnHeader(rptHIST, "0, 25, 15, 8, 8, 15, 10, 20")
    Call flex_FillReportView(RSTMP, rptHIST)

    Text1.SetFocus
    Screen.MousePointer = 0
End Sub

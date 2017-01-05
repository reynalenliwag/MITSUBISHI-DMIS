VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmCSMSCustomerHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Customer and Vehicle Service History Inquiry"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   Icon            =   "frmCSMSCustomerServiceHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAllCustomer 
      Caption         =   "All Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   9195
      Begin VB.OptionButton Otp 
         Caption         =   "By Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Otp 
         Caption         =   "By FirstName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2130
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Otp 
         Caption         =   "By Plate No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3690
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.TextBox txtSearchKey_All 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5070
         TabIndex        =   10
         Top             =   210
         Width           =   3975
      End
   End
   Begin MSComctlLib.ListView lstActiveCustomer 
      Height          =   1965
      Left            =   60
      TabIndex        =   6
      Top             =   840
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
      MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":058A
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Acct No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Model"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   5292
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12000
      MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":06EC
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSCustomerServiceHistory.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit Window"
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin XtremeSuiteControls.TabControl OtherTab 
      Height          =   2295
      Left            =   30
      TabIndex        =   4
      Top             =   5460
      Width           =   12105
      _Version        =   655364
      _ExtentX        =   21352
      _ExtentY        =   4048
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   1
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":0BA4
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
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   2540
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":0D06
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
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   2540
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":0E68
         NumItems        =   22
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
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
            Text            =   "Detcde"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Detdsc"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Technician"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Flat Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Det Hrs"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Detunt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Det Prc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Det Cost"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "DetAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Tax Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Tax Val"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Detail"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Det Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Ref Riv Adb"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Save date "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Save Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":0FCA
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
            Object.Width           =   2540
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
      Top             =   2820
      Width           =   12135
      _Version        =   655364
      _ExtentX        =   21405
      _ExtentY        =   4524
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
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
         Width           =   12015
         _ExtentX        =   21193
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":112C
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
            Text            =   "Vcond No"
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "del_date"
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
         Width           =   12015
         _ExtentX        =   21193
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":128E
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NO"
            Object.Width           =   882
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
            Text            =   "Acct no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Plate no"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "xhrswork"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Percentage"
            Object.Width           =   2540
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
      Caption         =   "W/Transaction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   9195
      Begin VB.TextBox txtSearchKey_Active 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5100
         TabIndex        =   17
         Top             =   270
         Width           =   3975
      End
      Begin VB.OptionButton Wotp 
         Caption         =   "By Plate no."
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
         Index           =   2
         Left            =   3690
         TabIndex        =   16
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton Wotp 
         Caption         =   "By First Name"
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
         Index           =   1
         Left            =   2130
         TabIndex        =   15
         Top             =   300
         Width           =   1935
      End
      Begin VB.OptionButton Wotp 
         Caption         =   "By Customer Name"
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
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10710
      TabIndex        =   8
      ToolTipText     =   "View All Customers Service History Inquiry"
      Top             =   240
      Width           =   1425
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "W/ Transaction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      ToolTipText     =   "View Service History Inquiry with Transaction"
      Top             =   240
      Width           =   1365
   End
   Begin MSComctlLib.ListView lstAllCustomer 
      Height          =   1965
      Left            =   60
      TabIndex        =   0
      Top             =   840
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
      MouseIcon       =   "frmCSMSCustomerServiceHistory.frx":13F0
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CUSCDE"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Acount No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Customer"
         Object.Width           =   5292
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
         Object.Width           =   5292
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
'FUNCTION /FEATURE:Display All The Customer Where Use To Locate the History Of the Car
'DATE STARTED:05/12/2007
'LAST UPDATE:
'DATABASE UPDATE:
'WHO UPDATE:HardNard
'UPDATING CODE:BTT - 05102007
'*******************************************************************************************
'FUNCTION /FEATURE:Add New Control For Searching,Complete The Algo For Searching
'DATE STARTED:05/15/2007
'LAST UPDATE:
'DATABASE UPDATE:
'WHO UPDATE:HardNard
'UPDATING CODE:BTT - 05102007
'*******************************************************************************************
'FUNCTION /FEATURE:Displaying The Job Of particular Customer using ListView
'DATE STARTED:05/16/2007
'LAST UPDATE:
'DATABASE UPDATE:Created New ViewTable For all the Customer with the Transaction only;ViewTable Name = CSMS_vw_AllCustomer
'WHO UPDATE:HardNard
'UPDATING CODE:BTT - 05102007
'********************************************************************************************
'==========================================================================================
'FUNCTION / FEATURE :Added Auto Select For LISTRO FOR ALL CUST and Active CUST
'DATE STARTED       :5/31/200719:46
'LAST UPDATED       :5/31/200719:46
'DATABASE UPDATES   :
'WHO UPDATED        :AXP  5/31/2007
'UDPATING CODE    :AXP-5312007-A
'==========================================================================================
Option Explicit
Dim TheCuscde                           As String
Dim TheEstimate                         As String
Dim TheOR                               As String

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


    On Error Resume Next
    txtSearchKey_All.SetFocus

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdTrans_Click()
    Call Myrefresh
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
    On Error Resume Next
    txtSearchKey_Active.SetFocus
End Sub

Sub Fill_ActiveCustomer()
    Dim SQL                             As String
    Dim Item                            As ListItem
    Dim cnt                             As Integer
    Dim RS                              As New ADODB.Recordset

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
            Set Item = lstActiveCustomer.ListItems.Add(, , cnt)
            Item.SubItems(1) = Null2String(!ACCT_NO)
            Item.SubItems(2) = Null2String(!Customer)
            Item.SubItems(3) = Null2String(!Model)
            Item.SubItems(4) = Null2String(!CUSTOMERADD)
            Item.SubItems(5) = Null2String(!HomePhone)
            Item.SubItems(6) = Null2String(!Mobile)

            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub Fill_ActiveCustomerSearch()
    Dim SQL                             As String
    Dim RS                              As New ADODB.Recordset
    Dim cnt                             As String
    Dim Item                            As ListItem
    Dim Keyword                         As String

    lstActiveCustomer.Enabled = False

    SQL = "SELECT DISTINCT * FROM CSMS_vw_Activecust WHERE"

    Keyword = Trim(txtSearchKey_Active.Text)


    If Len(Keyword) = 0 Then Exit Sub

    If Wotp(0).Value = True Then
        SQL = SQL & " lastname LIKE '" & ReplaceQuote(Keyword) & "%'"
    End If

    If Wotp(1).Value = True Then
        SQL = SQL & " firstname LIKE '" & ReplaceQuote(Keyword) & "%'"
    End If

    If Wotp(2).Value = True Then
        SQL = SQL & " Plate_no LIKE '" & ReplaceQuote(Keyword) & "%'"
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
            Set Item = lstActiveCustomer.ListItems.Add(, , cnt)
            Item.SubItems(1) = Null2String(!ACCT_NO)
            Item.SubItems(2) = Null2String(!Customer)
            Item.SubItems(3) = Null2String(!Model)
            Item.SubItems(4) = Null2String(!CUSTOMERADD)
            Item.SubItems(5) = Null2String(!HomePhone)
            Item.SubItems(6) = Null2String(!lastupdate)
            Item.SubItems(7) = Null2String(!TIMEUPDATE)
            Item.SubItems(8) = Null2String(!Model)
            .MoveNext
        Loop
    End With



    Set RS = Nothing
End Sub

Sub Fill_AllCustomer()
    Dim SQL                             As String
    Dim RS                              As New ADODB.Recordset
    Dim Item                            As ListItem
    Dim cnt                             As Integer

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

            Set Item = lstAllCustomer.ListItems.Add(, , cnt)
            Item.SubItems(1) = Null2String(!CUSCDE)
            Item.SubItems(2) = Null2String(!ACCOUNTNO)
            Item.SubItems(3) = Null2String(!AcctName)
            Item.SubItems(4) = Null2String(!lastname)
            Item.SubItems(5) = Null2String(!Firstname)
            Item.SubItems(6) = Null2String(!MiddleInitial)
            Item.SubItems(7) = Null2String(!Sex)
            Item.SubItems(8) = Null2String(!CUSTOMERADD)
            Item.SubItems(9) = Null2String(!provincialadd)
            Item.SubItems(10) = Null2String(!TelephoneNo)
            Item.SubItems(11) = Null2String(!Plateno)
            Item.SubItems(12) = Null2String(!Department)
            Item.SubItems(13) = Null2String(!EMAIL)
            Item.SubItems(14) = Null2String(!Mobile)
            Item.SubItems(15) = Null2String(!Fax)
            Item.SubItems(16) = Null2String(!City)
            Item.SubItems(17) = Null2String(!BirthDate)
            .MoveNext

        Loop
    End With

    Set RS = Nothing

    Exit Sub



End Sub

Sub Fill_AllCustomerSearch()
    Dim SQL                             As String
    Dim RS                              As New ADODB.Recordset
    Dim Item                            As ListItem
    Dim cnt                             As Integer
    Dim Keyword                         As String

    '    On Error GoTo loaderror

    lstAllCustomer.Enabled = False

    SQL = "SELECT cuscde,accountno,acctname,lastname,firstname,middleinitial,sex,customeradd,Provincialadd,telephoneno,plateno,department,email,mobile,fax,city,birthdate,description FROM all_customer WHERE"
    Keyword = Trim(txtSearchKey_All.Text)

    If Len(Keyword) = 0 Then Exit Sub

    If Otp(0).Value = True Then
        SQL = SQL & " Acctname LIKE '" & ReplaceQuote(Keyword) & "%'"
    End If

    If Otp(1).Value = True Then
        SQL = SQL & " Firstname LIKE'" & ReplaceQuote(Keyword) & "%'"
    End If

    If Otp(2).Value = True Then
        SQL = SQL & " Plateno like '" & ReplaceQuote(Keyword) & "%'"
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

            Set Item = lstAllCustomer.ListItems.Add(, , cnt)
            Item.SubItems(1) = Null2String(!CUSCDE)
            Item.SubItems(2) = Null2String(!ACCOUNTNO)
            Item.SubItems(3) = Null2String(!AcctName)
            Item.SubItems(4) = Null2String(!lastname)
            Item.SubItems(5) = Null2String(!Firstname)
            Item.SubItems(6) = Null2String(!MiddleInitial)
            Item.SubItems(7) = Null2String(!Sex)
            Item.SubItems(8) = Null2String(!CUSTOMERADD)
            Item.SubItems(9) = Null2String(!provincialadd)
            Item.SubItems(10) = Null2String(!TelephoneNo)
            Item.SubItems(11) = Null2String(!Plateno)
            Item.SubItems(12) = Null2String(!Department)
            Item.SubItems(13) = Null2String(!EMAIL)
            Item.SubItems(14) = Null2String(!Mobile)
            Item.SubItems(15) = Null2String(!Fax)
            Item.SubItems(16) = Null2String(!City)
            Item.SubItems(17) = Null2String(!BirthDate)
            '            Item.SubItems(18) = Null2String(!Description)
            'Item.SubItems(19) = Null2String(!Description)
            '  DoEvents
            '  I = I + 1
            '  loadme.Value = (I / rs.RecordCount) * 100
            '  Rema.Caption = Int(loadme.Value) & "% Completed"
            '  DoEvents
            .MoveNext

        Loop
    End With


    Set RS = Nothing

    Exit Sub

    'lstAllCustomer.Enabled = True
    'loadme.Visible = False
    'Rema.Visible = False
    'loaderror:

    '   MsgBox "Please Select A criteria", vbInformation, "Information"
    '   txtSearchKey_All.Text = ""

End Sub

Sub FillJob()
    Dim SQL                             As String
    Dim RS                              As New ADODB.Recordset
    Dim Item                            As ListItem
    Dim cnt                             As Integer
    Dim Estimate                        As String
    Dim ro                              As String

    Estimate = TheEstimate
    ro = TheOR

    Listjob.Enabled = False

    SQL = "SELECT * FROM CSMS_RO_DET where rep_or='" & ro & "' or estimateno='" & Estimate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    Dim i                               As Integer
    Listjob.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set Item = Listjob.ListItems.Add(, , cnt)
            Item.SubItems(1) = Null2String(!REP_OR)
            Item.SubItems(2) = Null2String(!EstimateNo)
            Item.SubItems(3) = Null2String(!Line_No)
            Item.SubItems(4) = Null2String(!DETCDE)
            Item.SubItems(5) = Null2String(!DETDSC)
            Item.SubItems(6) = Null2String(!TECHNICIAN)
            Item.SubItems(7) = Null2String(!FLATRATE)
            Item.SubItems(8) = Null2String(!DET_HRS)
            Item.SubItems(9) = Null2String(!detunt)
            Item.SubItems(10) = Null2String(!DetPrc)
            Item.SubItems(11) = FormatNumber(NumericVal(!DetCost))
            Item.SubItems(12) = FormatNumber(NumericVal(!detamt))
            Item.SubItems(13) = FormatNumber(NumericVal(!taxrate))
            Item.SubItems(14) = FormatNumber(NumericVal(!TAXVAL))
            Item.SubItems(15) = Null2String(!Detail)
            Item.SubItems(16) = FormatNumber(NumericVal(!DET_AMT))
            Item.SubItems(17) = FormatNumber(NumericVal(!Discount_2))
            Item.SubItems(18) = Null2String(!REF_RIV_ADB)
            Item.SubItems(19) = Null2String(!savedate)
            Item.SubItems(20) = Null2String(!savetime)

            If Null2String(!Approve) = "Rejected" Then
                Item.SubItems(21) = "Rejected"
                For i = 1 To Listjob.ColumnHeaders.Count - 1
                    Listjob.ListItems(cnt).ListSubItems(i).ForeColor = vbRed
                    Listjob.ListItems(cnt).ListSubItems(i).Bold = True
                Next

            ElseIf Null2String(!Approve) = "Approved" Then
                Item.SubItems(21) = "Approved"
                For i = 1 To Listjob.ColumnHeaders.Count - 1
                    Listjob.ListItems(cnt).ListSubItems(i).ForeColor = &H8000&
                Next
            Else
                Item.SubItems(21) = "Waiting For:QC"

                For i = 1 To Listjob.ColumnHeaders.Count - 1
                    Listjob.ListItems(cnt).ListSubItems(i).ForeColor = vbBlue
                Next

            End If

            .MoveNext
        Loop
    End With

    Listjob.Enabled = True

    Set RS = Nothing
End Sub

Sub FillMaterial()
    Dim rsTmp                           As New ADODB.Recordset
    Dim Item                            As ListItem

    Set rsTmp = gconDMIS.Execute("Select * From CSMS_ro_det Where Rep_OR = '" & ListRo.SelectedItem.ListSubItems(2).Text & "' And Livil = '3' Order By Line_NO")
    ListMat.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = ListMat.ListItems.Add(, , Null2String(rsTmp!Line_No))
            Item.SubItems(1) = Null2String(rsTmp!DETCDE)
            Item.SubItems(2) = Null2String(rsTmp!DETDSC)
            Item.SubItems(3) = Null2String(rsTmp!detvol)
            Item.SubItems(4) = Format(Null2String(rsTmp!DetPrc), "#,###,##0.00")
            Item.SubItems(5) = Format(Null2String(rsTmp!DetCost), "#,###,##0.00")
            Item.SubItems(6) = Format(Item.SubItems(3) * Item.SubItems(4), "#,###,##0.00")
            Item.SubItems(7) = Format((Item.SubItems(6) / 100) * CDbl(Null2String(rsTmp!discrate)), "#,###,##0.00")
            Item.SubItems(8) = Format(Item.SubItems(6) - Item.SubItems(7), "#,###,##0.00")

            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
End Sub
'UPDATE BY : MJP 09-12-2007 04:15 PM ---------------------------------------------------End Sub

Sub FillParts()
    Dim rsTmp                           As New ADODB.Recordset
    Dim Item                            As ListItem

    Set rsTmp = gconDMIS.Execute("Select * From CSMS_ro_det Where Rep_OR = '" & ListRo.SelectedItem.ListSubItems(2).Text & "' And Livil = '2' Order By Line_NO")
    ListParts.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = ListParts.ListItems.Add(, , Null2String(rsTmp!Line_No))
            Item.SubItems(1) = Null2String(rsTmp!DETCDE)
            Item.SubItems(2) = Null2String(rsTmp!DETDSC)
            Item.SubItems(3) = Null2String(rsTmp!detvol)
            Item.SubItems(4) = Format(Null2String(rsTmp!DetPrc), "#,###,##0.00")
            Item.SubItems(5) = Format(Null2String(rsTmp!DetCost), "#,###,##0.00")
            Item.SubItems(6) = Format(Item.SubItems(3) * Item.SubItems(4), "#,###,##0.00")
            Item.SubItems(7) = Format((Item.SubItems(6) / 100) * CDbl(Null2String(rsTmp!discrate)), "#,###,##0.00")
            Item.SubItems(8) = Format(Item.SubItems(6) - Item.SubItems(7), "#,###,##0.00")

            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
End Sub

Sub FillRO()
    Dim RS                              As New ADODB.Recordset
    Dim SQL                             As String
    Dim Item                            As ListItem
    Dim cnt                             As Integer
    Dim CUSCDE                          As String

    CUSCDE = TheCuscde

    ListRo.Enabled = False

    SQL = "SELECT * from CSMS_RepairOrder WHERE acct_no='" & CUSCDE & "'"

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
            Set Item = ListRo.ListItems.Add(, , cnt)
            Item.SubItems(1) = Null2String(!AppointmentDate)
            Item.SubItems(2) = Null2String(!RO_NO)
            Item.SubItems(3) = Null2String(!EstimateNo)
            Item.SubItems(4) = Null2String(!APPTNO)
            Item.SubItems(5) = Null2String(!ACCT_NO)
            Item.SubItems(6) = Null2String(!PLATE_NO)
            Item.SubItems(7) = Null2String(!xHrsWork)
            Item.SubItems(8) = Null2String(!percentage)
            Item.SubItems(9) = Null2String(!promisedate)
            Item.SubItems(10) = Null2String(!STATUS)
            Item.SubItems(11) = Null2String(!writer)
            Item.SubItems(12) = Null2String(!tech1)
            Item.SubItems(13) = Null2String(!tech2)
            Item.SubItems(14) = Null2String(!tech3)
            Item.SubItems(15) = Null2String(!savedate)
            Item.SubItems(16) = Null2String(!savetime)
            .MoveNext
        Loop
    End With


    Set RS = Nothing
    ListRo.Refresh

    Exit Sub

loaderror:

    MsgBox Err.Description

End Sub

Sub FillVehicle()
    Dim RS                              As New ADODB.Recordset
    Dim SQL                             As String
    Dim Item                            As ListItem
    Dim CUSCDE                          As String
    Dim cnt                             As String

    CUSCDE = TheCuscde
    'COMMENT BY : MJP 09-12-2007 04:04 PM -----------------------------------
    'DESCRIPTION : IT LOCKS THE LISTVIEW THATS WHY SA CANNOT SCROLL THE SCROLL BAR TO SEE THE VEHICLE INFO.
    'listCusveh.Enabled = False
    'COMMENT BY : MJP 09-12-2007 04:04 PM -----------------------------------

    SQL = "SELECT * FROM CSMS_cusveh WHERE cuscde='" & CUSCDE & "'"
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
            Set Item = listCusveh.ListItems.Add(, , cnt)
            Item.SubItems(1) = Null2String(!NIYM)
            Item.SubItems(2) = Null2String(!Vin)
            Item.SubItems(3) = Null2String(!PLATE_NO)
            Item.SubItems(4) = Null2String(!VCOND_NO)
            Item.SubItems(5) = Null2String(!ClrCde)
            Item.SubItems(6) = Null2String(!YER)
            Item.SubItems(7) = Null2String(!Make)
            Item.SubItems(8) = Null2String(!Model)
            Item.SubItems(9) = Null2String(!Engine)
            Item.SubItems(10) = Null2String(!KMReading)
            Item.SubItems(11) = Null2String(!ProdNo)
            Item.SubItems(12) = Null2String(!Serial)
            Item.SubItems(13) = Null2String(!TIN_Number)
            Item.SubItems(14) = Null2String(!D_Sold)
            '            Item.SubItems(15) = Null2String(!xvar_cert)
            Item.SubItems(16) = Null2String(!Del_Date)
            '            Item.SubItems(17) = Null2String(!selling_deadler)
            .MoveNext
        Loop
    End With

    Set RS = Nothing
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    cmdTrans_Click
End Sub

Private Sub ListRo_ItemClick(ByVal Item As MSComctlLib.ListItem)

    TheEstimate = ListRo.ListItems.Item(ListRo.SelectedItem.Index)
    TheEstimate = ListRo.SelectedItem.SubItems(3)

    TheOR = ListRo.ListItems.Item(ListRo.SelectedItem.Index)
    TheOR = ListRo.SelectedItem.SubItems(2)

    Call FillJob
    Call FillParts
    Call FillMaterial

    'UPDATE BY : MJP 09-12-2007 04:15 PM ---------------------------------------------------
    'DESCRIPTION : TO DISPLAY THE ACCESSORIES ISSUE TO THAT REPAIR ORDER
    Call FillAccessories
    'UPDATE BY : MJP 09-12-2007 04:15 PM ---------------------------------------------------

    If OtherTab(1).Selected = True Then
        OtherTab(0).Selected = True
    End If
End Sub

'UPDATE BY : MJP 09-12-2007 04:15 PM ---------------------------------------------------
'DESCRIPTION : TO DISPLAY THE ACCESSORIES ISSUE TO THAT REPAIR ORDER
Sub FillAccessories()
    Dim rsTmp                           As New ADODB.Recordset
    Dim Item                            As ListItem

    Set rsTmp = gconDMIS.Execute("Select * From CSMS_ro_det Where Rep_OR = '" & ListRo.SelectedItem.ListSubItems(2).Text & "' And Livil = '4' Order By Line_NO")
    lsvACC.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = lsvACC.ListItems.Add(, , Null2String(rsTmp!Line_No))
            Item.SubItems(1) = Null2String(rsTmp!DETCDE)
            Item.SubItems(2) = Null2String(rsTmp!DETDSC)
            Item.SubItems(3) = Null2String(rsTmp!detvol)
            Item.SubItems(4) = Format(Null2String(rsTmp!DetPrc), "#,###,##0.00")
            Item.SubItems(5) = Format(Null2String(rsTmp!DetCost), "#,###,##0.00")
            Item.SubItems(6) = Format(Item.SubItems(3) * Item.SubItems(4), "#,###,##0.00")
            Item.SubItems(7) = Format((Item.SubItems(6) / 100) * CDbl(Null2String(rsTmp!discrate)), "#,###,##0.00")
            Item.SubItems(8) = Format(Item.SubItems(6) - Item.SubItems(7), "#,###,##0.00")

            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
End Sub
'UPDATE BY : MJP 09-12-2007 04:15 PM ---------------------------------------------------

Private Sub lstActiveCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next

    TheCuscde = lstActiveCustomer.ListItems.Item(lstActiveCustomer.SelectedItem.Index)
    TheCuscde = lstActiveCustomer.SelectedItem.SubItems(1)

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

End Sub



Sub Myrefresh()
    ListRo.ListItems.Clear
    Listjob.ListItems.Clear
    listCusveh.ListItems.Clear
    lstAllCustomer.ListItems.Clear
    lstActiveCustomer.ListItems.Clear
End Sub

Private Sub lstAllCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'UDPATING CODE    :AXP-5312007-A
    TheCuscde = lstAllCustomer.ListItems.Item(lstAllCustomer.SelectedItem.Index)
    TheCuscde = lstAllCustomer.SelectedItem.SubItems(1)
    Call FillVehicle
    Call FillRO


    If ListRo.ListItems.Count > 0 Then
        ListRo.ListItems(1).Selected = True
        ListRo.ListItems(1).EnsureVisible
        ListRo_ItemClick ListRo.SelectedItem
    Else
        ListRo.ListItems.Clear
        Listjob.ListItems.Clear
        'listCusveh.ListItems.Clear

    End If

End Sub

Private Sub Otp_Click(Index As Integer)
    On Error Resume Next
    txtSearchKey_All.SetFocus
End Sub

Private Sub txtSearchKey_Active_Change()
    Call Fill_ActiveCustomerSearch
    If lstActiveCustomer.ListItems.Count > 0 Then
        lstActiveCustomer.ListItems(1).Selected = True
        lstActiveCustomer.ListItems(1).EnsureVisible
        lstActiveCustomer_ItemClick lstActiveCustomer.SelectedItem
    Else
        ListRo.ListItems.Clear
        Listjob.ListItems.Clear
        listCusveh.ListItems.Clear
    End If

End Sub

Private Sub txtSearchKey_All_Change()
    Fill_AllCustomerSearch
    If lstAllCustomer.ListItems.Count > 0 Then
        lstAllCustomer.ListItems(1).Selected = True
        lstAllCustomer.ListItems(1).EnsureVisible
        lstAllCustomer_ItemClick lstActiveCustomer.SelectedItem
    Else
        ListRo.ListItems.Clear
        Listjob.ListItems.Clear
        'listCusveh.ListItems.Clear
    End If
End Sub

Private Sub Wotp_Click(Index As Integer)
    On Error Resume Next
    txtSearchKey_Active.SetFocus
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmCSMSEstiSearchCustomer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estimate Search Customer"
   ClientHeight    =   6240
   ClientLeft      =   2835
   ClientTop       =   3390
   ClientWidth     =   10785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "EstiSearchCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10785
   Begin XtremeSuiteControls.TabControl SearchTab 
      Height          =   6195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _Version        =   655364
      _ExtentX        =   19076
      _ExtentY        =   10927
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   130
      ItemCount       =   6
      Item(0).Caption =   "By  &Customer Name"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "By &Estimate No"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "By &Invoice Number"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Item(3).Caption =   "By &Vehicle Model"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage4"
      Item(4).Caption =   "By &Service Adviser"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "TabControlPage5"
      Item(5).Caption =   "By &Plate Number"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "tbPlateNo"
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   5580
         Left            =   -69970
         TabIndex        =   2
         Top             =   585
         Visible         =   0   'False
         Width           =   10755
         _Version        =   655364
         _ExtentX        =   18971
         _ExtentY        =   9842
         _StockProps     =   0
         Begin VB.TextBox txtVehicleModel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1350
            TabIndex        =   13
            Top             =   60
            Width           =   9315
         End
         Begin MSComctlLib.ListView ListVehicleModel 
            Height          =   5055
            Left            =   0
            TabIndex        =   12
            Top             =   480
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "EstiSearchCustomer.frx":000C
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VEHICLE MODEL"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ESTIMATE NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "PLATE NUMBER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage5 
         Height          =   5580
         Left            =   -69970
         TabIndex        =   1
         Top             =   585
         Visible         =   0   'False
         Width           =   10755
         _Version        =   655364
         _ExtentX        =   18971
         _ExtentY        =   9842
         _StockProps     =   0
         Begin VB.TextBox txtServiceAdviser 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1350
            TabIndex        =   14
            Top             =   60
            Width           =   9315
         End
         Begin MSComctlLib.ListView ListServiceAdviser 
            Height          =   5055
            Left            =   0
            TabIndex        =   15
            Top             =   480
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "EstiSearchCustomer.frx":0326
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. ADVISER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ESTIMATE NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "PLATE NUMBER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   23
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   5580
         Left            =   -69970
         TabIndex        =   3
         Top             =   585
         Visible         =   0   'False
         Width           =   10755
         _Version        =   655364
         _ExtentX        =   18971
         _ExtentY        =   9842
         _StockProps     =   0
         Begin VB.TextBox txtInvoiceNumber 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1350
            TabIndex        =   10
            Top             =   60
            Width           =   9315
         End
         Begin MSComctlLib.ListView ListInvoiceNumber 
            Height          =   5055
            Left            =   0
            TabIndex        =   11
            Top             =   480
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "EstiSearchCustomer.frx":0640
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ESTIMATE NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   5580
         Left            =   -69970
         TabIndex        =   4
         Top             =   585
         Visible         =   0   'False
         Width           =   10755
         _Version        =   655364
         _ExtentX        =   18971
         _ExtentY        =   9842
         _StockProps     =   0
         Begin VB.TextBox txtEstimateNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1350
            TabIndex        =   8
            Top             =   60
            Width           =   9315
         End
         Begin MSComctlLib.ListView ListEstimateNo 
            Height          =   5055
            Left            =   0
            TabIndex        =   9
            Top             =   480
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "EstiSearchCustomer.frx":095A
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ESTIMATE NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   19
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5580
         Left            =   30
         TabIndex        =   5
         Top             =   585
         Width           =   10755
         _Version        =   655364
         _ExtentX        =   18971
         _ExtentY        =   9842
         _StockProps     =   0
         Begin VB.TextBox txtCustomerName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1350
            TabIndex        =   6
            Top             =   60
            Width           =   9315
         End
         Begin MSComctlLib.ListView ListCustomerName 
            Height          =   5055
            Left            =   0
            TabIndex        =   7
            Top             =   480
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "EstiSearchCustomer.frx":0C74
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6526
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ESTIMATE NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   20
            Top             =   90
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPlateNo 
         Height          =   5580
         Left            =   -69970
         TabIndex        =   16
         Top             =   585
         Visible         =   0   'False
         Width           =   10755
         _Version        =   655364
         _ExtentX        =   18971
         _ExtentY        =   9842
         _StockProps     =   0
         Begin VB.TextBox txtPlateNumber 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1350
            TabIndex        =   17
            Top             =   60
            Width           =   9315
         End
         Begin MSComctlLib.ListView ListPlateNumber 
            Height          =   5055
            Left            =   0
            TabIndex        =   18
            Top             =   480
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
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
            MouseIcon       =   "EstiSearchCustomer.frx":0F8E
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ESTIMATE NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   90
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frmCSMSEstiSearchCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEsti_HD                                          As New ADODB.Recordset
Dim y                                                  As Long
Dim k                                                  As Long

Sub clearListView()
    For y = 1 To Me.ListCustomerName.ListItems.Count
        If Me.ListCustomerName.ListItems.Count <= 0 Then Exit For
        Me.ListCustomerName.Sorted = False
        Me.ListCustomerName.ListItems.Remove Me.ListCustomerName.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListEstimateNo.ListItems.Count
        If Me.ListEstimateNo.ListItems.Count <= 0 Then Exit For
        Me.ListEstimateNo.Sorted = False
        Me.ListEstimateNo.ListItems.Remove Me.ListEstimateNo.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListInvoiceNumber.ListItems.Count
        If Me.ListInvoiceNumber.ListItems.Count <= 0 Then Exit For
        Me.ListInvoiceNumber.Sorted = False
        Me.ListInvoiceNumber.ListItems.Remove Me.ListInvoiceNumber.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListPlateNumber.ListItems.Count
        If Me.ListPlateNumber.ListItems.Count <= 0 Then Exit For
        Me.ListPlateNumber.Sorted = False
        Me.ListPlateNumber.ListItems.Remove Me.ListPlateNumber.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListVehicleModel.ListItems.Count
        If Me.ListVehicleModel.ListItems.Count <= 0 Then Exit For
        Me.ListVehicleModel.Sorted = False
        Me.ListVehicleModel.ListItems.Remove Me.ListVehicleModel.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListServiceAdviser.ListItems.Count
        If Me.ListServiceAdviser.ListItems.Count <= 0 Then Exit For
        Me.ListServiceAdviser.Sorted = False
        Me.ListServiceAdviser.ListItems.Remove Me.ListServiceAdviser.SelectedItem.INDEX
    Next y
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Select Case SEARCH_TAB
            Case 0:
                If Trim(txtCustomerName) <> "" Then
                    On Error Resume Next
                    txtCustomerName.SetFocus
                Else
                    Unload Me
                End If
            Case 1:
                If Trim(txtEstimateno) <> "" Then
                    On Error Resume Next
                    txtEstimateno.SetFocus
                Else
                    Unload Me
                End If
            Case 2:
                If Trim(txtInvoiceNumber) <> "" Then
                    On Error Resume Next
                    txtInvoiceNumber.SetFocus
                Else
                    Unload Me
                End If
            Case 3:
                If Trim(txtVehicleModel) <> "" Then
                    On Error Resume Next
                    txtVehicleModel.SetFocus
                Else
                    Unload Me
                End If
            Case 4:
                If Trim(txtServiceAdviser) <> "" Then
                    On Error Resume Next
                    txtServiceAdviser.SetFocus
                Else
                    Unload Me
                End If
            Case 5:
                If Trim(txtPlateNumber) <> "" Then
                    On Error Resume Next
                    txtPlateNumber.SetFocus
                Else
                    Unload Me
                End If
        End Select
    End If
    If Shift = 2 Then
        On Error GoTo Errorcode:
        Select Case KeyCode
            Case vbKeyC: SearchTab.SelectedItem = 0
            Case vbKeyE: SearchTab.SelectedItem = 1
            Case vbKeyI: SearchTab.SelectedItem = 2
            Case vbKeyV: SearchTab.SelectedItem = 3
            Case vbKeyS: SearchTab.SelectedItem = 4
            Case vbKeyP: SearchTab.SelectedItem = 5
        End Select
        SEARCH_TAB = SearchTab.Selected.INDEX
        SearchTab.SelectedItem = SEARCH_TAB
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If SEARCH_TAB = 0 Then txtCustomerName.Text = SEARCHCUSTOMERNAME
End Sub

Private Sub ListCustomerName_DblClick()
    SEARCHCUSTOMERNAME = txtCustomerName.Text
    frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListCustomerName.SelectedItem.SubItems(1)))
    Unload Me
End Sub

Private Sub ListCustomerName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtCustomerName.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListCustomerName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SEARCHCUSTOMERNAME = txtCustomerName.Text
        frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListCustomerName.SelectedItem.SubItems(1)))
        Unload Me
    End If
End Sub

Private Sub ListInvoiceNumber_DblClick()
    frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListInvoiceNumber.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListPlateNumber_DblClick()
    frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListPlateNumber.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListEstimateNo_DblClick()
    frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListEstimateNo.SelectedItem))
    Unload Me
End Sub

Private Sub ListEstimateNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtEstimateno.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListEstimateNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListEstimateNo.SelectedItem))
        Unload Me
    End If
End Sub

Private Sub ListInvoiceNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtInvoiceNumber.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListInvoiceNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListInvoiceNumber.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub ListPlateNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtPlateNumber.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListPlateNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListPlateNumber.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub ListServiceAdviser_DblClick()
    frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListServiceAdviser.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListVehicleModel_DblClick()
    frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListVehicleModel.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListVehicleModel_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtVehicleModel.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListVehicleModel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListVehicleModel.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub ListServiceAdviser_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtServiceAdviser.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListServiceAdviser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSEstimateEntry.SearchEstimateNo (Trim(Me.ListServiceAdviser.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub SearchTab_Click(PreviousTab As Integer)
    SEARCH_TAB = SearchTab.Tab
    DoEvents
    txtCustomerName.Enabled = False: txtEstimateno.Enabled = False
    txtInvoiceNumber.Enabled = False: txtPlateNumber.Enabled = False
    txtVehicleModel.Enabled = False: txtServiceAdviser.Enabled = False
    ListCustomerName.Enabled = False: ListEstimateNo.Enabled = False
    ListInvoiceNumber.Enabled = False: ListPlateNumber.Enabled = False
    ListVehicleModel.Enabled = False: ListServiceAdviser.Enabled = False
    Select Case SEARCH_TAB
        Case 0
            txtCustomerName.Enabled = True: ListCustomerName.Enabled = True
            Me.Caption = "Search Item by Customer Name"
            On Error Resume Next
            txtCustomerName.SetFocus
        Case 1
            txtEstimateno.Enabled = True: ListEstimateNo.Enabled = True
            Me.Caption = "Search Item by Estimate Number"
            On Error Resume Next
            txtEstimateno.SetFocus
        Case 2
            txtInvoiceNumber.Enabled = True: ListInvoiceNumber.Enabled = True
            Me.Caption = "Search Item by Invoice Number"
            On Error Resume Next
            txtInvoiceNumber.SetFocus
        Case 3
            txtPlateNumber.Enabled = True: ListPlateNumber.Enabled = True
            Me.Caption = "Search Item by Plate Number Order"
            On Error Resume Next
            txtPlateNumber.SetFocus
        Case 4
            txtVehicleModel.Enabled = True: ListVehicleModel.Enabled = True
            Me.Caption = "Search Item by Vehicle Model"
            On Error Resume Next
            txtVehicleModel.SetFocus
        Case 5
            txtServiceAdviser.Enabled = True: ListServiceAdviser.Enabled = True
            Me.Caption = "Search Item by Service Adviser"
            On Error Resume Next
            txtServiceAdviser.SetFocus
    End Select
End Sub

Private Sub txtCustomerName_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCustomerName.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListCustomerName.Enabled = True And ListCustomerName.ListItems.Count > 0 Then
            ListCustomerName.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCustomerName_Change()
    If txtCustomerName = "" Then
        ListCustomerName.Enabled = False
        Me.ListCustomerName.Sorted = False: Me.ListCustomerName.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 niym,estimateno,invoice,plate_no,model,recd_by,ro_amount from CSMS_Esti_Hd order by niym asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListCustomerName.ListItems, rsEsti_HD
            ListCustomerName.Enabled = True
        End If
    Else
        Me.ListCustomerName.Sorted = False: Me.ListCustomerName.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 niym,estimateno,invoice,plate_no,model,recd_by,ro_amount from CSMS_Esti_Hd Where niym like '" & Trim(Replace(Me.txtCustomerName, "'", "")) & "%' order by niym asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListCustomerName.ListItems, rsEsti_HD
            ListCustomerName.Enabled = True
        End If
    End If
End Sub

Private Sub txtEstimateNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtEstimateno.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListEstimateNo.Enabled = True And ListEstimateNo.ListItems.Count > 0 Then
            ListEstimateNo.SetFocus
        End If

    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtEstimateNo_Change()
    If txtEstimateno = "" Then
        ListEstimateNo.Enabled = False
        Me.ListEstimateNo.Sorted = False: Me.ListEstimateNo.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 estimateno,niym,invoice,plate_no,model,recd_by,ro_amount from CSMS_Esti_Hd order by estimateno asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListEstimateNo.ListItems, rsEsti_HD
            ListEstimateNo.Enabled = True
        End If
    Else
        Dim EstimateNo                                 As String
        EstimateNo = UCase(txtEstimateno.Text)
        Me.ListEstimateNo.Sorted = False: Me.ListEstimateNo.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 estimateno,niym,invoice,plate_no,model,recd_by,ro_amount from CSMS_Esti_Hd Where estimateno like '" & Replace(EstimateNo, "'", "") & "%' order by estimateno asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListEstimateNo.ListItems, rsEsti_HD
            ListEstimateNo.Enabled = True
        End If
    End If
End Sub

Private Sub txtInvoiceNumber_Change()
    If txtInvoiceNumber = "" Then
        ListInvoiceNumber.Enabled = False
        Me.ListInvoiceNumber.Sorted = False: Me.ListInvoiceNumber.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 invoice,niym,estimateno,plate_no,model,recd_by,ro_amount from CSMS_Esti_Hd order by invoice asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListInvoiceNumber.ListItems, rsEsti_HD
            ListInvoiceNumber.Enabled = True
        End If
    Else
        Dim InvoiceNumber, InvoiceNumber2, InvoiceNumber3 As String
        InvoiceNumber = UCase(txtInvoiceNumber.Text)
        If InvoiceNumber <> "" Then
            If IsNumeric(InvoiceNumber) = True Then
                InvoiceNumber = Format(Right(InvoiceNumber, 6), "000000")
            Else
                For k = 1 To Len(InvoiceNumber)
                    InvoiceNumber2 = Mid(InvoiceNumber, k, 1)
                    If IsNumeric(InvoiceNumber2) = True Then InvoiceNumber3 = InvoiceNumber3 + InvoiceNumber2
                Next
                InvoiceNumber = Format(InvoiceNumber3, "000000")
            End If
        End If
        If IsNumeric(InvoiceNumber) = True Then
            Me.ListInvoiceNumber.Sorted = False: Me.ListInvoiceNumber.ListItems.Clear
            Set rsEsti_HD = New ADODB.Recordset
            Set rsEsti_HD = gconDMIS.Execute("select top 100 invoice,niym,estimateno,plate_no,model,recd_by,ro_amount from CSMS_Esti_Hd Where invoice like'" & InvoiceNumber & "%' order by invoice asc")
            If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
                Listview_Loadval Me.ListInvoiceNumber.ListItems, rsEsti_HD
                ListInvoiceNumber.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub txtPlateNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtPlateNumber.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListPlateNumber.Enabled = True And ListPlateNumber.ListItems.Count > 0 Then
            ListPlateNumber.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtPlateNumber_Change()
    If txtPlateNumber = "" Then
        ListPlateNumber.Enabled = False
        Me.ListPlateNumber.Sorted = False: Me.ListPlateNumber.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 plate_no,niym,estimateno,invoice,model,recd_by,ro_amount from CSMS_Esti_Hd order by plate_no asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListPlateNumber.ListItems, rsEsti_HD
            ListPlateNumber.Enabled = True
        End If
    Else
        Me.ListPlateNumber.Sorted = False: Me.ListPlateNumber.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 plate_no,niym,estimateno,invoice,model,recd_by,ro_amount from CSMS_Esti_Hd Where plate_no like '" & Trim(Replace(Me.txtPlateNumber, "'", "")) & "%' order by plate_no asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListPlateNumber.ListItems, rsEsti_HD
            ListPlateNumber.Enabled = True
        End If
    End If
End Sub

Private Sub txtVehicleModel_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtVehicleModel.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListVehicleModel.Enabled = True And ListVehicleModel.ListItems.Count > 0 Then
            ListVehicleModel.SetFocus
        End If

    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVehicleModel_Change()
    If txtVehicleModel = "" Then
        Me.ListVehicleModel.Sorted = False: Me.ListVehicleModel.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 model,niym,estimateno,invoice,plate_no,recd_by,ro_amount from CSMS_Esti_Hd order by model asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListVehicleModel.ListItems, rsEsti_HD
        End If
    Else
        Me.ListVehicleModel.Sorted = False: Me.ListVehicleModel.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 model,niym,estimateno,invoice,plate_no,recd_by,ro_amount from CSMS_Esti_Hd Where model like '" & Trim(Replace(Me.txtVehicleModel, "'", "")) & "%' order by model asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListVehicleModel.ListItems, rsEsti_HD
        End If
    End If
End Sub

Private Sub txtServiceAdviser_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtServiceAdviser.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListServiceAdviser.Enabled = True And ListServiceAdviser.ListItems.Count > 0 Then
            ListServiceAdviser.SetFocus
        End If

    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtServiceAdviser_Change()
    If txtServiceAdviser = "" Then
        ListServiceAdviser.Enabled = False
        Me.ListServiceAdviser.Sorted = False: Me.ListServiceAdviser.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 recd_by,niym,estimateno,invoice,plate_no,model,ro_amount from CSMS_Esti_Hd Where recd_by like '" & Trim(Me.txtServiceAdviser) & "%' order by recd_by asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListServiceAdviser.ListItems, rsEsti_HD
            ListServiceAdviser.Enabled = True
        End If
    Else
        Me.ListServiceAdviser.Sorted = False: Me.ListServiceAdviser.ListItems.Clear
        Set rsEsti_HD = New ADODB.Recordset
        Set rsEsti_HD = gconDMIS.Execute("select top 100 recd_by,niym,estimateno,invoice,plate_no,model,ro_amount from CSMS_Esti_Hd order by recd_by asc")
        If Not (rsEsti_HD.EOF And rsEsti_HD.BOF) Then
            Listview_Loadval Me.ListServiceAdviser.ListItems, rsEsti_HD
            ListServiceAdviser.Enabled = True
        End If
    End If
End Sub


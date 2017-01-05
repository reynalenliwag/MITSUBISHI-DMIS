VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSQualityInformation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Car Quality information Report"
   ClientHeight    =   10935
   ClientLeft      =   150
   ClientTop       =   525
   ClientWidth     =   11010
   Icon            =   "frmCSMSQuailityInformation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCSMSQuailityInformation.frx":27A2
   ScaleHeight     =   10935
   ScaleWidth      =   11010
   Begin VB.PictureBox picdelete 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   11100
      ScaleHeight     =   3825
      ScaleWidth      =   6075
      TabIndex        =   226
      Top             =   4020
      Width           =   6105
      Begin VB.CommandButton cmdx 
         Caption         =   "X"
         Height          =   225
         Left            =   5700
         TabIndex        =   232
         Top             =   30
         Width           =   345
      End
      Begin VB.CommandButton cmdprintOk 
         Caption         =   "Ok"
         Height          =   405
         Left            =   4740
         TabIndex        =   231
         Top             =   3330
         Width           =   1215
      End
      Begin MSComctlLib.ListView listpartoption 
         Height          =   2970
         Left            =   6990
         TabIndex        =   227
         Top             =   1680
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   5239
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         MouseIcon       =   "frmCSMSQuailityInformation.frx":4F44
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Part No"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Part name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "QTY"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cost"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "LTS"
            Object.Width           =   882
         EndProperty
      End
      Begin MSComctlLib.ListView isdeletelist 
         Height          =   2925
         Left            =   90
         TabIndex        =   228
         Top             =   330
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   5159
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Part No"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Part name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "QTY"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cost"
            Object.Width           =   1586
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "LTS"
            Object.Width           =   882
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   315
         Left            =   -90
         TabIndex        =   230
         Top             =   -30
         Width           =   6165
         _Version        =   655364
         _ExtentX        =   10874
         _ExtentY        =   556
         _StockProps     =   14
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
   Begin VB.PictureBox thepicControl 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11010
      TabIndex        =   181
      Top             =   10110
      Width           =   11010
      Begin VB.CommandButton cmdexit 
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
         Height          =   735
         Left            =   10260
         MouseIcon       =   "frmCSMSQuailityInformation.frx":50A6
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":51F8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
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
         Height          =   735
         Left            =   9540
         MouseIcon       =   "frmCSMSQuailityInformation.frx":555E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":56B0
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
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
         Height          =   735
         Left            =   8820
         MouseIcon       =   "frmCSMSQuailityInformation.frx":5A16
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":5B68
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdedit 
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
         Height          =   735
         Left            =   8100
         MouseIcon       =   "frmCSMSQuailityInformation.frx":5E93
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":5FE5
         Style           =   1  'Graphical
         TabIndex        =   189
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdadd 
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
         Height          =   735
         Left            =   7380
         MouseIcon       =   "frmCSMSQuailityInformation.frx":6341
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":6493
         Style           =   1  'Graphical
         TabIndex        =   186
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Last"
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
         Left            =   6660
         MouseIcon       =   "frmCSMSQuailityInformation.frx":67A6
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":68F8
         Style           =   1  'Graphical
         TabIndex        =   187
         ToolTipText     =   "Move to Last Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdlast 
         Caption         =   "First"
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
         Left            =   5940
         MouseIcon       =   "frmCSMSQuailityInformation.frx":6C48
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":6D9A
         Style           =   1  'Graphical
         TabIndex        =   188
         ToolTipText     =   "Move to First Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdfind 
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
         Height          =   735
         Left            =   5220
         MouseIcon       =   "frmCSMSQuailityInformation.frx":70F8
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":724A
         Style           =   1  'Graphical
         TabIndex        =   185
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdnext 
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
         Height          =   735
         Left            =   4500
         MouseIcon       =   "frmCSMSQuailityInformation.frx":7544
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":7696
         Style           =   1  'Graphical
         TabIndex        =   182
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdprev 
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
         Height          =   735
         Left            =   3780
         MouseIcon       =   "frmCSMSQuailityInformation.frx":79EE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":7B40
         Style           =   1  'Graphical
         TabIndex        =   183
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.Label lblID 
         Caption         =   "theID"
         Height          =   375
         Left            =   3765
         TabIndex        =   184
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.PictureBox thepicSave 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   11010
      TabIndex        =   178
      Top             =   9315
      Width           =   11010
      Begin VB.CommandButton cmdCancel 
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
         Height          =   735
         Left            =   10230
         MouseIcon       =   "frmCSMSQuailityInformation.frx":7E9F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":7FF1
         Style           =   1  'Graphical
         TabIndex        =   179
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdsave 
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
         Height          =   735
         Left            =   9510
         MouseIcon       =   "frmCSMSQuailityInformation.frx":832F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSQuailityInformation.frx":8481
         Style           =   1  'Graphical
         TabIndex        =   180
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox Pictops 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2925
      ScaleWidth      =   11010
      TabIndex        =   0
      Top             =   0
      Width           =   11010
      Begin VB.OptionButton otpRequestN 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   15
         Width           =   645
      End
      Begin VB.OptionButton otpRequestY 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1500
         TabIndex        =   9
         Top             =   15
         Width           =   615
      End
      Begin VB.Frame frameA 
         Height          =   2670
         Left            =   30
         TabIndex        =   12
         Top             =   255
         Width           =   5835
         Begin Crystal.CrystalReport rptCQIRReport 
            Left            =   4245
            Top             =   285
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.OptionButton OtpTransTypeAuto 
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2550
            TabIndex        =   29
            Top             =   2115
            Width           =   735
         End
         Begin VB.OptionButton OtpTransTypeManual 
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1635
            TabIndex        =   28
            Top             =   2115
            Width           =   945
         End
         Begin VB.TextBox TxtPWANo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Top             =   120
            Width           =   1635
         End
         Begin VB.TextBox txtPWAType 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            Top             =   405
            Width           =   1635
         End
         Begin VB.TextBox txtVIN 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            TabIndex        =   19
            Top             =   960
            Width           =   3765
         End
         Begin VB.TextBox TxtCustomer 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            TabIndex        =   21
            Top             =   1245
            Width           =   3780
         End
         Begin VB.TextBox txtEngineNo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            TabIndex        =   23
            Top             =   1530
            Width           =   3765
         End
         Begin VB.TextBox txtTMAxleNo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            TabIndex        =   26
            Top             =   1815
            Width           =   3765
         End
         Begin VB.CheckBox checkPhoto 
            Caption         =   "Photo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1665
            TabIndex        =   32
            Top             =   2355
            Width           =   855
         End
         Begin VB.CheckBox checkSamplePart 
            Caption         =   "Sample Part"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2550
            TabIndex        =   31
            Top             =   2340
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker DeviverDate 
            Height          =   255
            Left            =   1680
            TabIndex        =   190
            Top             =   690
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            Format          =   51773441
            CurrentDate     =   39283
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   30
            Left            =   1860
            TabIndex        =   229
            Top             =   2670
            Width           =   2925
            _Version        =   655364
            _ExtentX        =   5159
            _ExtentY        =   53
            _StockProps     =   14
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PWA No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   810
            TabIndex        =   13
            Top             =   165
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Delevery Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   18
            Top             =   735
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "VIN No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   810
            TabIndex        =   20
            Top             =   1020
            Width           =   795
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Engine No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   405
            TabIndex        =   24
            Top             =   1605
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   765
            TabIndex        =   22
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TM/Axle No:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   840
            TabIndex        =   25
            Top             =   1860
            Width           =   795
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transmission type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -75
            TabIndex        =   27
            Top             =   2130
            Width           =   1665
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Attachments"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   510
            TabIndex        =   30
            Top             =   2370
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PWA Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   450
            TabIndex        =   16
            Top             =   435
            Width           =   1185
         End
      End
      Begin VB.Frame FrameB 
         Height          =   2655
         Left            =   5940
         TabIndex        =   33
         Top             =   270
         Width           =   4725
         Begin VB.TextBox txtMileage 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1515
            TabIndex        =   197
            Top             =   1350
            Width           =   1215
         End
         Begin VB.TextBox txtsaleAdvisor 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1515
            TabIndex        =   196
            Top             =   1695
            Width           =   2430
         End
         Begin VB.ComboBox Cbotech 
            Height          =   315
            Left            =   1515
            TabIndex        =   195
            Top             =   2025
            Width           =   2460
         End
         Begin VB.TextBox txtRO 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   194
            Top             =   1050
            Width           =   1965
         End
         Begin VB.TextBox txtDealer 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1515
            TabIndex        =   193
            Top             =   120
            Width           =   1920
         End
         Begin VB.OptionButton OtpSentCardN 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2235
            TabIndex        =   47
            Top             =   2340
            Width           =   645
         End
         Begin VB.OptionButton OtpSentCardY 
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1530
            TabIndex        =   46
            Top             =   2340
            Width           =   615
         End
         Begin VB.TextBox txtplateNo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3540
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   1350
            Width           =   1110
         End
         Begin MSComCtl2.DTPicker RepairDate 
            Height          =   285
            Left            =   1515
            TabIndex        =   191
            Top             =   750
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   51773441
            CurrentDate     =   39283
         End
         Begin MSComCtl2.DTPicker InspectionDate 
            Height          =   315
            Left            =   1515
            TabIndex        =   192
            Top             =   420
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   51773441
            CurrentDate     =   39283
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Repair Order No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   39
            Top             =   1065
            Width           =   1185
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Technician"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   180
            TabIndex        =   44
            Top             =   2025
            Width           =   1245
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Service Advisor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   43
            Top             =   1695
            Width           =   1395
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sent Reg Card"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   45
            Top             =   2355
            Width           =   1395
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Plate No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2730
            TabIndex        =   41
            Top             =   1380
            Width           =   765
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mileage(KM)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   40
            Top             =   1380
            Width           =   1125
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Repair Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   38
            Top             =   765
            Width           =   1035
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inspection date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   35
            Top             =   435
            Width           =   1005
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dealer Code"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   750
            TabIndex        =   34
            Top             =   135
            Width           =   675
         End
      End
      Begin VB.TextBox txtReferenceNo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8190
         TabIndex        =   8
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DLR CQIR Reference #"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6120
         TabIndex        =   11
         Top             =   30
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PWA REQUEST"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   7
         Top             =   45
         Width           =   1395
      End
   End
   Begin VB.PictureBox picunder 
      Align           =   1  'Align Top
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7875
      ScaleWidth      =   10950
      TabIndex        =   56
      Top             =   2925
      Width           =   11010
      Begin VB.PictureBox Picover 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   7665
         Left            =   60
         ScaleHeight     =   7665
         ScaleWidth      =   10515
         TabIndex        =   57
         Top             =   15
         Width           =   10515
         Begin VB.TextBox txtpassSublet 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   239
            Top             =   5550
            Width           =   1485
         End
         Begin XtremeSuiteControls.TabControl TabControl1 
            Height          =   1635
            Left            =   -30
            TabIndex        =   220
            Top             =   3090
            Width           =   6015
            _Version        =   655364
            _ExtentX        =   10610
            _ExtentY        =   2884
            _StockProps     =   64
            Appearance      =   3
            Color           =   8
            PaintManager.BoldSelected=   -1  'True
            PaintManager.DisableLunaColors=   0   'False
            PaintManager.FixedTabWidth=   100
            PaintManager.MaxTabWidth=   100
            PaintManager.MinTabWidth=   100
            ItemCount       =   2
            Item(0).Caption =   "Parts"
            Item(0).ControlCount=   1
            Item(0).Control(0)=   "listpart"
            Item(1).Caption =   "Sublet"
            Item(1).ControlCount=   1
            Item(1).Control(0)=   "listsublet"
            Begin MSComctlLib.ListView listpart 
               Height          =   1230
               Left            =   30
               TabIndex        =   221
               Top             =   330
               Width           =   5925
               _ExtentX        =   10451
               _ExtentY        =   2170
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
               MouseIcon       =   "frmCSMSQuailityInformation.frx":87D1
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Part No"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Part name"
                  Object.Width           =   4410
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "QTY"
                  Object.Width           =   970
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Cost"
                  Object.Width           =   1587
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Code"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "LTS"
                  Object.Width           =   882
               EndProperty
            End
            Begin MSComctlLib.ListView listsublet 
               Height          =   1155
               Left            =   -69940
               TabIndex        =   222
               Top             =   420
               Visible         =   0   'False
               Width           =   5835
               _ExtentX        =   10292
               _ExtentY        =   2037
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Description"
                  Object.Width           =   7056
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Charge"
                  Object.Width           =   2469
               EndProperty
            End
         End
         Begin VB.CommandButton cmdsublet 
            Caption         =   ".."
            Height          =   345
            Left            =   3420
            TabIndex        =   219
            ToolTipText     =   "Add Sublet"
            Top             =   5550
            Width           =   375
         End
         Begin VB.TextBox txttotalLTS 
            Alignment       =   1  'Right Justify
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
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   4980
            Width           =   1470
         End
         Begin VB.TextBox txttotalcost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4380
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   4740
            Width           =   975
         End
         Begin VB.PictureBox Picbottoms 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   0
            ScaleHeight     =   540
            ScaleWidth      =   10395
            TabIndex        =   170
            Top             =   6975
            Width           =   10395
            Begin VB.TextBox txtRequested 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   315
               TabIndex        =   172
               Top             =   30
               Width           =   2295
            End
            Begin VB.TextBox txtgeneralManager 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3810
               TabIndex        =   173
               Top             =   30
               Width           =   2655
            End
            Begin VB.TextBox txtServiceDept 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   7830
               TabIndex        =   171
               Top             =   15
               Width           =   2295
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Dealer Warranty Administrator"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   174
               Top             =   285
               Width           =   3075
            End
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Dealer General Manager"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3735
               TabIndex        =   175
               Top             =   300
               Width           =   2595
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "HARI SERVICE DEPARTMENT"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   7575
               TabIndex        =   176
               Top             =   300
               Width           =   2745
            End
         End
         Begin VB.Frame FrameD 
            Caption         =   "Condition"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6960
            Left            =   6000
            TabIndex        =   93
            Top             =   0
            Width           =   4485
            Begin VB.Frame Frame1 
               Caption         =   "Vehicle Maintenance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   120
               TabIndex        =   157
               Top             =   5025
               Width           =   4230
               Begin VB.OptionButton maintenance 
                  Caption         =   "Dealer"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   75
                  TabIndex        =   162
                  Top             =   585
                  Width           =   885
               End
               Begin VB.OptionButton maintenance 
                  Caption         =   "3-star shop"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   930
                  TabIndex        =   163
                  Top             =   585
                  Width           =   1185
               End
               Begin VB.OptionButton maintenance 
                  Caption         =   "Gas Station"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   2085
                  TabIndex        =   164
                  Top             =   585
                  Width           =   1215
               End
               Begin VB.OptionButton maintenance 
                  Caption         =   "Others"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   3
                  Left            =   3345
                  TabIndex        =   165
                  Top             =   585
                  Width           =   855
               End
               Begin VB.TextBox TxtmaintenaceKMS 
                  Height          =   255
                  Left            =   2700
                  TabIndex        =   159
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox TxtmaintenaceEvry 
                  Height          =   255
                  Left            =   630
                  TabIndex        =   160
                  Top             =   255
                  Width           =   735
               End
               Begin VB.Label Label50 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Month/s,Every"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   1440
                  TabIndex        =   161
                  Top             =   255
                  Width           =   1245
               End
               Begin VB.Label Label49 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Every"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   30
                  TabIndex        =   158
                  Top             =   255
                  Width           =   525
               End
            End
            Begin VB.Frame Frame9 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   1320
               TabIndex        =   94
               Top             =   210
               Width           =   2985
               Begin VB.OptionButton Engine 
                  Caption         =   "All Temp"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   1950
                  TabIndex        =   97
                  Top             =   90
                  Width           =   975
               End
               Begin VB.OptionButton Engine 
                  Caption         =   "Cold"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   96
                  Top             =   75
                  Width           =   615
               End
               Begin VB.OptionButton Engine 
                  Caption         =   "Hot"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   45
                  TabIndex        =   95
                  Top             =   75
                  Width           =   615
               End
            End
            Begin VB.TextBox txtOthercondition 
               Height          =   285
               Left            =   2040
               TabIndex        =   167
               Top             =   6270
               Width           =   2310
            End
            Begin VB.Frame Frame10 
               BorderStyle     =   0  'None
               Height          =   405
               Left            =   1365
               TabIndex        =   100
               Top             =   420
               Width           =   3045
               Begin VB.OptionButton Weather 
                  Caption         =   "Cold "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   101
                  Top             =   120
                  Width           =   645
               End
               Begin VB.OptionButton Weather 
                  Caption         =   "Warm"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1005
                  TabIndex        =   102
                  Top             =   120
                  Width           =   885
               End
               Begin VB.OptionButton Weather 
                  Caption         =   "All temp"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   1920
                  TabIndex        =   103
                  Top             =   150
                  Width           =   1005
               End
            End
            Begin VB.Frame Frame11 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   1290
               TabIndex        =   104
               Top             =   690
               Width           =   3045
               Begin VB.OptionButton Shifting 
                  Caption         =   "Fast"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   106
                  Top             =   120
                  Width           =   855
               End
               Begin VB.OptionButton Shifting 
                  Caption         =   "Normal"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   90
                  TabIndex        =   105
                  Top             =   120
                  Width           =   855
               End
            End
            Begin VB.Frame Frame12 
               BorderStyle     =   0  'None
               Height          =   810
               Left            =   1260
               TabIndex        =   113
               Top             =   1320
               Width           =   2895
               Begin VB.OptionButton MT 
                  Caption         =   "Reverse"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   120
                  Top             =   510
                  Width           =   915
               End
               Begin VB.OptionButton MT 
                  Caption         =   "Neutral"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   2040
                  TabIndex        =   119
                  Top             =   300
                  Width           =   915
               End
               Begin VB.OptionButton MT 
                  Caption         =   "5th"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   1110
                  TabIndex        =   118
                  Top             =   300
                  Width           =   705
               End
               Begin VB.OptionButton MT 
                  Caption         =   "4th"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   116
                  Top             =   270
                  Width           =   705
               End
               Begin VB.OptionButton MT 
                  Caption         =   "3rd"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   2040
                  TabIndex        =   117
                  Top             =   60
                  Width           =   705
               End
               Begin VB.OptionButton MT 
                  Caption         =   "2nd"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1110
                  TabIndex        =   115
                  Top             =   60
                  Width           =   705
               End
               Begin VB.OptionButton MT 
                  Caption         =   "1st"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   114
                  Top             =   30
                  Width           =   705
               End
            End
            Begin VB.Frame Frame13 
               BorderStyle     =   0  'None
               Height          =   330
               Left            =   1320
               TabIndex        =   109
               Top             =   1035
               Width           =   2085
               Begin VB.OptionButton ShiftPosition 
                  Caption         =   "2wd"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   111
                  Top             =   45
                  Width           =   705
               End
               Begin VB.OptionButton ShiftPosition 
                  Caption         =   "4wd"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   60
                  TabIndex        =   110
                  Top             =   60
                  Width           =   705
               End
            End
            Begin VB.Frame Frame14 
               BorderStyle     =   0  'None
               Height          =   510
               Left            =   1260
               TabIndex        =   136
               Top             =   3330
               Width           =   3075
               Begin VB.OptionButton Location 
                  Caption         =   "Stop & go traffic"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   3
                  Left            =   120
                  TabIndex        =   140
                  Top             =   210
                  Width           =   1425
               End
               Begin VB.OptionButton Location 
                  Caption         =   "Downhill"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   2070
                  TabIndex        =   139
                  Top             =   -45
                  Width           =   945
               End
               Begin VB.OptionButton Location 
                  Caption         =   "uphill"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1185
                  TabIndex        =   138
                  Top             =   -30
                  Width           =   735
               End
               Begin VB.OptionButton Location 
                  Caption         =   "Highway"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   137
                  Top             =   -15
                  Width           =   975
               End
            End
            Begin VB.Frame Frame15 
               BorderStyle     =   0  'None
               Height          =   645
               Left            =   1290
               TabIndex        =   122
               Top             =   2085
               Width           =   3105
               Begin VB.OptionButton AT 
                  Caption         =   "R"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   90
                  TabIndex        =   129
                  Top             =   450
                  Width           =   1035
               End
               Begin VB.OptionButton AT 
                  Caption         =   "OverDrive"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   2010
                  TabIndex        =   128
                  Top             =   255
                  Width           =   1035
               End
               Begin VB.OptionButton AT 
                  Caption         =   "N"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   1110
                  TabIndex        =   127
                  Top             =   270
                  Width           =   705
               End
               Begin VB.OptionButton AT 
                  Caption         =   "D"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   90
                  TabIndex        =   126
                  Top             =   240
                  Width           =   705
               End
               Begin VB.OptionButton AT 
                  Caption         =   "3"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   2010
                  TabIndex        =   125
                  Top             =   30
                  Width           =   705
               End
               Begin VB.OptionButton AT 
                  Caption         =   "2"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1110
                  TabIndex        =   124
                  Top             =   30
                  Width           =   705
               End
               Begin VB.OptionButton AT 
                  Caption         =   "L"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   90
                  TabIndex        =   123
                  Top             =   30
                  Width           =   705
               End
            End
            Begin VB.Frame Frame16 
               BorderStyle     =   0  'None
               Height          =   600
               Left            =   1245
               TabIndex        =   130
               Top             =   2745
               Width           =   2895
               Begin VB.OptionButton Road 
                  Caption         =   "Muddy"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   135
                  TabIndex        =   134
                  Top             =   285
                  Width           =   855
               End
               Begin VB.OptionButton Road 
                  Caption         =   "Rocky"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   2100
                  TabIndex        =   133
                  Top             =   60
                  Width           =   855
               End
               Begin VB.OptionButton Road 
                  Caption         =   "unpaved"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1110
                  TabIndex        =   132
                  Top             =   60
                  Width           =   945
               End
               Begin VB.OptionButton Road 
                  Caption         =   "paved"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   135
                  TabIndex        =   131
                  Top             =   60
                  Width           =   825
               End
            End
            Begin VB.Frame Frame17 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   1350
               TabIndex        =   154
               Top             =   4695
               Width           =   2295
               Begin VB.OptionButton Accesories 
                  Caption         =   "A/c on"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   30
                  TabIndex        =   155
                  Top             =   -15
                  Width           =   885
               End
               Begin VB.OptionButton Accesories 
                  Caption         =   "Electric load"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1125
                  TabIndex        =   156
                  Top             =   -30
                  Width           =   1215
               End
            End
            Begin VB.Frame Frame18 
               BorderStyle     =   0  'None
               Height          =   600
               Left            =   1305
               TabIndex        =   143
               Top             =   3795
               Width           =   3015
               Begin VB.OptionButton Action 
                  Caption         =   "Decelaration"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   1155
                  TabIndex        =   148
                  Top             =   285
                  Width           =   1275
               End
               Begin VB.OptionButton Action 
                  Caption         =   "Accelation"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   2025
                  TabIndex        =   146
                  Top             =   15
                  Width           =   1245
               End
               Begin VB.OptionButton Action 
                  Caption         =   "Cruising"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   75
                  TabIndex        =   147
                  Top             =   285
                  Width           =   885
               End
               Begin VB.OptionButton Action 
                  Caption         =   "idline"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1155
                  TabIndex        =   145
                  Top             =   45
                  Width           =   705
               End
               Begin VB.OptionButton Action 
                  Caption         =   "Cranking"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   75
                  TabIndex        =   144
                  Top             =   45
                  Width           =   945
               End
            End
            Begin VB.Frame Frame19 
               BorderStyle     =   0  'None
               Height          =   330
               Left            =   1365
               TabIndex        =   150
               Top             =   4335
               Width           =   2445
               Begin VB.OptionButton Occurence 
                  Caption         =   "Consistent"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   15
                  TabIndex        =   151
                  Top             =   60
                  Width           =   1065
               End
               Begin VB.OptionButton Occurence 
                  Caption         =   "intermittent"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1110
                  TabIndex        =   152
                  Top             =   45
                  Width           =   1395
               End
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Other(Pls Specify)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   90
               TabIndex        =   166
               Top             =   6300
               Width           =   1875
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Engine Temp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   150
               TabIndex        =   98
               Top             =   270
               Width           =   1155
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Weather"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   270
               TabIndex        =   99
               Top             =   540
               Width           =   1035
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Shifting"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   90
               TabIndex        =   107
               Top             =   825
               Width           =   1215
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -510
               TabIndex        =   108
               Top             =   1110
               Width           =   1815
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "M/T"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   825
               TabIndex        =   112
               Top             =   1425
               Width           =   435
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "A/T"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   825
               TabIndex        =   121
               Top             =   2130
               Width           =   435
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Road"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   705
               TabIndex        =   135
               Top             =   2820
               Width           =   555
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Location "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   360
               TabIndex        =   141
               Top             =   3360
               Width           =   975
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Action"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -480
               TabIndex        =   142
               Top             =   3840
               Width           =   1785
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Occurrence"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   90
               TabIndex        =   149
               Top             =   4365
               Width           =   1245
            End
            Begin VB.Label lblAccesories_ACON 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Accesories"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   225
               TabIndex        =   153
               Top             =   4680
               Width           =   1095
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Other Commenst"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   30
            TabIndex        =   168
            Top             =   6180
            Width           =   5925
            Begin VB.TextBox txtOtherComments 
               Height          =   435
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   169
               Top             =   240
               Width           =   5775
            End
         End
         Begin VB.Frame FrameC 
            Height          =   2475
            Left            =   30
            TabIndex        =   58
            Top             =   -30
            Width           =   5865
            Begin VB.TextBox txtsub 
               Height          =   285
               Left            =   3720
               TabIndex        =   223
               Top             =   120
               Width           =   1695
            End
            Begin VB.ComboBox cbojob 
               Height          =   315
               Left            =   1650
               TabIndex        =   200
               Text            =   "Combo1"
               Top             =   120
               Width           =   2085
            End
            Begin VB.TextBox txtRecommendation 
               Height          =   285
               Left            =   1635
               TabIndex        =   69
               Top             =   2100
               Width           =   3795
            End
            Begin VB.TextBox txtCorrectiveAct 
               Height          =   285
               Left            =   1635
               TabIndex        =   67
               Top             =   1800
               Width           =   3795
            End
            Begin VB.TextBox txthistory 
               Height          =   285
               Left            =   1650
               TabIndex        =   61
               Top             =   420
               Width           =   3765
            End
            Begin VB.TextBox txtDescription 
               Height          =   285
               Left            =   1650
               TabIndex        =   63
               Top             =   750
               Width           =   3765
            End
            Begin VB.TextBox txtAnalysis 
               Height          =   765
               Left            =   1635
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   65
               Top             =   1035
               Width           =   3795
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Recommendation"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   90
               TabIndex        =   68
               Top             =   2115
               Width           =   1515
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Corrective Action"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   90
               TabIndex        =   66
               Top             =   1815
               Width           =   1515
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Subject"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   810
               TabIndex        =   59
               Top             =   180
               Width           =   795
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "History"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   810
               TabIndex        =   60
               Top             =   450
               Width           =   795
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Descripion"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   735
               TabIndex        =   62
               Top             =   750
               Width           =   885
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Analysis (Cause of the problem)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   75
               TabIndex        =   64
               Top             =   1050
               Width           =   1545
            End
         End
         Begin VB.TextBox InvoiceNo 
            Alignment       =   2  'Center
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3900
            TabIndex        =   84
            Top             =   5250
            Width           =   1965
         End
         Begin VB.TextBox txtGrandtotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   92
            Top             =   5850
            Width           =   1485
         End
         Begin VB.TextBox txtTotalSubletRepair 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   90
            Top             =   5820
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox txtTotalLaborCost 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   5265
            Width           =   1485
         End
         Begin VB.TextBox txtCausalPartNo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            TabIndex        =   76
            Top             =   2670
            Width           =   1395
         End
         Begin VB.TextBox txtNatureCode 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2655
            TabIndex        =   73
            Top             =   2490
            Width           =   615
         End
         Begin VB.TextBox txtPaintCode 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4410
            TabIndex        =   75
            Top             =   2490
            Width           =   1440
         End
         Begin VB.TextBox txtCauseCode 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2655
            TabIndex        =   78
            Top             =   2790
            Width           =   615
         End
         Begin VB.TextBox txtSubletCode 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4410
            TabIndex        =   80
            Top             =   2820
            Width           =   1455
         End
         Begin VB.PictureBox opCodePic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1725
            Left            =   1140
            ScaleHeight     =   1695
            ScaleWidth      =   3195
            TabIndex        =   199
            Top             =   3690
            Width           =   3225
            Begin VB.CommandButton Edt 
               Caption         =   ".."
               Height          =   345
               Left            =   2280
               TabIndex        =   212
               Top             =   990
               Width           =   345
            End
            Begin VB.TextBox txtCost 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1080
               TabIndex        =   210
               Text            =   "247.50"
               Top             =   990
               Width           =   1155
            End
            Begin VB.TextBox txtLTS 
               Height          =   345
               Left            =   1080
               TabIndex        =   202
               Top             =   300
               Width           =   645
            End
            Begin VB.TextBox txtOpCode 
               Height          =   315
               Left            =   1080
               TabIndex        =   201
               Top             =   660
               Width           =   1935
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption5 
               Height          =   225
               Left            =   -1260
               TabIndex        =   236
               Top             =   0
               Width           =   5685
               _Version        =   655364
               _ExtentX        =   10028
               _ExtentY        =   397
               _StockProps     =   14
               Caption         =   "::LTS Information::"
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   1
               ForeColor       =   0
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "F5 -Save"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2310
               TabIndex        =   211
               Top             =   1410
               Width           =   735
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Labor Cost"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   -210
               TabIndex        =   209
               Top             =   1020
               Width           =   1245
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ecs-To Cancel"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   1110
               TabIndex        =   207
               Top             =   1410
               Width           =   1185
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "OpCode"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   210
               TabIndex        =   206
               Top             =   690
               Width           =   825
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "LTS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   540
               TabIndex        =   203
               Top             =   360
               Width           =   465
            End
         End
         Begin VB.PictureBox picsublet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2745
            Left            =   150
            ScaleHeight     =   2715
            ScaleWidth      =   5625
            TabIndex        =   214
            Top             =   3510
            Width           =   5655
            Begin VB.TextBox txtsubletopcode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1590
               TabIndex        =   3
               Top             =   1230
               Width           =   1365
            End
            Begin VB.TextBox txtsubletLTS 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1590
               TabIndex        =   5
               Top             =   1890
               Width           =   1365
            End
            Begin VB.TextBox txtqty 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1590
               TabIndex        =   4
               Top             =   1560
               Width           =   1365
            End
            Begin VB.TextBox txtpartno 
               Height          =   315
               Left            =   1590
               TabIndex        =   1
               Top             =   390
               Width           =   2055
            End
            Begin VB.CommandButton cmdarnie 
               Caption         =   "Cancel"
               Height          =   435
               Left            =   4470
               TabIndex        =   218
               Top             =   2220
               Width           =   1095
            End
            Begin VB.CommandButton cmdaddsub 
               Caption         =   "Add"
               Height          =   435
               Left            =   3390
               TabIndex        =   217
               Top             =   2220
               Width           =   1095
            End
            Begin VB.TextBox txtsubletprc 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1590
               TabIndex        =   6
               Top             =   2220
               Width           =   1365
            End
            Begin VB.TextBox txtsubletDesc 
               Height          =   495
               Left            =   1590
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   2
               Top             =   720
               Width           =   3975
            End
            Begin VB.Label Label69 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "OpCode"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   510
               TabIndex        =   240
               Top             =   1230
               Width           =   1035
            End
            Begin VB.Label Label68 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "LTS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   720
               TabIndex        =   238
               Top             =   1890
               Width           =   795
            End
            Begin VB.Label Label67 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Part Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   300
               TabIndex        =   237
               Top             =   870
               Width           =   1245
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
               Height          =   285
               Left            =   0
               TabIndex        =   235
               Top             =   -30
               Width           =   5685
               _Version        =   655364
               _ExtentX        =   10028
               _ExtentY        =   503
               _StockProps     =   14
               Caption         =   "::Sublet Information::"
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               VisualTheme     =   0
               Alignment       =   1
               ForeColor       =   0
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "QTY"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   510
               TabIndex        =   234
               Top             =   1560
               Width           =   1035
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Part Number"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   300
               TabIndex        =   233
               Top             =   390
               Width           =   1245
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Sublet Price"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   150
               TabIndex        =   216
               Top             =   2190
               Width           =   1365
            End
            Begin VB.Label Label62 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Sublet decription/"
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
               Left            =   -420
               TabIndex        =   215
               Top             =   690
               Width           =   1965
            End
         End
         Begin VB.Line Line2 
            X1              =   5340
            X2              =   4380
            Y1              =   5010
            Y2              =   5010
         End
         Begin VB.Line Line1 
            X1              =   5340
            X2              =   4380
            Y1              =   4980
            Y2              =   4980
         End
         Begin VB.Label Label61 
            Caption         =   "F3 - Add Sublet "
            Height          =   255
            Left            =   3840
            TabIndex        =   213
            Top             =   5610
            Width           =   1185
         End
         Begin VB.Label Label59 
            Caption         =   "NOTE(Labor Cost) 247.50 x Total LTS"
            Height          =   255
            Left            =   30
            TabIndex        =   208
            Top             =   4740
            Width           =   2805
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Total LTS"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1035
            TabIndex        =   85
            Top             =   5010
            Width           =   795
         End
         Begin VB.Label Label43 
            Caption         =   "Total Cost"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3420
            TabIndex        =   82
            Top             =   4770
            Width           =   930
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Invioce/OR #"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4290
            TabIndex        =   81
            Top             =   5040
            Width           =   1215
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   720
            TabIndex        =   91
            Top             =   5880
            Width           =   1065
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Sublet repair"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   180
            TabIndex        =   89
            Top             =   5565
            Width           =   1635
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Labor Charge"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   345
            TabIndex        =   87
            Top             =   5280
            Width           =   1455
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Causal Part No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   71
            Top             =   2460
            Width           =   1305
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nature Code"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1500
            TabIndex        =   72
            Top             =   2490
            Width           =   1125
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cause Code"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1560
            TabIndex        =   77
            Top             =   2790
            Width           =   1065
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Paint Code "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3435
            TabIndex        =   74
            Top             =   2505
            Width           =   915
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Code"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   750
            TabIndex        =   70
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet Code"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3210
            TabIndex        =   79
            Top             =   2805
            Width           =   1185
         End
      End
      Begin VB.VScrollBar ScrollBar1 
         Height          =   4575
         LargeChange     =   500
         Left            =   10620
         Max             =   10000
         MouseIcon       =   "frmCSMSQuailityInformation.frx":8933
         MousePointer    =   99  'Custom
         SmallChange     =   250
         TabIndex        =   177
         Top             =   0
         Value           =   50
         Width           =   315
      End
   End
   Begin VB.PictureBox thePic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   870
      ScaleHeight     =   5415
      ScaleWidth      =   9495
      TabIndex        =   48
      Top             =   1500
      Width           =   9525
      Begin VB.OptionButton otpSearch 
         BackColor       =   &H00C0FFFF&
         Caption         =   "RO Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   198
         Top             =   390
         Value           =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4410
         Left            =   90
         TabIndex        =   52
         Top             =   690
         Width           =   9360
         Begin MSComctlLib.ListView listcust 
            Height          =   4110
            Left            =   90
            TabIndex        =   53
            Top             =   210
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   7250
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
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCSMSQuailityInformation.frx":8C3D
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No "
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Repair Order"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "First Name"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Last Name"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Address"
               Object.Width           =   6350
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Customer"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Plate No"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "VIn"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Engine"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "service advi"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.TextBox txtkeyword 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3930
         TabIndex        =   51
         Top             =   390
         Width           =   2865
      End
      Begin VB.OptionButton otpSearch 
         BackColor       =   &H00C0FFFF&
         Caption         =   "LastName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2730
         TabIndex        =   50
         Top             =   390
         Width           =   1200
      End
      Begin VB.OptionButton otpSearch 
         BackColor       =   &H00C0FFFF&
         Caption         =   "firstName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1500
         TabIndex        =   49
         Top             =   390
         Width           =   1200
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   ":::Warranty List:::"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   3780
         TabIndex        =   225
         Top             =   60
         Width           =   1545
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   -30
         TabIndex        =   224
         Top             =   0
         Width           =   9495
         _Version        =   655364
         _ExtentX        =   16748
         _ExtentY        =   556
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Double Click The Item To Select"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   90
         TabIndex        =   55
         Top             =   5100
         Width           =   2775
      End
      Begin VB.Label lblclose 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8850
         MouseIcon       =   "frmCSMSQuailityInformation.frx":8D9F
         MousePointer    =   99  'Custom
         TabIndex        =   54
         ToolTipText     =   "Close"
         Top             =   5100
         Width           =   570
      End
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "LTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   205
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label57 
      BackStyle       =   0  'Transparent
      Caption         =   "LTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   204
      Top             =   0
      Width           =   465
   End
   Begin VB.Menu mnuremove 
      Caption         =   "Remove Temporary"
      Visible         =   0   'False
      Begin VB.Menu mnuref 
         Caption         =   "Refresh Repor"
      End
      Begin VB.Menu mnuremoveparts 
         Caption         =   "Make a report"
      End
   End
End
Attribute VB_Name = "frmCSMSQualityInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS                                                 As New ADODB.Recordset
Dim TheWcode_Mode                                      As Boolean

Dim UPDATE_MODE                                        As Boolean
Dim XtheRo, XCustomer, xVIN, xENGINE                   As String
Attribute XCustomer.VB_VarUserMemId = 1073938435
Attribute xVIN.VB_VarUserMemId = 1073938435
Attribute xENGINE.VB_VarUserMemId = 1073938435
Dim thecosttotal, theQTYTotal                          As Double
Attribute thecosttotal.VB_VarUserMemId = 1073938438
Attribute theQTYTotal.VB_VarUserMemId = 1073938438
Dim theID                                              As String
Attribute theID.VB_VarUserMemId = 1073938440
Dim theReference                                       As String
Attribute theReference.VB_VarUserMemId = 1073938441
Dim theReferenceNo, thePartNo, thepartName, theQty, thePCost, theopcode, theLts, theTotalCost, TheTotalLTS As String
Attribute theReferenceNo.VB_VarUserMemId = 1073938442
Attribute thePartNo.VB_VarUserMemId = 1073938442
Attribute thepartName.VB_VarUserMemId = 1073938442
Attribute theQty.VB_VarUserMemId = 1073938442
Attribute thePCost.VB_VarUserMemId = 1073938442
Attribute theopcode.VB_VarUserMemId = 1073938442
Attribute theLts.VB_VarUserMemId = 1073938442
Attribute theTotalCost.VB_VarUserMemId = 1073938442
Attribute TheTotalLTS.VB_VarUserMemId = 1073938442
Dim IsParts                                            As Boolean
Attribute IsParts.VB_VarUserMemId = 1073938444
Dim DefaultLabor                                       As Double
Attribute DefaultLabor.VB_VarUserMemId = 1073938445
Dim thesubletTotal                                     As Double
Attribute thesubletTotal.VB_VarUserMemId = 1073938446
Dim thesublet                                          As Double
Attribute thesublet.VB_VarUserMemId = 1073938447
Dim pindot                                             As Boolean
Attribute pindot.VB_VarUserMemId = 1073938448

Sub FillCBOTechnician()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT TECH_NAME FROM CSMS_vw_TechnicianAvailability"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    cboTECH.Clear
    With RS
        Do While Not .EOF
            cboTECH.AddItem !TECH_NAME
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub Saveinformation()
    Dim thePWADATE                                     As Date
    Dim theRo, TheDLR, ThePWARequest, theCustomer, theVinNo, EngineNo, theTMAxle, TheTransmisionType As String
    Dim theAttachment, theSubject, thehistory, theDescription, theAnalysis, theCorrective, TheRecommendation As String
    Dim theCasualPartno, TheNatureCode, thecausecode, ThePaintCode, theSubletCode As String
    Dim theTotalLaborCost, theTotalSubletRepair, TheGrandTotal As String
    Dim TheInvoice, ThedealerCode                      As String
    Dim theInspectionDate, theRepairOrderDate          As Date
    Dim theMileage, thePlateNo, theSentReistrationCard, theSaleAdvisor, thetechnician As String
    Dim TheEngineTemp, TheWeather, theShifting, theshiftPosition, theMT, theAT, theRoad, thelocation, theAction As String
    Dim theOccurence, theAccessories, theVechicleMaintenance, theEveryMoth, theEveryKMS, theOtherCondition, theOtherComments As String
    Dim TheRequestedby, theCheckedBy, theApproveby     As String
    Dim thecost, theLts                                As String

    On Error GoTo ErrorCode



    thecost = thecosttotal

    theRo = Trim(txtRO.Text)
    TheDLR = Trim(txtReferenceNo.Text)
    If otpRequestY.Value = True Then
        ThePWARequest = "Y"
    End If
    If otpRequestN.Value = True Then
        ThePWARequest = "N"
    End If

    thePWADATE = (DeviverDate)
    theCustomer = Trim(txtCustomer.Text)
    theVinNo = Trim(txtVIN.Text)
    EngineNo = Trim(txtEngineNo.Text)
    theTMAxle = Trim(txtTMAxleNo.Text)

    If OtpTransTypeManual.Value = True Then
        TheTransmisionType = "M"
    End If
    If OtpTransTypeAuto.Value = True Then
        TheTransmisionType = "A"
    End If
    theAttachment = "Test"
    theSubject = txtsub.Text
    thehistory = Trim(txthistory.Text)
    theDescription = Trim(txtDescription.Text)
    theAnalysis = Trim(txtAnalysis.Text)
    theCorrective = Trim(txtCorrectiveAct.Text)
    TheRecommendation = Trim(txtRecommendation.Text)
    theCasualPartno = Trim(txtCausalPartNo.Text)
    TheNatureCode = Trim(txtNatureCode.Text)
    ThePaintCode = Trim(txtPaintCode.Text)
    theSubletCode = Trim(txtSubletCode.Text)
    theTotalLaborCost = txtTotalLaborCost.Text
    thesublet = txtpassSublet.Text
    TheGrandTotal = txtGrandtotal.Text
    TheInvoice = Trim(InvoiceNo.Text)
    theOtherComments = Trim(txtOtherComments.Text)
    ThedealerCode = Trim(txtDealer.Text)
    theInspectionDate = InspectionDate
    theRepairOrderDate = RepairDate
    thecausecode = Trim(txtCauseCode.Text)


    theMileage = Trim(txtMileage.Text)
    thePlateNo = Trim(txtPLATENO.Text)
    theSaleAdvisor = Trim(txtsaleAdvisor.Text)
    thetechnician = cboTECH

    If OtpSentCardY.Value = True Then
        theSentReistrationCard = "Y"
    ElseIf OtpSentCardN.Value = True Then
        theSentReistrationCard = "N"
    End If
    'Engine
    If Engine(0).Value = True Then
        TheEngineTemp = "Hot"
    ElseIf Engine(1).Value = True Then TheEngineTemp = "Cold"
    ElseIf Engine(2).Value = True Then TheEngineTemp = "AllTemp"
    End If
    'Weather
    If Weather(0).Value = True Then
        TheWeather = "Cold "
    ElseIf Weather(1).Value = True Then TheWeather = "Warm"
    ElseIf Weather(2).Value = True Then TheWeather = "AllTemp"
    End If
    'shifting
    If Shifting(0).Value = True Then
        theShifting = "Normal"
    ElseIf Shifting(1).Value = True Then theShifting = "Fast"
    End If
    'Shift Position
    If ShiftPosition(0).Value = True Then
        theshiftPosition = "4wd"
    ElseIf ShiftPosition(0).Value = True Then
        theshiftPosition = "4wd"
    Else
        theshiftPosition = "2wd"
    End If
    'MT
    If MT(0).Value = True Then
        theMT = "1st"
    ElseIf MT(1).Value = True Then theMT = "2nd"
    ElseIf MT(2).Value = True Then theMT = "3rd"
    ElseIf MT(3).Value = True Then theMT = "4th"
    ElseIf MT(4).Value = True Then theMT = "5th"
    ElseIf MT(5).Value = True Then theMT = "Neutral"
    Else
        theMT = "Reverse"
    End If
    If AT(0).Value = True Then
        theAT = "L"
    ElseIf AT(1).Value = True Then theAT = "2"
    ElseIf AT(2).Value = True Then theAT = "3"
    ElseIf AT(3).Value = True Then theAT = "D"
    ElseIf AT(4).Value = True Then theAT = "N"
    ElseIf AT(5).Value = True Then theAT = "R"
    Else
        theAT = "OverDrive"
    End If
    If Road(0).Value = True Then
        theRoad = "Paved"
    ElseIf Road(1).Value = True Then theRoad = "Unpaved"
    ElseIf Road(2).Value = True Then theRoad = "Rocky"
    ElseIf Road(3).Value = True Then theRoad = "Muddy"
    End If
    If Location(0).Value = True Then
        thelocation = "HighWay"
    ElseIf Location(1).Value = True Then thelocation = "Uphill"
    ElseIf Location(2).Value = True Then thelocation = "Downhill"
    Else
        thelocation = "StopGoTraffic"
    End If
    If Action(0).Value = True Then
        theAction = "Cranking "
    ElseIf Action(1).Value = True Then theAction = "idling"
    ElseIf Action(2).Value = True Then theAction = "Cruising"
    ElseIf Action(3).Value = True Then theAction = "Accelarating"
    ElseIf Action(4).Value = True Then theAction = "Deccelaration"
    End If
    If Occurence(0).Value = True Then
        theOccurence = "Consistent"
    Else
        theOccurence = "intermittent"
    End If
    If maintenance(0).Value = True Then
        theVechicleMaintenance = "dealer"
    ElseIf maintenance(1).Value = True Then theVechicleMaintenance = "3starshop"
    ElseIf maintenance(2).Value = True Then theVechicleMaintenance = "gasStation"
    Else
        theVechicleMaintenance = "Other"
    End If
    If Accesories(0).Value = True Then
        theAccessories = "a/cON"
    Else
        theAccessories = "Heavy Electrical load"
    End If

    theOtherCondition = Trim(txtOthercondition)
    theEveryMoth = Trim(TxtmaintenaceEvry.Text)
    theEveryKMS = Trim(TxtmaintenaceKMS.Text)
    TheRequestedby = Trim(txtRequested.Text)
    theCheckedBy = Trim(txtgeneralManager.Text)
    theApproveby = Trim(txtServiceDept.Text)

    If Len(TheDLR) = 0 Then
        MsgBox "Missing parameters...DLR CQIR Reference #", vbExclamation, "Warning"
        On Error Resume Next
        txtReferenceNo.SetFocus
        Exit Sub
    End If

    If UPDATE_MODE = False Then
        IsParts = True
        gconDMIS.Execute "INSERT INTO CSMS_CQIR " & _
                         "(Ro_no,DLR_CQIR_referenceno,PWA_Request,PWA_date,customer,VinNo,EngineNo,TM_axleno,transmissionType,attachments,subject,history,description,analysis,CorrectiveAction,Recommendation,causalPartno,natureCode,causeCode,paintCode,totalcost,totallts,subletCode,totallaborcost,totalsubletRepair,grandtotal,invoiceno_ORno,othercomments,dealerCode,inspectiondate,repairdate,mileage,plateno,sentregistrationcard,serviceAdvisor,technician,con_enginetemp,con_weather,con_shifting,con_shifposition,con_mt,Con_at,con_road,con_location,con_Action,con_occurence,con_accessories,VechicleMaintenance,everymonth,everyKMS,OtherCondition,requestedby,checkedby,approvedby)" & _
                       " VALUES('" & theRo & "','" & TheDLR & "','" & ThePWARequest & "','" & thePWADATE & "','" & theCustomer & "','" & theVinNo & "','" & EngineNo & "','" & theTMAxle & "','" & TheTransmisionType & "','" & theAttachment & "','" & theSubject & "','" & thehistory & "','" & theDescription & "','" & theAnalysis & "','" & theCorrective & "','" & TheRecommendation & "','" & theCasualPartno & "','" & TheNatureCode & "','" & thecausecode & "','" & ThePaintCode & "','" & thecost & "','" & theLts & "','" & theSubletCode & "','" & theTotalLaborCost & _
                         "','" & thesublet & "','" & TheGrandTotal & "','" & TheInvoice & "','" & theOtherComments & "','" & ThedealerCode & "','" & theInspectionDate & "','" & theRepairOrderDate & "','" & theMileage & "','" & thePlateNo & "','" & theSentReistrationCard & "','" & theSaleAdvisor & "','" & thetechnician & "','" & TheEngineTemp & "','" & TheWeather & "','" & theShifting & "','" & theshiftPosition & "','" & theMT & "','" & theAT & "','" & theRoad & "','" & thelocation & "','" & theAction & _
                         "','" & theOccurence & "','" & theAccessories & "','" & theVechicleMaintenance & "','" & theEveryMoth & "','" & theEveryKMS & "','" & theOtherCondition & "','" & TheRequestedby & "','" & theCheckedBy & "','" & theApproveby & "')"

        displayPartsinfo
        gconDMIS.Execute "update CSMS_CQIRPARTS set DLR_CQIR_ReferenceNo='" & txtReferenceNo & "' where Ro='" & txtRO & "'"
        IsParts = False
        MsgBox "All information Has been Save!", vbInformation, "information"
        LogAudit "A", "CQIR REPORT", "PWA/RO" & TxtPWANo & "/" & txtRO
    Else


        gconDMIS.Execute "update CSMS_CQIR set Ro_no='" & theRo & "',DLR_CQIR_referenceno='" & TheDLR & "',PWA_Request='" & ThePWARequest & "',PWA_date='" & thePWADATE & _
                         "',customer='" & theCustomer & "',VinNo='" & theVinNo & "',EngineNo='" & EngineNo & "',TM_axleno='" & theTMAxle & "',transmissionType='" & TheTransmisionType & _
                         "',attachments='" & theAttachment & "',subject='" & theSubject & "',history='" & thehistory & "',description='" & theDescription & "',analysis='" & theAnalysis & _
                         "',CorrectiveAction='" & theCorrective & "',Recommendation='" & TheRecommendation & "',causalPartno='" & theCasualPartno & "',natureCode='" & TheNatureCode & _
                         "' ,causeCode='" & thecausecode & "',paintCode='" & ThePaintCode & "',subletCode='" & theSubletCode & "',totallaborcost='" & theTotalLaborCost & _
                         "', totalsubletrepair ='" & thesublet & "',grandtotal='" & TheGrandTotal & "',invoiceno_ORno='" & TheInvoice & "',othercomments='" & theOtherComments & _
                         "',dealerCode='" & ThedealerCode & "',inspectiondate='" & theInspectionDate & "',repairdate='" & theRepairOrderDate & "',mileage='" & theMileage & "',plateno='" & thePlateNo & _
                         "',sentregistrationcard='" & theSentReistrationCard & "',serviceAdvisor='" & theSaleAdvisor & "',technician='" & thetechnician & _
                         "',con_enginetemp='" & TheEngineTemp & "',con_weather='" & TheWeather & "',con_shifting='" & theShifting & "',con_shifposition='" & theshiftPosition & "',con_mt='" & theMT & _
                         "',Con_at='" & theAT & "',con_road='" & theRoad & "',con_location='" & thelocation & "',con_Action='" & theAction & "',con_occurence='" & theOccurence & "',con_accessories='" & theAccessories & _
                         "',VechicleMaintenance='" & theVechicleMaintenance & "',everymonth='" & theEveryMoth & "',everyKMS='" & theEveryKMS & "',OtherCondition='" & theOtherCondition & "',requestedby='" & TheRequestedby & _
                         "',checkedby='" & theCheckedBy & "',approvedby='" & theApproveby & "' WHERE ID='" & lblID.Caption & "'"
        LogAudit "E", "CQIR REPORT", "PWA/RO" & TxtPWANo & "/" & txtRO
        MsgBox "All Information has Been Updated", vbInformation, "Information"
        UPDATE_MODE = False
    End If
    rsRefresh
    Call initMemvars
    StoreMemVars
    thepicSave.Visible = False
    thepicControl.Visible = True

    Exit Sub

ErrorCode:

    ShowVBError
    Exit Sub
End Sub

Sub StoreMemVars()

    If Not RS.EOF And Not RS.BOF Then
        DeviverDate = Null2String(RS!pwa_date)
        txtRO.Text = Null2String(RS!RO_NO)
        txtReferenceNo.Text = Null2String(RS!DLR_CQIR_REFERENCENO)
        txtTotalLaborCost.Text = Null2String(RS!TotalLaborCost)
        If RS!pwa_request = "Y" Then
            otpRequestY.Value = True
        Else
            otpRequestN.Value = False
        End If
        txtCustomer.Text = Null2String(RS!Customer)
        txtVIN.Text = Null2String(RS!VINNO)
        txtEngineNo.Text = Null2String(RS!EngineNo)
        txtTMAxleNo.Text = Null2String(RS!tm_axleno)
        If RS!transmissiontype = "M" Then
            OtpTransTypeManual.Value = True
        Else
            OtpTransTypeAuto.Value = True
        End If
        If RS!attachments = "Test" Then
            checkPhoto.Value = 1
        Else

        End If
        cbojob.Text = Null2String(RS!Subject)
        txtsub.Text = Null2String(RS!Subject)

        txthistory.Text = Null2String(RS!history)
        txtDescription.Text = Null2String(RS!Description)
        txtAnalysis.Text = Null2String(RS!ANALYSIS)
        txtCorrectiveAct.Text = Null2String(RS!correctiveAction)
        txtRecommendation.Text = Null2String(RS!RECOMMENDATION)
        txtCausalPartNo.Text = Null2String(RS!CAUSALPARTNO)
        txtNatureCode.Text = Null2String(RS!NATURECODE)
        txtCauseCode.Text = Null2String(RS!CAUSECODE)
        txtPaintCode.Text = Null2String(RS!paintcode)
        txtSubletCode.Text = Null2String(RS!subletcode)
        txtpassSublet.Text = Null2String(RS!TotalSUBLETREPAIR)
        txtGrandtotal.Text = Null2String(RS!grandtotal)
        InvoiceNo.Text = Null2String(RS!invoiceno_orno)
        txtOtherComments.Text = Null2String(RS!othercomments)
        txtDealer.Text = Null2String(RS!DEALERCODE)
        InspectionDate = Null2String(RS!InspectionDate)
        RepairDate = Null2String(RS!RepairDate)
        txtMileage.Text = Null2String(RS!MILEAGE)
        txtPLATENO.Text = Null2String(RS!PlateNo)
        If RS!sentregistrationcard = "Y" Then
            OtpSentCardY.Value = True
        Else
            OtpSentCardN.Value = True
        End If
        txtsaleAdvisor.Text = Null2String(RS!serviceAdvisor)
        cboTECH = Null2String(RS!Technician)

        If RS!con_enginetemp = "Hot" Then
            Engine(0).Value = True
        ElseIf RS!con_enginetemp = "Cold" Then Engine(1).Value = True
        Else
            Engine(2).Value = True
        End If


        If RS!con_weather = "Cold" Then
            Weather(0).Value = True
        ElseIf RS!con_weather = "Warm" Then Weather(1).Value = True
        Else
            Weather(2).Value = True
        End If

        If RS!con_shifting = "Normal" Then
            Shifting(0).Value = True
        Else
            Shifting(1).Value = True
        End If

        If RS!con_shifposition = "4wd" Then
            ShiftPosition(0).Value = True
        Else
            ShiftPosition(0).Value = True
        End If

        If RS!con_MT = "1st" Then
            MT(0).Value = True
        ElseIf RS!con_MT = "2nd" Then MT(1).Value = True
        ElseIf RS!con_MT = "3rd" Then MT(2).Value = True
        ElseIf RS!con_MT = "4th" Then MT(3).Value = True
        ElseIf RS!con_MT = "5th" Then MT(4).Value = True
        ElseIf RS!con_MT = "Neutral" Then MT(5).Value = True
        Else
            MT(6).Value = True
        End If

        If RS!con_AT = "L" Then
            AT(0).Value = True
        ElseIf RS!con_AT = "2" Then AT(1).Value = True
        ElseIf RS!con_AT = "3" Then AT(2).Value = True
        ElseIf RS!con_AT = "D" Then AT(3).Value = True
        ElseIf RS!con_AT = "N" Then AT(4).Value = True
        ElseIf RS!con_AT = "L" Then AT(5).Value = True
        Else
            AT(6).Value = True
        End If

        If RS!con_road = "Paved" Then
            Road(0).Value = True
        ElseIf RS!con_road = "Unpaved" Then Road(1).Value = True
        ElseIf RS!con_road = "Rocky" Then Road(2).Value = True
        Else
            Road(3).Value = True
        End If

        If RS!con_location = "HighWay" Then
            Location(0).Value = True
        ElseIf RS!con_location = "Uphill" Then Location(1).Value = True
        ElseIf RS!con_location = "Downhill" Then Location(2).Value = True
        Else
            Location(3).Value = True
        End If

        If RS!con_action = "Cranking" Then
            Action(0).Value = True
        ElseIf RS!con_action = "idling" Then Action(1).Value = True
        ElseIf RS!con_action = "Cruising" Then Action(2).Value = True
        ElseIf RS!con_action = "Accelarating" Then Action(3).Value = True
        Else
            Action(4).Value = True
        End If

        If RS!con_Occurence = "Consistent" Then
            Occurence(0).Value = True
        Else
            Occurence(1).Value = True
        End If

        If RS!VechicleMaintenance = "dealer" Then
            maintenance(0).Value = True
        ElseIf RS!VechicleMaintenance = "3starshop" Then maintenance(1).Value = True
        ElseIf RS!VechicleMaintenance = "gasStation" Then maintenance(2).Value = True
        Else
            maintenance(3).Value = True
        End If

        If RS!con_accessories = "a/cON" Then
            Accesories(0).Value = True
        Else
            Accesories(1).Value = True
        End If

        txtOthercondition = Null2String(RS!othercondition)
        TxtmaintenaceEvry.Text = Null2String(RS!everymonth)
        TxtmaintenaceKMS.Text = Null2String(RS!everykms)
        txtRequested.Text = Null2String(RS!requestedby)
        txtgeneralManager.Text = Null2String(RS!CheckedBy)
        txtServiceDept.Text = Null2String(RS!ApprovedBy)
        lblID.Caption = Null2String(RS!ID)
        theID = lblID.Caption
    End If
    'listpart_Click
    listpart.Enabled = True
    Exit Sub

ErrorCode:

    ShowVBError
    Exit Sub
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    Call RS.Open("SELECT * FROM CSMS_CQIR", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Sub initMemvars()
    txtRO.Text = "": txtReferenceNo.Text = "":
    txtCustomer.Text = "": txtVIN.Text = ""
    txtEngineNo.Text = "": txtTMAxleNo.Text = ""
    cbojob.Text = "": txthistory.Text = "":
    txtDescription.Text = "": txtAnalysis.Text = ""
    txtCorrectiveAct.Text = "": txtRecommendation.Text = ""
    txtCausalPartNo.Text = "": txtNatureCode.Text = ""
    txtCauseCode.Text = "": txtPaintCode.Text = ""
    txtSubletCode.Text = "": txtTotalLaborCost.Text = ""
    txtTotalSubletRepair.Text = "": txtGrandtotal.Text = ""
    InvoiceNo.Text = "": txtOtherComments.Text = ""
    txtDealer.Text = "": txtMileage.Text = "": txtMileage.Text = ""
    txtPLATENO.Text = "": txtPLATENO.Text = "": txtsaleAdvisor.Text = ""
    cboTECH = "": txtRequested.Text = "": txtgeneralManager.Text = ""
    txtServiceDept = ""
    theID = ""
    txtsub.Text = ""
    thePartNo = ""
    thepartName = ""
    theQty = ""
    thePCost = ""
    theReference = ""
End Sub

Sub displayCustomer()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer

    'SQL = "SELECT Firstname,LastName,customeradd,ro_no,engine,Vin,plate_no,customer,writer from CSMS_vw_CQIRactiveCust WHERE Transtype='R'"
    SQL = "SELECT * FROM CSMS_CQIRWARANTY"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listcust.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = listcust.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!RO_NO)
            ITEM.SubItems(2) = Null2String(!Firstname)
            ITEM.SubItems(3) = Null2String(!lastname)
            ITEM.SubItems(4) = Null2String(!CUSTOMERADD)
            ITEM.SubItems(5) = Null2String(!ACCTNAME)
            ITEM.SubItems(6) = Null2String(!PLATE_NO)
            ITEM.SubItems(7) = Null2String(!Vin)
            ITEM.SubItems(8) = Null2String(!Engine)
            ITEM.SubItems(9) = Null2String(!writer)


            '            CheckIfWcode Item.SubItems(1)
            '            Dim I As Integer
            '            Dim b As Integer
            '            If TheWcode_Mode = True Then
            '
            '                 For I = 1 To listcust.ColumnHeaders.Count - 1
            '                    listcust.ListItems(cnt).ListSubItems(I).Bold = True
            '
            '                 Next
            '
            '                Else
            '                For I = 1 To listcust.ColumnHeaders.Count - 1
            '
            '
            '                    listcust.ListItems(cnt).ListSubItems(I).ForeColor = &H8000&
            '
            '                Next
            '
            '            End If

            .MoveNext



        Loop
    End With
    Set RS = Nothing
End Sub

Sub searchCustomer()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim Keyword                                        As String
    Dim cnt                                            As Integer

    SQL = "SELECT * FROM CSMS_vw_CQIRactivecust WHERE"
    Keyword = Trim(txtkeyword.Text)
    If Len(Keyword) = 0 Then Exit Sub

    If otpSearch(0).Value = True Then
        SQL = SQL & " Firstname LIKE '" & Keyword & "%'"
    ElseIf otpSearch(1).Value = True Then
        SQL = SQL & " Lastname LIKE '" & Keyword & "%'"
    Else
        SQL = SQL & " Ro_no LIKE '%" & Keyword & "%'"
    End If

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cnt = 0
    listcust.ListItems.Clear

    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = listcust.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!RO_NO)
            ITEM.SubItems(2) = Null2String(!Firstname)
            ITEM.SubItems(3) = Null2String(!lastname)
            ITEM.SubItems(4) = Null2String(!CUSTOMERADD)
            .MoveNext
        Loop
    End With
    Set RS = Nothing

End Sub

Sub displayPartsinfo()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim SQLPARTS                                       As String
    Dim TheTotalLTS                                    As Double

    SQL = "SELECT detcde,detdsc,detvol,det_amt,det_hrs  from CSMS_ro_det WHERE  rep_or='" & txtRO.Text & "' and livil='2' and wcode='W'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If RS.EOF And RS.BOF Then
        Exit Sub
    End If

    listpart.ListItems.Clear
    cnt = 0
    thecosttotal = 0
    'theQTYTotal = 0
    With RS
        Do While Not .EOF
            Set ITEM = listpart.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!DETCDE)
            ITEM.SubItems(2) = Null2String(!DETDSC)
            ITEM.SubItems(3) = Null2String(!detvol)
            ITEM.SubItems(4) = Null2String(!DET_AMT)
            ITEM.SubItems(6) = Null2String(!DET_HRS)
            thecosttotal = thecosttotal + N2Str2IntZero(!DET_AMT)
            TheTotalLTS = TheTotalLTS + NumericVal(!DET_HRS)

            If IsParts = True Then
                SQLPARTS = "INSERT INTO CSMS_CQIRparts VALUES(NULL,'" & ITEM.SubItems(1) & "','" & ITEM.SubItems(2) & "','" & ITEM.SubItems(3) & "','" & ITEM.SubItems(4) & "','" & theopcode & _
                           "','" & theLts & "','" & txtRO.Text & "','" & txttotalcost.Text & "','" & TheTotalLTS & "',null,null)"
                gconDMIS.Execute (SQLPARTS)
            End If



            .MoveNext
        Loop
    End With

    txttotalcost.Text = thecosttotal

    txttotalLTS.Text = Format(TheTotalLTS, "#,###,##0.0")
    Set RS = Nothing
End Sub

Sub displaylaborCost()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT labor From CSMS_Repor WHERE Rep_or='" & txtRO.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtTotalLaborCost.Text = Null2String(RS!labor)
    End If
    Set RS = Nothing
End Sub

Sub DeleteRecord()
    Dim RS                                             As New ADODB.Recordset
    Dim SQL                                            As String
    Dim SQLPart                                        As String
    Dim ans                                            As String

    ans = MsgBox("Are You Sure Do you Want to Delete this Record?", vbInformation + vbYesNo, "Information!")
    If ans = vbYes Then
        SQL = "DELETE FROM CSMS_CQIR WHERE ID='" & theID & "'"
        gconDMIS.Execute (SQL)
        MsgBox "All information Has Been Deleted!", vbInformation, "Information"
        SQLPart = "DELETE From CSMS_CQIRPARTS WHERE DLR_CQIR_referenceNo='" & txtReferenceNo.Text & "' "
        gconDMIS.Execute (SQLPart)
        Call initMemvars
        Call rsRefresh
    End If

End Sub

Sub lockedAll(ByVal b As Boolean)

    frameA.Enabled = b
    FrameB.Enabled = b
    FrameC.Enabled = b
    FrameD.Enabled = b
    txtOtherComments.Enabled = b
    txtRequested.Enabled = b
    txtgeneralManager.Enabled = b
    txtServiceDept.Enabled = b
End Sub

Sub GetTheInvoice()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT invoice,plate_no,km_rdg FROM CSMS_Repor Where Rep_or='" & txtRO.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        InvoiceNo.Text = Null2String(RS!invoice)
        txtPLATENO.Text = Null2String(RS!PLATE_NO)
        txtMileage.Text = Null2String(RS!km_rdg)
    End If
    Set RS = Nothing
End Sub

Sub SaveParts()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    SQL = "INSERT INTO CSMS_CQIRparts VALUES('" & txtReferenceNo.Text & "','" & thePartNo & "','" & thepartName & "','" & theQty & "','" & thePCost & "','" & theopcode & _
          "','" & theLts & "','" & txttotalcost.Text & "','" & TheTotalLTS & "')"
    gconDMIS.Execute (SQL)
End Sub

Sub PrintMe()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ishot, iscold, AllTemp                         As String
    Dim Wishot, WisWarm, Walltemp                      As String
    Dim shiftingF, shiftingN                           As String
    Dim Shift4w, Shift2w                               As String
    Dim MT1, MT2, MT3, MT4, MT5, MT6, MTN, MTR         As String
    Dim ATL, AT2, AT3, ATD, ATN, ATR, ATO              As String
    Dim MRoadP, MRoadU, MRoadR, MRoadM                 As String
    Dim MlocationH, MLocationU, MLocation, MLocationD, MLocationS As String
    Dim MActionC, MActioni, MactionCr, MActionA, MActionD As String
    Dim MOccurenceC, MOccurenceI                       As String
    Dim MaintenanceD, Maintenance3, MaintenanceG, MaintenanceO As String
    Dim AccessoriesA, AccessoriesH                     As String
    Dim SentY, SentN                                   As String
    Dim PwaY, PwaN                                     As String
    SQL = "SELECT * FROM CSMS_CQIR WHERE DLR_CQIR_ReferenceNo ='" & txtReferenceNo.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    Screen.MousePointer = 11

    If Engine(0).Value = True Then
        ishot = "X"
    End If
    If Engine(1).Value = True Then
        iscold = "X"
    End If
    If Engine(2).Value = True Then
        AllTemp = "X"
    End If

    If Weather(0).Value = True Then
        Wishot = "X"
    End If
    If Weather(1).Value = True Then
        WisWarm = "X"
    End If
    If Weather(2).Value = True Then
        Walltemp = "X"
    End If

    If Shifting(0).Value = True Then
        shiftingF = "X"
    Else
        shiftingN = "X"
    End If

    If ShiftPosition(0).Value = True Then
        Shift4w = "X"
    Else
        Shift2w = "X"
    End If

    If MT(0).Value = True Then
        MT1 = "X"
    ElseIf MT(1).Value = True Then MT2 = "X"
    ElseIf MT(2).Value = True Then MT3 = "X"
    ElseIf MT(3).Value = True Then MT4 = "X"
    ElseIf MT(4).Value = True Then MT5 = "X"
    ElseIf MT(5).Value = True Then MTN = "X"
    Else
        MTR = "X"
    End If

    If AT(0).Value = True Then
        ATL = "X"
    ElseIf AT(1).Value = True Then AT2 = "X"
    ElseIf AT(2).Value = True Then AT3 = "X"
    ElseIf AT(3).Value = True Then ATD = "X"
    ElseIf AT(4).Value = True Then ATN = "X"
    ElseIf AT(5).Value = True Then ATR = "X"
    Else
        ATO = "X"
    End If

    If Road(0).Value = True Then
        MRoadP = "X"
    ElseIf Road(1).Value = True Then MRoadU = "X"
    ElseIf Road(2).Value = True Then MRoadR = "X"
    ElseIf Road(3).Value = True Then MRoadM = "X"
    End If

    If Location(0).Value = True Then
        MlocationH = "X"
    ElseIf Location(1).Value = True Then MLocationU = "X"
    ElseIf Location(2).Value = True Then MLocationD = "X"
    Else
        MLocationS = "X"
    End If

    If Action(0).Value = True Then
        MActionC = "X"
    ElseIf Action(1).Value = True Then MActioni = "X"
    ElseIf Action(2).Value = True Then MactionCr = "X"
    ElseIf Action(3).Value = True Then MActionA = "X"
    ElseIf Action(4).Value = True Then MActionD = "X"
    End If

    If Occurence(0).Value = True Then
        MOccurenceC = "X"
    Else
        MOccurenceI = "X"
    End If

    If maintenance(0).Value = True Then
        MaintenanceD = "X"
    ElseIf maintenance(1).Value = True Then Maintenance3 = "X"
    ElseIf maintenance(2).Value = True Then MaintenanceG = "X"
    Else
        MaintenanceO = "X"
    End If

    If Accesories(0).Value = True Then
        AccessoriesA = "X"
    Else
        AccessoriesH = "X"
    End If

    If OtpSentCardY.Value = True Then
        SentY = "X"
    ElseIf OtpSentCardN.Value = True Then
        SentN = "X"
    End If

    If otpRequestY.Value = True Then
        PwaY = "X"
    Else
        PwaN = "X"
    End If

    rptCQIRReport.Formulas(1) = "engineIshot='" & ishot & "'"
    rptCQIRReport.Formulas(2) = "engineIscold='" & iscold & "'"
    rptCQIRReport.Formulas(3) = "engineIsalltemp='" & AllTemp & "'"

    rptCQIRReport.Formulas(4) = "WeatherIshot='" & Wishot & "'"
    rptCQIRReport.Formulas(5) = "WeatherIsWarm='" & WisWarm & "'"
    rptCQIRReport.Formulas(6) = "weatheralltemp='" & Walltemp & "'"

    rptCQIRReport.Formulas(7) = "shiftingIsFast='" & shiftingF & "'"
    rptCQIRReport.Formulas(8) = "shiftingIsnormal='" & shiftingN & "'"

    rptCQIRReport.Formulas(9) = "shifting4wd='" & Shift4w & "'"
    rptCQIRReport.Formulas(10) = "shifting2wd='" & Shift2w & "'"

    rptCQIRReport.Formulas(11) = "MT1='" & MT1 & "'"
    rptCQIRReport.Formulas(12) = "MT2='" & MT2 & "'"
    rptCQIRReport.Formulas(13) = "MT3='" & MT3 & "'"
    rptCQIRReport.Formulas(14) = "MT4='" & MT4 & "'"
    rptCQIRReport.Formulas(15) = "MT5='" & MT5 & "'"
    rptCQIRReport.Formulas(16) = "MTN='" & MTN & "'"
    rptCQIRReport.Formulas(17) = "MTR='" & MTR & "'"

    rptCQIRReport.Formulas(18) = "ATL='" & ATL & "'"
    rptCQIRReport.Formulas(19) = "AT2='" & AT2 & "'"
    rptCQIRReport.Formulas(20) = "AT3='" & AT3 & "'"
    rptCQIRReport.Formulas(21) = "ATD='" & ATD & "'"
    rptCQIRReport.Formulas(22) = "ATN='" & ATN & "'"
    rptCQIRReport.Formulas(23) = "ATR='" & ATR & "'"
    rptCQIRReport.Formulas(24) = "ATO='" & ATO & "'"

    If Road(0).Value = True Then
        MRoadP = "X"
    ElseIf Road(1).Value = True Then MRoadU = "X"
    ElseIf Road(2).Value = True Then MRoadR = "X"
    ElseIf Road(3).Value = True Then MRoadM = "X"
    End If

    rptCQIRReport.Formulas(25) = "RoadP='" & MRoadP & "'"
    rptCQIRReport.Formulas(26) = "RoadU='" & MRoadU & "'"
    rptCQIRReport.Formulas(27) = "RoadR='" & MRoadR & "'"
    rptCQIRReport.Formulas(28) = "RoadM='" & MRoadM & "'"



    rptCQIRReport.Formulas(29) = "LocationH='" & MlocationH & "'"
    rptCQIRReport.Formulas(30) = "LocationU='" & MLocationU & "'"
    rptCQIRReport.Formulas(31) = "LocationD='" & MLocationD & "'"
    rptCQIRReport.Formulas(32) = "LocationS='" & MLocationS & "'"

    rptCQIRReport.Formulas(33) = "ActionC='" & MActionC & "'"
    rptCQIRReport.Formulas(34) = "Actioni='" & MActioni & "'"
    rptCQIRReport.Formulas(35) = "ActionCr='" & MactionCr & "'"
    rptCQIRReport.Formulas(36) = "ActionA='" & MActionA & "'"
    rptCQIRReport.Formulas(37) = "ActionD='" & MActionD & "'"

    rptCQIRReport.Formulas(38) = "occurrenceC='" & MOccurenceC & "'"
    rptCQIRReport.Formulas(39) = "occurrenceI='" & MOccurenceI & "'"

    rptCQIRReport.Formulas(40) = "AccessoriesA='" & MOccurenceC & "'"
    rptCQIRReport.Formulas(41) = "AccessoriesH='" & MOccurenceI & "'"

    rptCQIRReport.Formulas(42) = "MaintenanceD='" & MaintenanceD & "'"
    rptCQIRReport.Formulas(43) = "Maintenance3='" & Maintenance3 & "'"
    rptCQIRReport.Formulas(44) = "MaintenanceG='" & MaintenanceG & "'"
    rptCQIRReport.Formulas(45) = "MaintenanceO='" & MaintenanceO & "'"

    rptCQIRReport.Formulas(46) = "sentCardY='" & SentY & "'"
    rptCQIRReport.Formulas(47) = "SentCardN='" & SentN & "'"

    rptCQIRReport.Formulas(48) = "PWAY='" & PwaY & "'"
    rptCQIRReport.Formulas(49) = "PWAN='" & PwaN & "'"

    PrintSQLReport rptCQIRReport, CSMS_REPORT_PATH & "CQIRReport.rpt", "{CSMS_CQIR.DLR_CQIR_ReferenceNo}='" & txtReferenceNo.Text & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub checkIfnoREc()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT * FROM CSMS_CQIR"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    With RS

        If .EOF And .BOF Then
            ShowNoRecord
        End If

        If Not .EOF And Not .BOF Then
        End If

    End With
    Set RS = Nothing
End Sub

Sub CheckIfWcode(theRo)
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT Rep_or ,wcode From CSMS_Ro_det Where Rep_or='" & theRo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    TheWcode_Mode = False
    With RS
        Do While Not .EOF
            If (RS!wCode) = "W" Then

                TheWcode_Mode = True
            Else: TheWcode_Mode = False
            End If
            .MoveNext
        Loop
    End With
End Sub

Sub filjob()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim tmp                                            As String

    SQL = "SELECT DETDSC from CSMS_Ro_det where rep_or='" & txtRO.Text & "' and livil='1' and Wcode ='W'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    cbojob.Clear
    Do While Not RS.EOF
        cbojob.AddItem RS!DETDSC
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub computeLTS()


    Dim lts                                            As Double
    Dim totalLabor                                     As Double

    lts = txttotalLTS.Text


    txtTotalLaborCost = CCur(txtCost.Text) * CCur(lts)


End Sub

Sub displaySubletparts()
    Dim SQL                                            As String
    Dim arnie                                          As ListItem
    Dim RS                                             As New ADODB.Recordset


    SQL = "SELECT * FROM CSMS_CQIRPARTS where RO='" & txtRO.Text & "' and issublet='S'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)


    thesubletTotal = 0
    listsublet.ListItems.Clear

    Do While Not RS.EOF

        Set arnie = listsublet.ListItems.Add(, , RS!partname)
        arnie.SubItems(1) = Null2String(RS!cost)
        thesubletTotal = thesubletTotal + N2Str2IntZero(RS!cost)
        RS.MoveNext

    Loop

    txtpassSublet.Text = thesubletTotal

    Set RS = Nothing


End Sub

Sub computeTheGrandTotal()
    txtGrandtotal.Text = Val(txttotalcost.Text) + Val(txtTotalLaborCost.Text) + Val(txtpassSublet.Text)
End Sub

Sub makeaNuLL()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim SQLPARTS                                       As String
    Dim TheTotalLTS                                    As Double

    SQL = "SELECT detcde,detdsc,detvol,det_amt,det_hrs  from CSMS_ro_det WHERE  rep_or='" & txtRO.Text & "' and livil='2' and wcode='W'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    isdeletelist.ListItems.Clear
    cnt = 0
    thecosttotal = 0

    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = isdeletelist.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!DETCDE)
            ITEM.SubItems(2) = Null2String(!DETDSC)
            ITEM.SubItems(3) = Null2String(!detvol)
            ITEM.SubItems(4) = Null2String(!DET_AMT)
            ITEM.SubItems(6) = Null2String(!DET_HRS)
            thecosttotal = thecosttotal + N2Str2IntZero(!DET_AMT)
            TheTotalLTS = TheTotalLTS + NumericVal(!DET_HRS)

            SQLPARTS = "INSERT INTO CSMS_CQIRparts VALUES(NULL,'" & ITEM.SubItems(1) & "','" & ITEM.SubItems(2) & "','" & ITEM.SubItems(3) & "','" & ITEM.SubItems(4) & "','" & theopcode & _
                       "','" & theLts & "','" & txtRO.Text & "','" & txttotalcost.Text & "','" & TheTotalLTS & "',null,null)"
            gconDMIS.Execute (SQLPARTS)


            .MoveNext
        Loop
    End With


    Set RS = Nothing
End Sub

Private Sub cbojob_Change()
    filjob
End Sub

Private Sub cmdaddsub_Click()
    Dim SQLP                                           As String
    Dim RS                                             As New ADODB.Recordset
    Dim xsubletPartNo                                  As String
    Dim xsubletlts                                     As String
    Dim xsubletqty                                     As String
    Dim totalLTS                                       As Double



    xsubletPartNo = txtPartNo.Text
    xsubletqty = txtQty.Text

    If txtsubletLTS.Text = "" Then
        txtsubletLTS.Text = 0
    End If
    theopcode = txtsubletopcode.Text
    xsubletlts = txtsubletLTS.Text
    totalLTS = txtsubletLTS.Text
    If txtSubletDesc.Text = "" Then
        MsgBox "Pls input a description..", vbExclamation, "WARNING"
        txtSubletDesc.SetFocus
        Exit Sub
    End If



    If txtReferenceNo.Text = "" Then
        MsgBox "Input DLR CQIR Reference #..", vbInformation, "WARNING"
        txtReferenceNo.SetFocus
        Exit Sub
    End If

    If txtsubletprc.Text = "" Then
        MsgBox "Pls Input sublet price..", vbExclamation, "WARNING"
        txtsubletprc.SetFocus
        Exit Sub

    End If
    SQLP = "INSERT INTO CSMS_CQIRparts VALUES('" & txtReferenceNo & "','" & xsubletPartNo & "','" & txtSubletDesc.Text & "','" & xsubletqty & "','" & txtsubletprc.Text & "','" & theopcode & _
           "','" & xsubletlts & "','" & txtRO.Text & "',null,null,'S','YES')"


    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQLP)



    displaySubletparts
    txtpassSublet.Text = thesubletTotal
    computeTheGrandTotal

    MsgBox "Sublet has been save!.", vbInformation, "Confirm"
    UPDATE_MODE = True
    Saveinformation
    picsublet.Visible = False
    txtSubletDesc.Text = ""
    txtsubletprc.Text = ""
    txttotalLTS.Text = txttotalLTS.Text + totalLTS
    computeLTS
End Sub

Private Sub cmdarnie_Click()
    picsublet.Visible = False
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next

    If Not RS.EOF And Not RS.BOF Then
        thepicSave.Visible = False
        thepicControl.Visible = True
        thepic.Visible = False
        txtReferenceNo.Locked = True
        cmdSave.Caption = "Save"
        initMemvars
        lockedAll (False)
    Else
        ShowNoRecord
    End If
End Sub

Private Sub cmdCompute_Click()
    txtGrandtotal.Text = Val(txttotalcost.Text) + Val(txtTotalLaborCost.Text) + Val(txtpassSublet.Text)
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "QUALITY INFORMATION") = False Then Exit Sub
    Call DeleteRecord
    LogAudit "X", "CQIR REPORT", "PWA/RO" & TxtPWANo & "/" & txtRO
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "QUALITY INFORMATION") = False Then Exit Sub
    On Error Resume Next
    UPDATE_MODE = True
    cmdSave.Caption = "Update"
    thepicSave.Visible = True
    thepicControl.Visible = False
    lockedAll (True)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    RS.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdFind_Click()
    Call initMemvars
    cmdSave.Caption = "Save"
    thepic.Visible = True
    thepic.ZOrder 0
    displayCustomer

End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    RS.MoveLast
    StoreMemVars
End Sub

Private Sub cmdme_Click()
    displaySubletparts

End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub CmdOpCode_Click()
    Dim SQL                                            As String
    Dim theopcode                                      As String
    Dim theLts                                         As String
    Dim SqlLTS                                         As String
    Dim ans                                            As String
    theopcode = txtOPCODE.Text
    theLts = txtLTS.Text
    ans = MsgBox("Are you sure do you want to save the opcode?", vbQuestion + vbYesNo, "Information")
    If ans = vbYes Then
        If txtOPCODE.Text = "" Then
            MsgBox "Pls input Opcode!!", vbExclamation, "Warning"
            Exit Sub
        End If
        SQL = "update CSMS_CQIRPARTS set opcode='" & theopcode & "',lts='" & theLts & "' where partno='" & txtCausalPartNo & "'"
        gconDMIS.Execute (SQL)
        SqlLTS = "update CSMS_RO_Det set det_hrs='" & theLts & "' where rep_or='" & txtRO.Text & "' and detcde='" & txtCausalPartNo & "'"
        gconDMIS.Execute (SqlLTS)
        Call displayPartsinfo
        MsgBox "OpCode Has Been Save!", vbInformation, "Confirm"
        opCodePic.Visible = False
        txtOPCODE.Text = ""
    Else
        txtOPCODE.SetFocus
    End If

    Call computeLTS

End Sub

Private Sub cmdprev_Click()
    On Error Resume Next
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    On Error Resume Next
    'cbojob.ListIndex = 0
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "QUALITY INFORMATION") = False Then Exit Sub
    Dim ans                                            As String
    Dim ans1                                           As String
    ans = MsgBox("Are You sure Do you want to Print This Record?", vbQuestion + vbYesNo, "Information")
    If ans = vbYes Then
        'cmdprintOk.Enabled = False
        'listpartoption.CheckBoxes = True
        'picdelete.Visible = True

        PrintMe

        LogAudit "V", "CQIR REPORT", "PWA/RO" & TxtPWANo & "/" & txtRO
        'gconDMIS.Execute "UPDATE CSMS_CQIRPARTS set isdelete=null where partno='" & thePartNo & "'"

    End If
End Sub

Private Sub cmdSave_Click()
    Dim ans                                            As String

    ans = MsgBox("Are you sure do you want to save this record?", vbQuestion + vbYesNo, "information")
    If ans = vbYes Then
        Saveinformation
        listpart.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "QUALITY INFORMATION") = False Then Exit Sub
    Call initMemvars
    txtReferenceNo.Locked = False
    cmdSave.Caption = "Save"
    thepic.Visible = True
    thepic.ZOrder 0
    displayCustomer

    thepicSave.Visible = True
    thepicControl.Visible = False
    listpart.Enabled = False
End Sub

Private Sub cmdPrevious_Click()

End Sub

Private Sub Command1_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub Command2_Click()
    Saveinformation
End Sub

Private Sub Command3_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub Command5_Click()
    Call initMemvars
    thepicControl.Visible = False
    thepicSave.Visible = True
    thepic.Visible = True
    thepic.ZOrder 0
    displayCustomer
End Sub

Private Sub Command6_Click()
    UPDATE_MODE = True
    thepicControl.Visible = False
    thepicSave.Visible = True

End Sub

Private Sub cmdsublet_Click()


    picsublet.Visible = True

End Sub

Private Sub cmdx_Click()
    picdelete.Visible = False
End Sub

Private Sub Command8_Click()
    On Error Resume Next
    RS.MoveLast
    StoreMemVars
End Sub

Private Sub Command9_Click()
    RS.MoveFirst
    StoreMemVars

End Sub

Private Sub Edt_Click()
    txtCost.Locked = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        If opCodePic.Visible = True Then
            Dim SQL                                    As String
            Dim theopcode                              As String
            Dim theLts                                 As String
            Dim SqlLTS                                 As String
            Dim ans                                    As String
            theopcode = txtOPCODE.Text
            theLts = txtLTS.Text
            ans = MsgBox("Are you sure do you want to save the opcode?", vbQuestion + vbYesNo, "Information")
            If ans = vbYes Then
                If txtOPCODE.Text = "" Then
                    MsgBox "Pls input Opcode!!", vbExclamation, "Warning"
                    Exit Sub
                End If
                SQL = "update CSMS_CQIRPARTS set opcode='" & theopcode & "',lts='" & theLts & "' where partno='" & txtCausalPartNo & "'"
                gconDMIS.Execute (SQL)
                SqlLTS = "update CSMS_RO_Det set det_hrs='" & theLts & "' where rep_or='" & txtRO.Text & "' and detcde='" & txtCausalPartNo & "'"
                gconDMIS.Execute (SqlLTS)
                Call displayPartsinfo
                MsgBox "OpCode Has Been Save!", vbInformation, "Confirm"
                opCodePic.Visible = False
                txtOPCODE.Text = ""
                Call computeLTS
                computeTheGrandTotal
                UPDATE_MODE = True
                Saveinformation
            Else
                txtOPCODE.SetFocus
            End If



        End If
    End If

    If KeyCode = vbKeyEscape Then
        If opCodePic.Visible = True Then
            opCodePic.Visible = False
        End If
    End If


    If KeyCode = vbKeyF3 = True Then
        picsublet.Visible = True
    End If


End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Me.Height = Screen.TwipsPerPixelY * 585
    picunder.Height = Me.ScaleHeight - Pictops.Height - Picbottoms.Height
    ScrollBar1.Max = Abs(picunder.ScaleHeight - Picover.Height) + 20
    rsRefresh
    txtReferenceNo.Locked = True
    thepicSave.Visible = False
    CenterMe frmMain, Me, 1
    Call StoreMemVars
    FillCBOTechnician
    UPDATE_MODE = False
    thepic.Visible = False
    lockedAll (False)
    InvoiceNo.Enabled = False
    IsParts = False
    opCodePic.Visible = False

    If RS.BOF And RS.EOF Then
        cmdAdd_Click
        checkIfnoREc
    End If

    picsublet.Visible = False
    listpart.Enabled = True
    picdelete.Visible = False

End Sub

Private Sub isdeletelist_ItemCheck(ByVal ITEM As MSComctlLib.ListItem)
    If ITEM.Checked = True Then
        cmdprintOk.Enabled = True
        'Exit Sub

    End If
    For Each ITEM In isdeletelist.ListItems
        If ITEM.Checked = True Then

            Exit Sub
        End If
    Next

End Sub

Private Sub lblclose_Click()

    If Not RS.EOF And Not RS.BOF Then
        thepic.Visible = False
        thepicSave.Visible = False
        thepicControl.Visible = True
        initMemvars
        txtReferenceNo.Locked = True
    End If
End Sub

Private Sub listcust_DblClick()
    Dim Customer, Vin                                  As String

    txtRO.Text = listcust.SelectedItem.SubItems(1)
    Customer = listcust.SelectedItem.SubItems(5)
    txtPLATENO = listcust.SelectedItem.SubItems(6)
    Vin = listcust.SelectedItem.SubItems(7)
    txtEngineNo.Text = listcust.SelectedItem.SubItems(8)
    txtsaleAdvisor.Text = listcust.SelectedItem.SubItems(9)
    thepic.Visible = False

    txtCustomer.Text = Customer
    txtVIN.Text = Vin



    filjob
    displaylaborCost
    lockedAll (True)
    On Error Resume Next
    TxtPWANo.SetFocus
    'listpart_Click
    listpart.Visible = True
End Sub

Private Sub listpart_Click()
    On Error Resume Next

    thePartNo = listpart.SelectedItem.SubItems(1)
    thepartName = listpart.SelectedItem.SubItems(2)
    theQty = listpart.SelectedItem.SubItems(3)
    thePCost = listpart.SelectedItem.SubItems(4)
    theReference = listpart.SelectedItem.SubItems(6)
    txtCausalPartNo = thePartNo
End Sub

Private Sub listpart_DblClick()
    DefaultLabor = 247.5
    txtCost.Text = DefaultLabor
    opCodePic.Visible = True
    txtLTS.SetFocus
    txtCost.Locked = True
End Sub

Private Sub listpart_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu MNUREMOVE
        Dim xdelete                                    As String
        Dim ans                                        As String

        'If pindot = True Then Exit Sub

        xdelete = "YES"
        ans = MsgBox("Are you sure do you want to include this in CQIR?", vbQuestion + vbYesNo)
        'this routin is for temporary deletion of part in printing porpuse only
        If ans = vbYes Then
            gconDMIS.Execute "UPDATE CSMS_CQIRPARTS set isdelete='" & xdelete & "' where partno='" & thePartNo & "'"
            MsgBox "All Information has been Added.", vbInformation, "Information"
        End If
    End If

    ' pindot = False

End Sub

Private Sub listpartoption_ItemCheck(ByVal ITEM As MSComctlLib.ListItem)
    If ITEM.Checked = True Then

    End If

End Sub

Private Sub mnuref_Click()
    'pindot = True


    gconDMIS.Execute "UPDATE CSMS_CQIRPARTS set isdelete=null where partno='" & thePartNo & "'"
    MsgBox "All information has been refresh.", vbInformation, "Information"

End Sub

Private Sub ScrollBar1_Change()
    Picover.Top = 0 - ScrollBar1.Value

End Sub

Private Sub txtKeyword_Change()
    searchCustomer
End Sub

Private Sub txtRO_Change()
    displayPartsinfo
    displaySubletparts
End Sub


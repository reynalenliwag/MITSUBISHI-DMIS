VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Trans_Ordering 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Purchase Order"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PurchaseOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   10875
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   3090
      ScaleHeight     =   915
      ScaleWidth      =   7740
      TabIndex        =   60
      Top             =   6390
      Width           =   7740
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
         Height          =   795
         Left            =   6870
         MouseIcon       =   "PurchaseOrder.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   705
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
         Height          =   795
         Left            =   6180
         MouseIcon       =   "PurchaseOrder.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdCancelCO 
         Caption         =   "Cancel Transaction"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5430
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PurchaseOrder.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Cancel this Transaction"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Transaction"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4680
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PurchaseOrder.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Post this Transaction"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost Transaction"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3930
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PurchaseOrder.frx":1B3D
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":1C8F
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Unpost this Transaction"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton cmdEdit 
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
         Left            =   3240
         MouseIcon       =   "PurchaseOrder.frx":1FD4
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":2126
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
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
         Left            =   2550
         MouseIcon       =   "PurchaseOrder.frx":2482
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":25D4
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
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
         Left            =   1860
         MouseIcon       =   "PurchaseOrder.frx":28E7
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":2A39
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
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
         Left            =   1170
         MouseIcon       =   "PurchaseOrder.frx":2D33
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":2E85
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
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
         Left            =   480
         MouseIcon       =   "PurchaseOrder.frx":31DD
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":332F
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7365
      Left            =   0
      ScaleHeight     =   7365
      ScaleWidth      =   3045
      TabIndex        =   0
      Top             =   0
      Width           =   3045
      Begin VB.Frame fraSearch 
         Height          =   7365
         Left            =   30
         TabIndex        =   1
         Top             =   -30
         Width           =   2955
         Begin VB.TextBox txtPODetailID 
            Alignment       =   2  'Center
            Height          =   495
            Left            =   1725
            TabIndex        =   6
            Text            =   "0"
            Top             =   840
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox txtID 
            Height          =   495
            Left            =   1710
            TabIndex        =   3
            Top             =   210
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.OptionButton optDate 
            Caption         =   "D&ate"
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
            Left            =   90
            TabIndex        =   7
            Top             =   960
            Width           =   1845
         End
         Begin VB.OptionButton optPO 
            Caption         =   "&PO Number"
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
            Left            =   90
            TabIndex        =   4
            Top             =   390
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton optVModel 
            Caption         =   "&Vehicle Model[Description]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   5
            Top             =   630
            Width           =   2835
         End
         Begin VB.TextBox textSearch 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   60
            MaxLength       =   35
            TabIndex        =   8
            Text            =   "TEXT"
            Top             =   1530
            Width           =   2835
         End
         Begin MSComctlLib.ListView lstPO 
            Height          =   5415
            Left            =   60
            TabIndex        =   9
            Top             =   1890
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   9551
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
            MouseIcon       =   "PurchaseOrder.frx":368E
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PO"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   2540
            EndProperty
         End
         Begin Crystal.CrystalReport rptPO 
            Left            =   1350
            Top             =   1260
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label Label22 
            Caption         =   "Search by:"
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
            Left            =   90
            TabIndex        =   2
            Top             =   150
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox picSaves 
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
      Left            =   9210
      ScaleHeight     =   885
      ScaleWidth      =   2130
      TabIndex        =   71
      Top             =   6390
      Width           =   2130
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
         Height          =   795
         Left            =   780
         MouseIcon       =   "PurchaseOrder.frx":37F0
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":3942
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   90
         MouseIcon       =   "PurchaseOrder.frx":3C80
         MousePointer    =   99  'Custom
         Picture         =   "PurchaseOrder.frx":3DD2
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   3000
      ScaleHeight     =   6375
      ScaleWidth      =   7875
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   7875
      Begin VB.TextBox txtDatePO 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5700
         TabIndex        =   15
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txtDueDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5700
         TabIndex        =   16
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox txtPONO 
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
         Left            =   90
         TabIndex        =   14
         Top             =   420
         Width           =   2265
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   7590
         ScaleHeight     =   195
         ScaleWidth      =   7605
         TabIndex        =   27
         Top             =   570
         Width           =   7605
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F3 - Add "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   1530
            TabIndex        =   28
            Top             =   0
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F4 - Edit "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   2310
            TabIndex        =   29
            Top             =   0
            Width           =   630
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "F12 - Un-Post Transaction"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   5700
            TabIndex        =   32
            Top             =   0
            Width           =   2445
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "F8 - Post Transaction"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   4050
            TabIndex        =   31
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F5 - Delete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3090
            TabIndex        =   30
            Top             =   0
            Width           =   780
         End
      End
      Begin VB.ComboBox cboModeOfPayment 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   345
         ItemData        =   "PurchaseOrder.frx":4122
         Left            =   90
         List            =   "PurchaseOrder.frx":4135
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   2550
         Width           =   2850
      End
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   4530
         Picture         =   "PurchaseOrder.frx":417F
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Select Customer"
         Top             =   6390
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtCusCode 
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
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   6390
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtSource 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2970
         TabIndex        =   19
         Top             =   1920
         Width           =   1800
      End
      Begin VB.TextBox txtFuel 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6240
         TabIndex        =   21
         Top             =   1920
         Width           =   1350
      End
      Begin VB.TextBox txtModelYear 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4800
         TabIndex        =   20
         Top             =   1920
         Width           =   1395
      End
      Begin VB.TextBox txtModelCode 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4170
         TabIndex        =   36
         Top             =   1290
         Width           =   1515
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1290
         Width           =   1965
      End
      Begin VB.ComboBox cboModelDescript 
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
         Height          =   345
         Left            =   90
         TabIndex        =   17
         Text            =   "txtDescript"
         Top             =   1290
         Width           =   4065
      End
      Begin VB.ComboBox cboColor 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   90
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   1920
         Width           =   2850
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   30
         Top             =   6570
      End
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   90
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   5400
         Width           =   7635
      End
      Begin VB.Frame fraCheckDetail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   3450
         TabIndex        =   23
         Top             =   2700
         Width           =   4305
         Begin VB.TextBox txtSubsidy 
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
            Height          =   405
            Left            =   1710
            TabIndex        =   25
            Top             =   1920
            Width           =   2505
         End
         Begin VB.TextBox txtPy_CD_Amount 
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
            Height          =   405
            Left            =   1710
            TabIndex        =   24
            Top             =   1470
            Width           =   2505
         End
         Begin VB.TextBox txtPy_CD_CheckNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   630
            Width           =   2505
         End
         Begin VB.ComboBox cboPy_CD_BankName 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00000000&
            Height          =   345
            ItemData        =   "PurchaseOrder.frx":4349
            Left            =   1710
            List            =   "PurchaseOrder.frx":434B
            TabIndex        =   49
            Text            =   "Combo1"
            Top             =   210
            Width           =   2505
         End
         Begin MSComCtl2.DTPicker txPy_CD_Date 
            Height          =   375
            Left            =   1710
            TabIndex        =   53
            Top             =   1050
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   57081857
            CurrentDate     =   39248
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Subsidy"
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
            Left            =   990
            TabIndex        =   74
            Top             =   2040
            Width           =   675
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   1275
            TabIndex        =   52
            Top             =   1080
            Width           =   390
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check No."
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
            Left            =   810
            TabIndex        =   50
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
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
            Left            =   1230
            TabIndex        =   48
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   1005
            TabIndex        =   54
            Top             =   1530
            Width           =   660
         End
      End
      Begin VB.Frame fraCrNo 
         Caption         =   "Financing Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   90
         TabIndex        =   43
         Top             =   2970
         Width           =   3345
         Begin VB.ComboBox cboPy_FinLcIssuingBank 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00000000&
            Height          =   345
            ItemData        =   "PurchaseOrder.frx":434D
            Left            =   120
            List            =   "PurchaseOrder.frx":434F
            TabIndex        =   45
            Text            =   "Combo1"
            Top             =   495
            Width           =   2970
         End
         Begin VB.TextBox txtPy_LCNo 
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
            Left            =   120
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1140
            Width           =   2985
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Financing /LC Issuing Bank"
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
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   2265
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Letter of Credit Number"
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
            Left            =   120
            TabIndex        =   46
            Top             =   885
            Width           =   1995
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date :"
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
         Left            =   5190
         TabIndex        =   13
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Due Date:"
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
         Left            =   4860
         TabIndex        =   26
         Top             =   570
         Width           =   810
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "POSTED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2370
         TabIndex        =   12
         Top             =   60
         Width           =   2715
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "PO No"
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   150
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mode of Payment "
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
         Left            =   90
         TabIndex        =   42
         Top             =   2310
         Width           =   1515
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cus Code"
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
         Left            =   1995
         TabIndex        =   57
         Top             =   6450
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Source"
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
         Left            =   2970
         TabIndex        =   39
         Top             =   1695
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fuel"
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
         Left            =   6240
         TabIndex        =   41
         Top             =   1695
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         Left            =   4800
         TabIndex        =   40
         Top             =   1695
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Left            =   4200
         TabIndex        =   34
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Description"
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
         Left            =   120
         TabIndex        =   33
         Top             =   1020
         Width           =   1530
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5730
         TabIndex        =   35
         Top             =   1020
         Width           =   510
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Left            =   90
         TabIndex        =   38
         Top             =   1695
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Notes"
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
         Left            =   90
         TabIndex        =   55
         Top             =   5160
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_Ordering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPO                                                              As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim DontChange                                                        As Boolean
Dim rsParts                                                           As ADODB.Recordset
Dim rsS_Model                                                         As ADODB.Recordset
Dim rsColor                                                           As ADODB.Recordset
Dim WithEvents SearchMaster                                           As frmSMIS_Mis_SearchMaster
Attribute SearchMaster.VB_VarHelpID = -1

Private Sub cboModelDescript_Change()
    If cboModelDescript.ListIndex <> -1 And DontChange = False Then
        SetModelLine cboModelDescript, False
    End If
    DontChange = False
End Sub

Private Sub cboModelDescript_Click()
    cboModelDescript_Change
End Sub

Private Sub cboModeOfPayment_Change()
    If SetModeOfPayment(cboModeOfPayment) = "CA" Then
        fraCrNo.Enabled = False
        '       fraCheckDetail.Enabled = False
    Else
        fraCrNo.Enabled = True
        '     fraCheckDetail.Enabled = True
    End If
End Sub

Private Sub cboModeOfPayment_Click()
    cboModeOfPayment_Change
End Sub

Private Sub cmdADD_Click()
    If Function_Access(LOGID, "Acess_ADD", "PURCHASE ORDER") = False Then Exit Sub
    AddorEdit = "ADD"
    initMemvars
    txtID = 0
    txtPONO = GenerateCode("SMIS_PO", "PO_NO", "000000")
    txtDatePO = Format(Now, "mm/dd/yyyy")
    picAdds.Visible = False
    picSaves.Visible = True
    fraCrNo.Enabled = True
    fraCheckDetail.Enabled = True
    picTop.Enabled = True
    fraSearch.Enabled = False
    On Error Resume Next
    txtPONO.SetFocus
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picAdds.Visible = True
    picSaves.Visible = False
    fraCrNo.Enabled = False
    fraCheckDetail.Enabled = False
    picTop.Enabled = False
    fraSearch.Enabled = True
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200713:26
Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo Errorcode:

    If MsgBox("Do you Want to Cancel this Transaction ", vbOKCancel + vbExclamation, "Confirm Posting") = vbCancel Then Exit Sub
    cmdCancelCO.Enabled = True
    gconDMIS.Execute ("UPDate PMIS_DPIHeader Set Status='C'  Where ID=" & txtID)
    rsRefresh
    rsPO.Find ("ID=" & txtID)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Cancelled", "Transaction Sucessfully Cancelled"





    Exit Sub
Errorcode:
    ShowVBError
End Sub


Private Sub cmdDelTechnicalInquiry_Click()
'    gconDMIS.Execute ("DELETE From PMIS_DPIDETAILS WHERE ID=" & txtPODetailID)
'    cleargrid Grid1
'    FillGrid
'    ShowHidePictureBox2 picInquiryTechincal, False
    Form_KeyDown 116, 1
End Sub

'Upating Code       : AXP-0707200713:26
Private Sub cmdEdit_Click()
    If lblStatus <> "" Then Exit Sub
    If Function_Access(LOGID, "Acess_EDIT", "PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo Errorcode:

    If NumericVal(txtID) <> 0 Then
        AddorEdit = "EDIT"
        picAdds.Visible = False
        picSaves.Visible = True
        fraCrNo.Enabled = True
        fraCheckDetail.Enabled = True
        picTop.Enabled = True
        On Error Resume Next
        cboModelDescript.SetFocus
    End If



    fraSearch.Enabled = False

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    rsPO.MoveNext

    If rsPO.EOF Then
        rsPO.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars



End Sub

'Upating Code       : AXP-0707200713:26
Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo Errorcode:

    If MsgBox("Do you Want to Post this Transaction ", vbOKCancel + vbExclamation, "Confirm Posting") = vbCancel Then Exit Sub
    cmdCancelCO.Enabled = False
    gconDMIS.Execute ("UPDate SMIS_PO  Set Status='P' Where ID=" & txtID)
    rsRefresh
    rsPO.Find ("ID=" & txtID)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Posted", "Transaction Sucessfully Posted"





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()


    rsPO.MovePrevious

    If rsPO.BOF Then
        rsPO.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars


End Sub

'Upating Code       : AXP-0707200713:26
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    Screen.MousePointer = 11
    '    rptPO.Formulas(0) = "SourceName = 'HUYNDAI ASIA RESOURCES INC (HARI)" & Chr(13) & "315 Sen Gil Puyat Ave., Makati City'"
    Dim rs                                                            As ADODB.Recordset
    Set rs = gconDMIS.Execute("Select * from SMIS_SIGNATORIES")
    If Not rs.EOF Or Not rs.BOF Then
        rptPO.Formulas(0) = "ABY = '" & Null2String(rs!SalesApproved) & "'"
        rptPO.Formulas(1) = "PBY= '" & Null2String(rs!PreparedBy) & "'"
        rptPO.Formulas(2) = "AVP= '" & Null2String(rs!GeneralManager) & "'"
        rptPO.Formulas(3) = "FMAN= '" & Null2String(rs!CheckedBy) & "'"
    End If

    rptPO.Formulas(4) = "COMPANYNAME= '" & Company_name & "'"
    PrintSQLReport rptPO, SMIS_REPORT_PATH & "PO.rpt", "{SMIS_PO.PO_NO}='" & txtPONO.Text & "'", DMIS_Connection, 1
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:26
Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo Errorcode:

    If MsgBox("Do you Want to Unpost this Transaction ", vbOKCancel + vbExclamation, "Confirm Posting") = vbCancel Then Exit Sub

    cmdCancelCO.Enabled = True
    gconDMIS.Execute ("UPDate SMIS_PO  Set Status='U' Where ID=" & txtID)

    rsRefresh

    rsPO.Find ("ID=" & txtID)

    StoreMemVars
    MessagePop RecSaveOk, "Transaction Unposted", "Transaction Sucessfully Unposted"





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    SearchMaster.SearchForCustomers
    SearchMaster.Show 1

End Sub

'Upating Code       : AXP-0707200713:27
Private Sub cmdSave_Click()

    On Error GoTo Errorcode:

    If RTrim(LTrim(txtPONO)) = "" Then
        MessagePop RecSaveError, "MISSING FIELDS", "PO NUMBER"
        On Error Resume Next
        txtPONO.SetFocus
        Exit Sub
    End If


    If IsDate(txtDatePO) = False Then
        MessagePop RecSaveError, "Invalid Date", "Date Of PO is Required Field"
        On Error Resume Next
        txtDatePO.SetFocus
        Exit Sub
    End If

    If NumericVal(txtPy_CD_Amount) = 0 Then
        If MsgBox(" Zero Amount ! Are You Sure ?", vbQuestion + vbYesNo) = vbNo Then
            On Error Resume Next
            txtPy_CD_Amount.SetFocus
            Exit Sub
        End If

    End If

    '''''''AXP063020071200
    Dim lng                                                           As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_PO  WHERE PO_NO=" & N2Str2Null(txtPONO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "PO Number Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsPO!PO_NO)) <> UCase(txtPONO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "PO Number Already Exist"
            Exit Sub
        End If
    End If

    If IsDate(txtDueDate) = False Then
        MessagePop RecSaveError, "Invalid Date", "Date Required is Required Field"
        On Error Resume Next
        txtDueDate.SetFocus
        Exit Sub
    End If

    If Null2String(txtModelCode) = "" Then
        MessagePop RecSaveError, "Invalid Code", "Code is Required Field"
        On Error Resume Next
        txtModelCode.SetFocus
        Exit Sub
    End If

    '''''''AXP063020071200

    lng = gconDMIS.Execute("select Count(*) from SMIS_PO WHERE PO_NO=" & N2Str2Null(txtPONO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Purchase Order Number Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsPO!PO_NO)) <> UCase(txtPONO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Sales Order Already Exist"
            Exit Sub
        End If
    End If



    If Null2String(cboModelDescript) = "" Then
        MessagePop RecSaveError, "Invalid Model Description", "Description is Required Field"
        On Error Resume Next
        cboModelDescript.SetFocus
        Exit Sub
    End If

    If Null2String(cboColor) = "" Then
        If MsgBox(" Color Information Missing Are You Sure?", vbYesNo + vbQuestion) = vbNo Then
        On Error Resume Next
        cboColor.SetFocus
            Exit Sub
        End If
    End If

    If Null2String(txtSource) = "" Then
        On Error Resume Next
        txtSource.SetFocus
        Exit Sub
    End If

    Dim TEMPRS                                                        As ADODB.Recordset

    If AddorEdit = "ADD" Then

        sql = " INSERT INTO SMIS_PO "
        sql = sql & " ( PO_NO, DateOrdered, ModelDescript"
        sql = sql & " , FinLcIssuingBank , LCNo "
        sql = sql & " , CD_BankName , CD_CheckNo, CD_Date,CD_Amount "
        sql = sql & " , Model, ModelYear, ModelCode,CUSCDE,DateReq, Source, Color, Fuel, Notes, subsidy) "
        sql = sql & " VALUES( "
        sql = sql & N2Str2Null(txtPONO) & "," & N2Str2Null(txtDatePO) & " ," & N2Str2Null(cboModelDescript) & ", "
        sql = sql & N2Str2Null(cboPy_FinLcIssuingBank) & "," & N2Str2Null(txtPy_LCNo) & " ,"
        sql = sql & N2Str2Null(cboPy_CD_BankName) & "," & N2Str2Null(txtPy_CD_CheckNo) & " ," & N2Str2Null(txPy_CD_Date) & "," & NumericVal(txtPy_CD_Amount) & ","
        sql = sql & N2Str2Null(txtModel) & "," & N2Str2Null(txtModelYear) & " ," & N2Str2Null(txtModelCode) & ", "
        sql = sql & N2Str2Null(txtCusCode) & "," & N2Date2Null(txtDueDate) & "," & vbCrLf
        sql = sql & N2Str2Null(txtSource) & "," & N2Str2Null(cboColor) & " ," & N2Str2Null(txtFuel) & ", " & N2Str2Null(txtNotes) & ", "
        sql = sql & NumericVal(txtSubsidy) & ")" & vbCrLf
        sql = sql & " SELECT @@IDENTITY "

    Else
        'FinLcIssuingBank
        'LCNo
        'CD_BankName
        'CD_CheckNo
        'CD_Date
        'CD_Amount

        sql = " Update SMIS_PO "
        sql = sql & " SET "
        sql = sql & " PO_NO=" & N2Str2Null(txtPONO) & ","
        sql = sql & " DateOrdered=" & N2Str2Null(txtDatePO) & ","
        sql = sql & " DateReq=" & N2Str2Null(txtDueDate) & ","
        sql = sql & " CUSCDE=" & N2Str2Null(txtCusCode) & ","
        sql = sql & " ModelDescript=" & N2Str2Null(cboModelDescript) & ","
        sql = sql & " Model=" & N2Str2Null(txtModel) & ","
        sql = sql & " ModelYear=" & N2Str2Null(txtModelYear) & ","
        sql = sql & " Source=" & N2Str2Null(txtSource) & ","
        sql = sql & " ModelCode=" & N2Str2Null(txtModelCode) & ","
        sql = sql & " Color=" & N2Str2Null(cboColor) & ","
        sql = sql & " Fuel=" & N2Str2Null(txtFuel) & ","
        sql = sql & " Notes=" & N2Str2Null(txtNotes) & ","
        sql = sql & " FinLcIssuingBank=" & N2Str2Null(cboPy_FinLcIssuingBank) & ","
        sql = sql & " LCNo=" & N2Str2Null(txtPy_LCNo) & ","
        sql = sql & " CD_BankName=" & N2Str2Null(cboPy_CD_BankName) & ","
        sql = sql & " CD_CheckNo=" & N2Str2Null(txtPy_CD_CheckNo) & ","
        sql = sql & " CD_Date=" & N2Date2Null(txPy_CD_Date) & ","
        sql = sql & " CD_Amount=" & NumericVal(txtPy_CD_Amount) & ","
        sql = sql & " SUBSIDY=" & NumericVal(txtSubsidy) & ","
        sql = sql & " ModeOfPayment=" & N2Str2Null(SetModeOfPayment(cboModeOfPayment))
        sql = sql & " WHERE ID=" & N2Str2Null(txtID)
    End If

    '    If LTrim(RTrim(txtCusCode)) <> "" Then
    '        gconDMIS.Execute "update SMIS_MRRINV SET CUSTOMERCODE='" & txtCusCode & "' where PONO='" & txtPONO & "'"
    '        gconDMIS.Execute "update SMIS_MRRINV SET ISTATUS='A' WHERE PONO='" & txtPONO & "' AND (ISTATUS='O')  "
    '
    '    End If

    'If AddorEdit = "ADD" And txtID = 0 Then
    '    AddDetails
    'End If

    Set TEMPRS = gconDMIS.Execute(sql)

    Set TEMPRS = TEMPRS.NextRecordset
    If Not TEMPRS Is Nothing Then
        txtID = TEMPRS.Collect(0)
    End If

    picAdds.Visible = True
    picSaves.Visible = False

    rsRefresh
    rsPO.Find ("ID=" & txtID)
    CboRefresh
    cmdCancel.Value = True

    FillSearchGrid



    Exit Sub
Errorcode:
    ShowVBError

End Sub
Sub SearchID(xxx)
    If Not (rsPO.EOF Or rsPO.BOF) Then
        rsPO.Find ("ID=" & xxx)
        StoreMemVars

    End If

End Sub

Sub FillSearchGrid()
    Dim TEMPRS                                                        As ADODB.Recordset
    lstPO.Enabled = False
    Dim xxx                                                           As String
    If optVModel.Value = True Then
        
        Set TEMPRS = gconDMIS.Execute("SELECT DateOrdered, PO_NO , ID FROM SMIS_PO WHERE ModelDescript Like " & N2Str2Null(ReplaceQuote(textSearch & "%")))
        
    ElseIf optDate.Value = True Then
        
        Set TEMPRS = gconDMIS.Execute("SELECT DateOrdered, PO_NO , ID FROM SMIS_PO WHERE  DateOrdered Like " & N2Str2Null(ReplaceQuote(textSearch & "%")))
        
    ElseIf optPO.Value = True Then
        
        Set TEMPRS = gconDMIS.Execute("SELECT DateOrdered, PO_NO ,  ID FROM SMIS_PO WHERE  PO_NO Like '%" & Repleys(textSearch) & "%'")
    End If


    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        flex_FillListView TEMPRS, lstPO

        'Listview_Loadval lstPO.ListItems, Temprs
        lstPO.Enabled = True
    End If

    Set TEMPRS = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    picAdds.Visible = True
    picSaves.Visible = False
    picTop.Enabled = False
    fraCrNo.Enabled = False
    fraCheckDetail.Enabled = False

    InitCombo
    CboRefresh
    Call AddColumnHeader("Date, PO", lstPO)
    Call ResizeColumnHeader(lstPO, "40,50")
    Set SearchMaster = New frmSMIS_Mis_SearchMaster
    SetCompany
    rsRefresh
    initMemvars
    StoreMemVars

End Sub


Sub InitCombo()
    Dim sql                                                           As String

    

    Set rsS_Model = New ADODB.Recordset
    Call rsS_Model.Open("Select descript from All_Model where LEN(code)<> 0 order by descript asc", gconDMIS, adOpenKeyset)
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboModelDescript.Clear
        Do While Not rsS_Model.EOF
            cboModelDescript.AddItem UCase(Null2String(rsS_Model!DESCRIPT))
            rsS_Model.MoveNext
        Loop
    End If
    Set rsColor = New ADODB.Recordset
    Call rsColor.Open("Select DISTINCT(Color_Desc) as Color_Desc from All_Color  order by 1 asc", gconDMIS, adOpenKeyset)
    If Not rsColor.EOF And Not rsColor.BOF Then
        rsColor.MoveFirst
        cboColor.Clear
        Do While Not rsColor.EOF
            cboColor.AddItem UCase(Null2String(rsColor!COLOR_DESC))
            rsColor.MoveNext
        Loop
    End If
End Sub
Sub CboRefresh()
    Dim rsTemp                                                        As ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("Select DISTINCT(FinLcIssuingBank) from SMIS_PO WHERE LEN(FinLcIssuingBank)>0 Order by 1 asc")

    Combo_Loadval cboPy_FinLcIssuingBank, rsTemp

    Set rsTemp = gconDMIS.Execute("Select DISTINCT(CD_BankName) from SMIS_PO  WHERE LEN(CD_BankName)>0 Order by 1 asc")

    Combo_Loadval cboPy_CD_BankName, rsTemp

End Sub


Sub initMemvars()
    Dim cntrl                                                         As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            cntrl.Text = vbNullString
        End If
    Next
    txtPODetailID = 0
    txtID = 0
    lblStatus = ""
    txtQty = 1
    txtPy_CD_Amount = "0.00"
    txtCompanyName = CompName
    txtReqBy = ReqBy
    txtNotedBy = NotedBy
End Sub

Function ItemExists(StringToFind As String, ColumnToLook As Integer) As Integer
    For I = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(I, ColumnToLook) = StringToFind Then
            ItemExists = ItemExists + 1
            Exit For
        End If
    Next
End Function

Private Sub lstPO_DblClick()
    If lstPO.SelectedItem Is Nothing Then: Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lstPO_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ADDER:
    rsPO.MoveFirst
    rsPO.Find ("ID=" & Item.ListSubItems(2).Text)
    StoreMemVars
    Exit Sub
ADDER:
    Err.Clear
End Sub

Private Sub optDate_Click()
    textSearch_Change
End Sub

Private Sub optPO_Click()
    textSearch_Change
End Sub

Private Sub optVModel_Click()
    textSearch_Change
End Sub

Sub rsRefresh()
    Set rsPO = New ADODB.Recordset
    Call rsPO.Open("SELECT  * FROM SMIS_PO order by id DESC ", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Function SelectCombo(C As ComboBox, str As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim I                                                             As Long
    Dim ItemDataX                                                     As Long
    If ByItemData = False Then
        For I = 0 To C.ListCount - 1
            If UCase(C.List(I)) = UCase(Trim(str)) Then
                SelectCombo = I
                Exit Function
            End If
        Next
    Else
        If str = vbNullString Then
            SelectCombo = -1
            Exit Function
        End If

        ItemDataX = CLng(str)

        For I = 0 To C.ListCount - 1
            If C.ItemData(I) = str Then
                SelectCombo = I
                Exit Function
            End If
        Next
    End If
    SelectCombo = -1
End Function

Sub SetCompany()
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select CompanyName, CompanyAddress,PreparedBy,ApprovedBy from ALL_PRofile Where ModuleName='PMIS'")
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        CompName = Null2String(TEMPRS!CompanyName)
        NotedBy = Null2String(TEMPRS!PreparedBy)
        ReqBy = Null2String(TEMPRS!ApprovedBy)
    End If

End Sub

Sub SetModelLine(xxx As String, ByCode As Boolean)
    Dim TEMPRS                                                        As ADODB.Recordset
    If ByCode = True Then
        Set TEMPRS = gconDMIS.Execute("Select DESCRIPT, MODEL, CODE from ALL_MODEL WHERE CODE=" & N2Str2Null(ReplaceQuote(xxx)))
    Else
        Set TEMPRS = gconDMIS.Execute("Select DESCRIPT, MODEL , CODE from ALL_MODEL WHERE DESCRIPT=" & N2Str2Null(ReplaceQuote(xxx)))
    End If
    DontChange = True

    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        If ByCode = True Then
            cboModelDescript.Text = Null2String(TEMPRS!DESCRIPT)
        Else
            txtModelCode = Null2String(TEMPRS!CODE)
        End If
        txtModel = Null2String(TEMPRS!Model)
        cmdSave.Enabled = True
    Else
        MsgBox " Invalid Model Code ! " & vbCrLf & "Try Again Or Select From Drop Down of Model Description", vbInformation
        If ByCode = True Then
            On Error Resume Next
            txtModelCode.SetFocus
        Else
            On Error Resume Next
            cboModelDescript.SetFocus
        End If
        cmdSave.Enabled = False
    End If
End Sub

Sub SetPartsLines(PartIDNo As Variant, ForCombo As Boolean)
    Dim TEMPRS                                                        As ADODB.Recordset
    If ForCombo = False Then
        Set TEMPRS = gconDMIS.Execute("SELECT  SRP, STOCKNO , STOCKDESC FROM PMIS_STOCKMAS WHERE ID=" & PartIDNo)
        If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
            txtUnitAmount = FormatNumber(NumericVal(TEMPRS!SRP))
            txtPartNo = Null2String(TEMPRS!STOCKNO)
            txtPART_NUMBER = Null2String(TEMPRS!STOCKNO)
            txtPartDescription = Null2String(TEMPRS!STOCKDESC)
        End If
    Else
        Set TEMPRS = gconDMIS.Execute("SELECT  SRP,  STOCKDESC , STOCKNO FROM PMIS_STOCKMAS WHERE STOCKNO=" & N2Str2Null(PartIDNo))
        If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
            txtUnitAmount = FormatNumber(NumericVal(TEMPRS!SRP))
            cboPART_NAME.Text = Null2String(TEMPRS!STOCKDESC)
            txtPART_NUMBER = Null2String(TEMPRS!STOCKNO)
            txtPartDescription = Null2String(TEMPRS!STOCKDESC)
        End If
    End If
End Sub

Function SetStatus(XString) As String
'when 'FO' then 'FOR ORDERING'
'when 'BO' then 'BACK ORDER STAGE'
'when 'AS' then 'ALLOCATION STAGE'
'when 'KS' then 'PICKING STAGE'
'when 'PS' then 'PACKING STAGE'
'when 'SS' then 'SHIPPING STAGE'

    Select Case XString
        Case "FO"
            SetStatus = "FOR ORDERING"
        Case "BO"
            SetStatus = "BACK ORDER STAGE"
        Case "AS"
            SetStatus = "ALLOCATION STAGE"
        Case "KS"
            SetStatus = "PICKING STAGE"
        Case "PS"
            SetStatus = "PACKING STAGE"
        Case "SS"
            SetStatus = "SHIPPING STAGE"
    End Select
End Function

Sub StoreMemVars()

    If Not (rsPO.EOF Or rsPO.BOF) Then
        cmdEdit.Enabled = True
        'ID, PO_NO, DateOrdered, Descript, Model, YearModel, Source, Color, Fuel, DatePullOut, DateReleased, DateInvoiced, CustomerCode, Status, Notes FROM
        txtSource = Null2String(rsPO!Source)
        txtPONO = Null2String(rsPO!PO_NO)
        txtDatePO = Null2String(rsPO!DateOrdered)
        cboModelDescript = Null2String(rsPO!ModelDescript)
        txtModel = Null2String(rsPO!Model)
        txtModelCode = Null2String(rsPO!ModelCode)
        txtModelYear = Null2String(rsPO!MODELYEAR)
        cboColor = Null2String(rsPO!Color)
        txtNotes = Null2String(rsPO!Notes)
        txtFuel = Null2String(rsPO!Fuel)
        txtID = Null2String(rsPO!ID)
        txtDueDate = Null2String(rsPO!DATEREQ)
        txtCusCode = Null2String(rsPO!CUSCDE)
        cboPy_FinLcIssuingBank = Null2String(rsPO!FinLcIssuingBank)
        txtPy_LCNo = Null2String(rsPO!LCNo)
        cboPy_CD_BankName = Null2String(rsPO!CD_BankName)
        txtPy_CD_CheckNo = Null2String(rsPO!CD_CheckNo)
        txtPY_CD_Date = Null2Date(rsPO!CD_Date)
        txtSubsidy = FormatNumber(NumericVal(rsPO!SUBSIDY))
        txtPy_CD_Amount = FormatNumber(NumericVal(rsPO!CD_AMOUNT))

        cboModeOfPayment = GetModeOfPayment(Null2String(rsPO!ModeOfPayment))

        If IsDate(rsPO!DateReceived) = True Then
            lblStatus = "***RECEIVED***"
            cmdCancelCO.Enabled = False
            cmdUnPost.Enabled = False
            cmdPost.Enabled = False
            cmdEdit.Enabled = False
        Else
            If Null2String(rsPO!status) = "C" Then
                cmdCancelCO.Enabled = False
                cmdUnPost.Enabled = False
                cmdPost.Enabled = False
                lblStatus = "***Cancelled***"
                cmdEdit.Enabled = False
            ElseIf Null2String(rsPO!status) = "P" Then
                cmdCancelCO.Enabled = False
                cmdUnPost.Enabled = True
                cmdPost.Enabled = False
                lblStatus = "***Posted ***"
                cmdEdit.Enabled = False
            Else
                cmdCancelCO.Enabled = True
                cmdUnPost.Enabled = False
                cmdPost.Enabled = True
                lblStatus = ""
                cmdEdit.Enabled = True
            End If
        End If
    Else

        ShowNoRecord
        cmdAdd.Value = True
    End If

End Sub

Private Sub SearchMaster_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    txtCusCode = oCusRs!CUSCDE
    Unload SearchMaster
End Sub

Private Sub txtDueDate_GotFocus()
    If IsDate(txtDueDate) = False Then
        txtDueDate = ""
    Else
        txtDueDate = Format(txtDueDate, "mm/dd/yyyy")
    End If

End Sub

Private Sub txtDueDate_LostFocus()
    If IsDate(txtDueDate) = False Then
        txtDueDate = ""
    Else
        txtDueDate = Format(txtDueDate, "mmm dd yyyy")
    End If
End Sub



Private Sub textSearch_Change()
    FillSearchGrid
End Sub

Private Sub Timer1_Timer()
    If lblStatus.Caption <> "" Then
        If lblStatus.Visible = True Then
            lblStatus.Visible = False
        Else
            lblStatus.Visible = True
        End If
    End If
End Sub

Private Sub txtACC_QTY_Change()
    txtACC_Total = NumericVal(txtACC_QTY) * NumericVal(txtACC_SRP)
End Sub

Private Sub txtACC_QTY_GotFocus()
    If NumericVal(txtACC_QTY.Text) <= 0 Then txtACC_QTY = ""
End Sub

Private Sub txtACC_QTY_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtACC_QTY_LostFocus()
    If NumericVal(txtACC_QTY.Text) <= 0 Then txtACC_QTY = "0"

End Sub

Private Sub txtACC_SRP_Change()
    txtACC_Total = NumericVal(txtACC_QTY) * NumericVal(txtACC_SRP)
End Sub

Private Sub txtACC_SRP_GotFocus()
    If NumericVal(txtACC_SRP.Text) <= 0 Then txtACC_SRP = ""

End Sub

Private Sub txtACC_SRP_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtACC_SRP_LostFocus()
    If NumericVal(txtACC_SRP.Text) <= 0 Then txtACC_SRP = "0.00"
    txtACC_SRP = FormatNumber(txtACC_SRP)
End Sub

Private Sub txtACC_Total_GotFocus()
    If NumericVal(txtACC_Total.Text) <= 0 Then txtACC_Total = ""

End Sub

Private Sub txtACC_Total_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtACC_Total_LostFocus()
    If NumericVal(txtACC_Total.Text) <= 0 Then txtACC_Total = "0.00"
    txtACC_Total = FormatNumber(txtACC_Total)
End Sub

Private Sub txtDatePO_GotFocus()
    If IsDate(txtDatePO) = False Then
        txtDatePO = ""
    Else
        txtDatePO = Format(txtDatePO, "mm/dd/yyyy")
    End If
End Sub

Private Sub txtDatePO_LostFocus()
    If IsDate(txtDatePO) = False Then
        txtDatePO = ""
    Else
        txtDatePO = Format(txtDatePO, "mmm dd yyyy")
    End If
End Sub

Private Sub txtFuel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)

End Sub

Private Sub txtModelCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If txtModelCode <> "" And KeyCode = 13 Then
        SetModelLine txtModelCode, True
    ElseIf KeyCode = vbKeyEscape And AddorEdit = "EDIT" Then
        DontChange = True
        txtModelCode = Null2String(rsPO!ModelCode)
        txtModel = Null2String(rsPO!Model)
        txtModelCode = Null2String(rsPO!ModelCode)

    End If
End Sub

Private Sub txtModelCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtModelCode_LostFocus()

    DontChange = False
End Sub

Private Sub txtModelYear_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub


Private Sub txtPONO_LostFocus()
    txtPONO = Format(txtPONO, "000000")
End Sub

Private Sub txtPy_CD_Amount_GotFocus()
    If NumericVal(txtPy_CD_Amount.Text) <= 0 Then txtPy_CD_Amount = ""
End Sub
Private Sub txtPy_CD_Amount_LostFocus()
    If NumericVal(txtPy_CD_Amount) <= 0 Then txtPy_CD_Amount = "0.00"
    txtPy_CD_Amount = FormatNumber(NumericVal(txtPy_CD_Amount))
End Sub
Private Sub txtSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub
Function SetModeOfPayment(xxx)
    xxx = UCase(xxx)
    If xxx = UCase("Letter of Credit") Then
        SetModeOfPayment = "LC"
    ElseIf xxx = UCase("Open Account") Then
        SetModeOfPayment = "OA"
    ElseIf xxx = UCase("Promissory Note") Then
        SetModeOfPayment = "PN"
    ElseIf xxx = UCase("Financing Co.") Then
        SetModeOfPayment = "FC"
    ElseIf xxx = UCase("Cash") Then
        SetModeOfPayment = "CA"
    End If
End Function
Function GetModeOfPayment(xxx)
    If xxx = "LC" Then
        GetModeOfPayment = "Letter of Credit"
    ElseIf xxx = "CA" Then
        GetModeOfPayment = "Cash"
    ElseIf xxx = "OA" Then
        GetModeOfPayment = "Open Account"
    ElseIf xxx = "PN" Then
        GetModeOfPayment = "Promissory Note"
    ElseIf xxx = "FC" Then
        GetModeOfPayment = "Financing Co."
    End If
End Function

Private Sub txtSubsidy_GotFocus()
    If NumericVal(txtSubsidy.Text) <= 0 Then txtSubsidy = ""
End Sub
Private Sub txtSubsidy_LostFocus()
    If NumericVal(txtSubsidy) <= 0 Then txtSubsidy = "0.00"
    txtSubsidy = FormatNumber(NumericVal(txtSubsidy))
End Sub

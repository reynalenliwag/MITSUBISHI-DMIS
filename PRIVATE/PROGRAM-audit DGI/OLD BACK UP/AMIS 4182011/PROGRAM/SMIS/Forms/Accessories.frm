VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_Inventory_AccessoriesEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accessories Master File"
   ClientHeight    =   6450
   ClientLeft      =   900
   ClientTop       =   315
   ClientWidth     =   8925
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Accessories.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   8925
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3120
      ScaleHeight     =   855
      ScaleWidth      =   5715
      TabIndex        =   55
      Top             =   5460
      Width           =   5715
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
         Left            =   4980
         MouseIcon       =   "Accessories.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   4980
         MouseIcon       =   "Accessories.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Print this Record"
         Top             =   30
         Visible         =   0   'False
         Width           =   705
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
         Height          =   795
         Left            =   4290
         MouseIcon       =   "Accessories.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
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
         Left            =   3600
         MouseIcon       =   "Accessories.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Left            =   2910
         MouseIcon       =   "Accessories.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Add Record"
         Top             =   30
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
         Left            =   2220
         MouseIcon       =   "Accessories.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Find a Record"
         Top             =   30
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
         Left            =   1530
         MouseIcon       =   "Accessories.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Move to Next Record"
         Top             =   30
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
         Left            =   840
         MouseIcon       =   "Accessories.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5355
      Left            =   2730
      TabIndex        =   22
      Top             =   30
      Width           =   6105
      Begin VB.TextBox txtMAC 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2190
         Width           =   1755
      End
      Begin MSMask.MaskEdBox txtWFP 
         Height          =   345
         Left            =   4230
         TabIndex        =   9
         Top             =   2580
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtOldNo 
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
         Height          =   345
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   8
         ToolTipText     =   "Type the part's old number, if there's any."
         Top             =   2580
         Width           =   1755
      End
      Begin VB.Frame fraSupervisor 
         BorderStyle     =   0  'None
         Height          =   2355
         Left            =   60
         TabIndex        =   34
         Top             =   2910
         Width           =   5925
         Begin VB.TextBox txtResService 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   4170
            MaxLength       =   10
            TabIndex        =   21
            ToolTipText     =   "Type the part's res. service (e.g. 80, 40)"
            Top             =   2010
            Width           =   1725
         End
         Begin VB.TextBox txtIssuances 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   4170
            MaxLength       =   10
            TabIndex        =   19
            ToolTipText     =   "Type the part's number of issuances (e.g. 5, 6, 0)"
            Top             =   1620
            Width           =   1725
         End
         Begin VB.TextBox txtAdjPhyCnt 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   4170
            MaxLength       =   10
            TabIndex        =   17
            ToolTipText     =   "Type the part's adjustment count (e.g. 58, 2, 3)"
            Top             =   1230
            Width           =   1725
         End
         Begin VB.TextBox txtTISSQty 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   4170
            MaxLength       =   10
            TabIndex        =   15
            ToolTipText     =   "Type the temporary ISS of the part (e.g. 55, 25)"
            Top             =   840
            Width           =   1725
         End
         Begin VB.TextBox txtLastM_Oh 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Height          =   345
            Left            =   4170
            MaxLength       =   10
            TabIndex        =   13
            ToolTipText     =   "Type the number of the last on-hand of the particular part (e.g. 2, 50, 65)"
            Top             =   450
            Width           =   1725
         End
         Begin VB.TextBox txtSRP 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   4170
            MaxLength       =   10
            TabIndex        =   11
            ToolTipText     =   "Type the Selling Retail Price of the part. Do not use peso symbol and comma as separator  (e.g. 265, 9500)"
            Top             =   60
            Width           =   1725
         End
         Begin VB.TextBox txtSStock 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   20
            ToolTipText     =   "Input the number of the part in the safety stock (e.g. 1, 5, 3) "
            Top             =   2010
            Width           =   1185
         End
         Begin VB.TextBox txtReceipts 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   18
            ToolTipText     =   "Type the number of receipts of the particular part, if there's any."
            Top             =   1620
            Width           =   1185
         End
         Begin VB.TextBox txtPhyCount 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   16
            ToolTipText     =   "Type the physical count of the particular part (e.g. 58, 60)"
            Top             =   1230
            Width           =   1185
         End
         Begin VB.TextBox txtTPOQty 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   14
            ToolTipText     =   "Type part's temporary PO."
            Top             =   840
            Width           =   1185
         End
         Begin VB.TextBox txtOnHand 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Height          =   345
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   12
            ToolTipText     =   "Type how many of that part are on hand (e.g. 5, 10)"
            Top             =   450
            Width           =   1185
         End
         Begin VB.TextBox txtLastM_Mac 
            Alignment       =   1  'Right Justify
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
            Height          =   345
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   10
            Top             =   60
            Width           =   1185
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
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
            Height          =   255
            Left            =   3390
            TabIndex        =   46
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Last MAC "
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
            Height          =   255
            Left            =   30
            TabIndex        =   45
            Top             =   90
            Width           =   1275
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Last On-Hand"
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
            Height          =   255
            Left            =   2820
            TabIndex        =   44
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "On Hand"
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
            Height          =   255
            Left            =   30
            TabIndex        =   43
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Temp. ISS"
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
            Height          =   255
            Left            =   2820
            TabIndex        =   42
            Top             =   870
            Width           =   1305
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Temp. PO"
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
            Height          =   255
            Left            =   30
            TabIndex        =   41
            Top             =   870
            Width           =   1275
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Phy Count"
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
            Height          =   255
            Left            =   30
            TabIndex        =   40
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Adj. Count"
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
            Height          =   255
            Left            =   2850
            TabIndex        =   39
            Top             =   1260
            Width           =   1305
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Receipts"
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
            Height          =   255
            Left            =   30
            TabIndex        =   38
            Top             =   1650
            Width           =   1275
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Issuances"
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
            Height          =   255
            Left            =   2820
            TabIndex        =   37
            Top             =   1650
            Width           =   1305
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Res. Service"
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
            Height          =   255
            Left            =   2820
            TabIndex        =   36
            Top             =   2040
            Width           =   1305
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Safety Stock"
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
            Height          =   255
            Left            =   30
            TabIndex        =   35
            Top             =   2040
            Width           =   1275
         End
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   0
         ToolTipText     =   "Type part number (e.g. 030202 G 504, 033581G55613)"
         Top             =   240
         Width           =   1995
      End
      Begin VB.TextBox txtLocation 
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
         Height          =   345
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Type the location where the part can be found (e.g. Q-1)"
         Top             =   1800
         Width           =   1755
      End
      Begin VB.TextBox txtModelCode 
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
         Height          =   345
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   3
         ToolTipText     =   "Type model code (e.g. FE 97, CK1 CK2 CK4)"
         Top             =   1410
         Width           =   4515
      End
      Begin VB.TextBox txtVehType 
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
         Height          =   345
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1020
         Width           =   4470
      End
      Begin VB.TextBox txtPartDesc 
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
         Height          =   345
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   1
         ToolTipText     =   "Type description of the part (e.g. WRANGLER DT 195, ANTENNA ROD)"
         Top             =   630
         Width           =   4515
      End
      Begin MSMask.MaskEdBox txtMAD 
         Height          =   345
         Left            =   4230
         TabIndex        =   7
         Top             =   2190
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport rptPrintParts 
         Left            =   5340
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "List of New Part Numbers"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSMask.MaskEdBox txtDNP 
         Height          =   345
         Left            =   4230
         TabIndex        =   5
         Top             =   1800
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   3480
         TabIndex        =   48
         Top             =   270
         Width           =   225
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "DNP"
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
         Height          =   255
         Left            =   3450
         TabIndex        =   47
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "WFP"
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
         Height          =   255
         Left            =   3450
         TabIndex        =   33
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Old No."
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
         Height          =   255
         Left            =   90
         TabIndex        =   32
         Top             =   2610
         Width           =   1275
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MAD"
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
         Height          =   255
         Left            =   3450
         TabIndex        =   31
         Top             =   2220
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MAC "
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
         Height          =   255
         Left            =   90
         TabIndex        =   30
         Top             =   2220
         Width           =   1275
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories Number"
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
         Height          =   255
         Left            =   90
         TabIndex        =   29
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   255
         Left            =   90
         TabIndex        =   26
         Top             =   1860
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   25
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Veh Type"
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
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   255
         Left            =   90
         TabIndex        =   23
         Top             =   690
         Width           =   1275
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6330
      Left            =   90
      TabIndex        =   49
      Top             =   30
      Width           =   2595
      Begin VB.OptionButton optPartNo 
         Caption         =   "Accesories # [Alt + &R]"
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
         Left            =   180
         TabIndex        =   53
         Top             =   510
         Value           =   -1  'True
         Width           =   2385
      End
      Begin VB.OptionButton optDescription 
         Caption         =   "Description [Alt + &E]"
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
         Left            =   180
         TabIndex        =   52
         Top             =   750
         Width           =   2385
      End
      Begin VB.TextBox textSearch 
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
         Left            =   60
         MaxLength       =   35
         TabIndex        =   50
         Top             =   1080
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstAccessories 
         Height          =   4770
         Left            =   30
         TabIndex        =   51
         Top             =   1470
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   8414
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
         MouseIcon       =   "Accessories.frx":2D71
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PART NO."
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label23 
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
         TabIndex        =   54
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7380
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   64
      Top             =   5460
      Width           =   1470
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
         Left            =   720
         MouseIcon       =   "Accessories.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Cancel"
         Top             =   30
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
         Left            =   30
         MouseIcon       =   "Accessories.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "Accessories.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Save Accessories"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   300
      TabIndex        =   28
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   3330
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmSMIS_Inventory_AccessoriesEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPartMas                                                         As ADODB.Recordset
Dim AddorEdit                                                         As String

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "ACCESSORIES") = False Then Exit Sub
    Screen.MousePointer = 11
    rptPrintParts.ReportTitle = "ALPHABETICAL LISTING OF PART NUMBERS"
    rptPrintParts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrintParts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPrintParts, SMIS_REPORT_PATH & "printaccesories.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

'Upating Code       : AXP-0707200712:36
Private Sub cmdADD_Click()

    If Function_Access(LOGID, "Acess_Add", "ACCESSORIES") = False Then Exit Sub
    On Error GoTo Errorcode:


    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    txtPartNo.Enabled = True
    '    txtPartNo.SetFocus





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    txtPartNo.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "ACCESSORIES") = False Then Exit Sub
    On Error GoTo Errorcode
    If Not rsPartMas.BOF Or Not rsPartMas.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from PMIS_STOCKMAS where id = " & labid.Caption
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

'Upating Code       : AXP-0707200712:36
Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "ACCESSORIES") = False Then Exit Sub
    On Error GoTo Errorcode:


    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtPartDesc.SetFocus





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsPartMas.MoveNext
    If rsPartMas.EOF Then
        rsPartMas.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsPartMas.MovePrevious
    If rsPartMas.BOF Then
        rsPartMas.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim rsfindDup                                                     As ADODB.Recordset

    Dim vtxtPARTNO, vtxtPARTDESC, VTXTVehType                         As String
    Dim vtxtModelCode, VTXTLocation                                   As String
    Dim VTXTMAC, VTXTDNP                                              As Double
    Dim VTXTMAD, VTXTOldNo, VTXTWFP                                   As String
    Dim VTXTSRP, VTXTLastM_Mac                                        As Double
    Dim VTXTLastM_Oh, VTXTOnHand, VTXTTISSQty                         As Integer
    Dim VTXTTPOQty, VTXTTPRQty, VTXTPhyCount                          As Integer
    Dim VTXTAdjPhyCnt, VTXTReceipts, VTXTIssuances                    As Integer
    Dim VTXTSStock, VTXTResService                                    As Integer

    If IsNull(txtPartNo.Text) = True Then
        MsgSpeechBox "Part Number must not be empty"
        On Error Resume Next
        txtPartNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select STOCKNO from PMIS_STOCKMAS where STOCKNO = '" & txtPartNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Part Number already exist!"
                On Error Resume Next
                txtPartNo.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtPartDesc.Text = "" Then
        ShowIsRequiredMsg "Description"
        On Error Resume Next
        txtPartDesc.SetFocus
        Exit Sub
    End If

    vtxtPARTNO = N2Str2Null(txtPartNo.Text)
    vtxtPARTDESC = N2Str2Null(txtPartDesc.Text)
    VTXTVehType = N2Str2Null(txtVehType.Text)
    vtxtModelCode = N2Str2Null(txtModelCode.Text)
    VTXTLocation = N2Str2Null(txtLocation.Text)
    VTXTDNP = NumericVal(txtDNP.Text)
    VTXTMAC = NumericVal(txtMAC.Text)
    VTXTMAD = NumericVal(txtMAD.Text)
    VTXTOldNo = N2Str2Null(txtOldNo.Text)
    VTXTWFP = NumericVal(txtWFP.Text)
    VTXTSRP = NumericVal(txtSRP.Text)
    VTXTLastM_Mac = NumericVal(txtLastM_Mac.Text)
    VTXTLastM_Oh = NumericVal(txtLastM_Oh.Text)
    VTXTOnHand = NumericVal(txtOnHand.Text)
    VTXTTISSQty = NumericVal(txtTISSQty.Text)
    VTXTTPOQty = NumericVal(txtTPOQty.Text)
    VTXTPhyCount = NumericVal(txtPhyCount.Text)
    VTXTAdjPhyCnt = NumericVal(txtAdjPhyCnt.Text)
    VTXTReceipts = NumericVal(txtReceipts.Text)
    VTXTIssuances = NumericVal(txtIssuances.Text)
    VTXTSStock = NumericVal(txtSStock.Text)
    VTXTResService = NumericVal(txtResService.Text)

    If AddorEdit = "ADD" Then
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            rsPartMas.MoveLast
            labid.Caption = NumericVal(rsPartMas!ID) + 1
        End If
        gconDMIS.Execute "Insert into PMIS_STOCKMAS" & _
                       " (Type , LastM_Mac,STOCKNO,STOCKDESC,vehtype,modelcode,location,oldno,wfp,srp,tissqty,tpoqty,phycount,adjphycnt,receipts,issuances,sstock,resservice,lastupdate,usercode)" & _
                       " values ('A', " & VTXTLastM_Mac & "," & vtxtPARTNO & "," & vtxtPARTDESC & ", " & VTXTVehType & ", " & _
                       " " & vtxtModelCode & ", " & VTXTLocation & _
                         ", " & VTXTOldNo & ", " & VTXTWFP & _
                         ", " & VTXTSRP & ", " & VTXTTISSQty & ", " & VTXTTPOQty & _
                         ", " & VTXTPhyCount & ", " & VTXTAdjPhyCnt & ", " & VTXTReceipts & ", " & VTXTIssuances & ", " & VTXTSStock & ", " & VTXTResService & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ")"
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " LastM_Mac = " & VTXTLastM_Mac & "," & _
                       " STOCKNO = " & vtxtPARTNO & "," & _
                       " STOCKDESC = " & vtxtPARTDESC & "," & _
                       " vehtype = " & VTXTVehType & "," & _
                       " modelcode = " & vtxtModelCode & "," & _
                       " location = " & VTXTLocation & "," & _
                       " oldno = " & VTXTOldNo & "," & _
                       " wfp = " & VTXTWFP & "," & _
                       " srp = " & VTXTSRP & "," & _
                       " tpoqty = " & VTXTTPOQty & "," & _
                       " phycount = " & VTXTPhyCount & "," & _
                       " adjphycnt = " & VTXTAdjPhyCnt & "," & _
                       " receipts = " & VTXTReceipts & "," & _
                       " issuances = " & VTXTIssuances & "," & _
                       " sstock = " & VTXTSStock & ", resservice = " & VTXTResService & ", " & _
                       " lastupdate = '" & LOGDATE & "'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & _
                       " where id = " & labid.Caption
        ShowSuccessFullyUpdated
    End If
    rsRefresh
    On Error Resume Next
    rsPartMas.Find "id =" & labid.Caption
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    cmdCancel.Value = True
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    rsRefresh
    Frame1.Enabled = False
    txtWFP.Enabled = False: txtDNP.Enabled = False: txtMAC.Enabled = False: txtLastM_Mac.Enabled = False
    If LOGLEVEL = "SUPERVISOR" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
        fraSupervisor.Enabled = True
        If LOGLEVEL = "AUTHOR" Then
            txtWFP.Enabled = True: txtDNP.Enabled = True: txtMAC.Enabled = True: txtLastM_Mac.Enabled = True
        End If
    Else
        fraSupervisor.Enabled = False
    End If
    textSearch.Text = "":                                   'Picture3.ZOrder 0
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
    txtPartNo.Text = ""
    txtPartDesc.Text = ""
    txtVehType.Text = ""
    txtModelCode.Text = ""
    txtLocation.Text = ""
    txtDNP.Text = ""
    txtMAC.Text = 0
    txtMAD.Text = 0
    txtOldNo.Text = ""
    txtWFP.Text = 0
    txtSRP.Text = 0
    txtLastM_Mac.Text = 0
    txtLastM_Oh.Text = 0
    txtOnHand.Text = 0
    txtTISSQty.Text = 0
    txtTPOQty.Text = 0
    txtPhyCount.Text = 0
    txtAdjPhyCnt.Text = 0
    txtReceipts.Text = 0
    txtIssuances.Text = 0
    txtSStock.Text = 0
    txtResService.Text = 0
End Sub

Sub StoreMemVars()
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        labid.Caption = rsPartMas!ID
        txtPartNo.Text = Null2String(rsPartMas!STOCKNO)
        txtPartDesc.Text = Null2String(rsPartMas!STOCKDESC)
        txtVehType.Text = Null2String(rsPartMas!vehtype)
        txtModelCode.Text = Null2String(rsPartMas!ModelCode)
        txtLocation.Text = Null2String(rsPartMas!Location)
        If frmMain.wizEnc1.DecryptAccess("524956t^yw9|kk") = LOGLEVEL Then
            txtWFP.Text = "Classified"
            txtDNP.Text = "Classified"
            txtMAC.Text = "Classified"
            txtLastM_Mac.Text = "Classified"
        Else
            txtWFP.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!WFP))
            txtDNP.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
            txtMAC.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            txtLastM_Mac.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!lastm_mac))
        End If
        txtMAD.Text = N2Str2IntZero(rsPartMas!mad)
        txtOldNo.Text = Null2String(rsPartMas!oldno)
        txtSRP.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
        txtLastM_Oh.Text = N2Str2IntZero(rsPartMas!lastm_oh)
        txtOnHand.Text = N2Str2IntZero(rsPartMas!ONHAND)
        txtTISSQty.Text = N2Str2IntZero(rsPartMas!TISSQTY)
        txtTPOQty.Text = N2Str2IntZero(rsPartMas!tpoqty)
        txtPhyCount.Text = N2Str2IntZero(rsPartMas!PHYCOUNT)
        txtAdjPhyCnt.Text = N2Str2IntZero(rsPartMas!ADJPHYCNT)
        txtReceipts.Text = N2Str2IntZero(rsPartMas!receipts)
        txtIssuances.Text = N2Str2IntZero(rsPartMas!issuances)
        txtSStock.Text = N2Str2IntZero(rsPartMas!SSTOCK)
        txtResService.Text = N2Str2IntZero(rsPartMas!RESSERVICE)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select * from PMIS_STOCKMAS  WHERE TYPE='A' order by STOCKNO DESC", gconDMIS, adOpenForwardOnly
End Sub


Private Sub lstAccessories_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optPartNo.Value = True Then
        rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "ID", Item.ListSubItems(1).Text).Bookmark
    Else
        rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "ID", Item.ListSubItems(1).Text).Bookmark
    End If
    StoreMemVars
End Sub

Private Sub lstAccessories_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstAccessories
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

Private Sub lstAccessories_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstAccessories_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then On Error Resume Next: textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    FillSearchGrid (textSearch.Text)

End Sub


Sub FillSearchGrid(xxx As String)
    Dim rsParts                                                       As ADODB.Recordset
    Dim TEMPSQL                                                       As String
    lstAccessories.Enabled = False
    lstAccessories.Sorted = False: lstAccessories.ListItems.Clear
    lstAccessories.Enabled = False
    Set rsParts = New ADODB.Recordset

    If Trim(xxx) = "" Then
        TEMPSQL = "select TOP 50 STOCKNO, ID from PMIS_STOCKMAS where TYPE='A' ORDER BY STOCKNO ASC "

    Else

        If optPartNo.Value = True Then
            TEMPSQL = "select STOCKNO, ID from PMIS_STOCKMAS where TYPE='A' AND STOCKNO like'" & xxx & "%'"
        Else
            TEMPSQL = "select STOCKDESC, ID from PMIS_STOCKMAS where TYPE='A' AND STOCKDESC like'" & xxx & "%'"
        End If
    End If

    Set rsParts = gconDMIS.Execute(TEMPSQL)
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstAccessories.Enabled = True
        Listview_Loadval Me.lstAccessories.ListItems, rsParts
        lstAccessories.Refresh
        lstAccessories.Enabled = True
    Else
        lstAccessories.Enabled = False
    End If

End Sub


Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstAccessories.ListItems.Count > 0 And lstAccessories.Enabled = True Then: lstAccessories.SetFocus
    End If
End Sub

Private Sub optDescription_Click()
    lstAccessories.ColumnHeaders(1).Text = "DESCRIPTION"
    FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optPARTNO_Click()
    lstAccessories.ColumnHeaders(1).Text = "PART NO."
    FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

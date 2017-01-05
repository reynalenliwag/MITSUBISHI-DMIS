VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISMaster_Parts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parts Master File"
   ClientHeight    =   7380
   ClientLeft      =   900
   ClientTop       =   315
   ClientWidth     =   8925
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Parts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   8925
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7380
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   63
      Top             =   6090
      Visible         =   0   'False
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
         MouseIcon       =   "Parts.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   735
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
         Left            =   0
         MouseIcon       =   "Parts.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   3015
      ScaleHeight     =   2025
      ScaleWidth      =   6015
      TabIndex        =   54
      Top             =   6090
      Width           =   6015
      Begin VB.TextBox TXT_ACTIVE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   840
         Width           =   2625
      End
      Begin VB.CommandButton cmdActiveInactive 
         Caption         =   "Tag It Active"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2700
         TabIndex        =   72
         Top             =   840
         Width           =   1365
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
         Height          =   795
         Left            =   5100
         MouseIcon       =   "Parts.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Exit Window"
         Top             =   0
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
         Height          =   795
         Left            =   4380
         MouseIcon       =   "Parts.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Print this Record"
         Top             =   0
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
         Height          =   795
         Left            =   3660
         MouseIcon       =   "Parts.frx":1B6C
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":1CBE
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
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
         Left            =   2940
         MouseIcon       =   "Parts.frx":1FE9
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":213B
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
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
         Left            =   2220
         MouseIcon       =   "Parts.frx":2497
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":25E9
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
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
         Left            =   1500
         MouseIcon       =   "Parts.frx":28FC
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":2A4E
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   735
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
         Left            =   780
         MouseIcon       =   "Parts.frx":2D48
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":2E9A
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   735
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
         Left            =   60
         MouseIcon       =   "Parts.frx":31F2
         MousePointer    =   99  'Custom
         Picture         =   "Parts.frx":3344
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6075
      Left            =   2640
      TabIndex        =   22
      Top             =   -30
      Width           =   6255
      Begin VB.TextBox txtPartType 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
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
         Height          =   315
         Left            =   4890
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   67
         Text            =   "Text1"
         ToolTipText     =   "Type part number (e.g. 030202 G 504, 033581G55613)"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.CheckBox chkNonHARI 
         Caption         =   "Non HARI Parts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2820
         TabIndex        =   66
         Top             =   990
         Width           =   1965
      End
      Begin VB.TextBox txtMAC 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2220
         Width           =   1755
      End
      Begin MSMask.MaskEdBox txtWFP 
         Height          =   345
         Left            =   4230
         TabIndex        =   9
         Top             =   3000
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
         Text            =   "Text1"
         ToolTipText     =   "Type the part's old number, if there's any."
         Top             =   2610
         Width           =   1755
      End
      Begin VB.Frame fraSupervisor 
         BorderStyle     =   0  'None
         Height          =   2355
         Left            =   60
         TabIndex        =   34
         Top             =   3330
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
            Text            =   "Text1"
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
            Text            =   "Text1"
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
            Text            =   "Text1"
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
            Text            =   "Text1"
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
            Text            =   "Text1"
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
            Text            =   "Text1"
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   20
            Text            =   "Text1"
            ToolTipText     =   "Input the number of the part in the safety stock (e.g. 1, 5, 3) "
            Top             =   2010
            Width           =   1275
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "Text1"
            ToolTipText     =   "Type the number of receipts of the particular part, if there's any."
            Top             =   1620
            Width           =   1275
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   16
            Text            =   "Text1"
            ToolTipText     =   "Type the physical count of the particular part (e.g. 58, 60)"
            Top             =   1230
            Width           =   1275
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   14
            Text            =   "Text1"
            ToolTipText     =   "Type part's temporary PO."
            Top             =   840
            Width           =   1275
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   12
            Text            =   "Text1"
            ToolTipText     =   "Type how many of that part are on hand (e.g. 5, 10)"
            Top             =   450
            Width           =   1275
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SRP :"
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
            Height          =   225
            Left            =   3630
            TabIndex        =   46
            Top             =   120
            Width           =   465
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LM MAC  :"
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
            Left            =   -150
            TabIndex        =   45
            ToolTipText     =   "Displays Last Month Moving Average Cost"
            Top             =   120
            Width           =   1545
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Last On-Hand :"
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
            Left            =   2820
            TabIndex        =   44
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "On Hand :"
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
            TabIndex        =   43
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Temp. ISS :"
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
            Left            =   2820
            TabIndex        =   42
            Top             =   870
            Width           =   1305
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Temp. PO :"
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
            TabIndex        =   41
            Top             =   870
            Width           =   1275
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Phy Count :"
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
            TabIndex        =   40
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Adj. Count :"
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
            Left            =   2820
            TabIndex        =   39
            Top             =   1260
            Width           =   1305
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Receipts :"
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
            TabIndex        =   38
            Top             =   1650
            Width           =   1275
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Issuances :"
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
            Left            =   2820
            TabIndex        =   37
            Top             =   1650
            Width           =   1305
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Res. Service :"
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
            Left            =   2820
            TabIndex        =   36
            Top             =   2040
            Width           =   1305
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Safety Stock :"
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
         MaxLength       =   30
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "Type part number (e.g. 030202 G 504, 033581G55613)"
         Top             =   240
         Width           =   3165
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
         MaxLength       =   100
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Type the location where the part can be found (e.g. Q-1)"
         Top             =   1800
         Width           =   4515
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
         Text            =   "Text1"
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
         Text            =   "Text1"
         Top             =   1020
         Width           =   360
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
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Type description of the part (e.g. WRANGLER DT 195, ANTENNA ROD)"
         Top             =   630
         Width           =   4515
      End
      Begin MSMask.MaskEdBox txtMAD 
         Height          =   345
         Left            =   4230
         TabIndex        =   7
         Top             =   2610
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
         Left            =   2010
         Top             =   2940
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
         Top             =   2220
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   0
         Left            =   4650
         TabIndex        =   48
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   3810
         TabIndex        =   47
         Top             =   2250
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   3810
         TabIndex        =   33
         Top             =   3045
         Width           =   390
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   3810
         TabIndex        =   31
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "**Part Number"
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
         Height          =   255
         Left            =   90
         TabIndex        =   29
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Alignment       =   1  'Right Justify
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
         Alignment       =   1  'Right Justify
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
         Alignment       =   1  'Right Justify
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
      Height          =   7410
      Left            =   30
      TabIndex        =   49
      Top             =   -30
      Width           =   2580
      Begin VB.OptionButton optPartNo 
         Caption         =   "Pa&rt Number [Alt + R]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   53
         Top             =   210
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton optDescription 
         Caption         =   "D&escription [Alt + E]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   52
         Top             =   450
         Width           =   1995
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
         Height          =   375
         Left            =   30
         MaxLength       =   35
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   690
         Width           =   2505
      End
      Begin MSComctlLib.ListView lstParts 
         Height          =   5385
         Left            =   30
         TabIndex        =   51
         Top             =   1080
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   9499
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
         MouseIcon       =   "Parts.frx":36A3
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
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   30
         TabIndex        =   68
         Top             =   6390
         Width           =   2505
         Begin VB.OptionButton OPT_ACTIVE 
            Caption         =   "Show All Active Parts"
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
            Left            =   60
            TabIndex        =   71
            Top             =   120
            Width           =   2025
         End
         Begin VB.OptionButton OPT_INACTIVE 
            Caption         =   "Show All Inactive Parts"
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
            Left            =   60
            TabIndex        =   70
            Top             =   390
            Width           =   2025
         End
         Begin VB.OptionButton OPT_ALL 
            Caption         =   "Show All Parts"
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
            Left            =   60
            TabIndex        =   69
            Top             =   630
            Value           =   -1  'True
            Width           =   2025
         End
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
Attribute VB_Name = "frmPMISMaster_Parts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPARTMAS                                          As ADODB.Recordset
Dim AddorEdit                                          As String
Dim LOCAL_STOCKTYPE                                    As String
Dim LOCAL_ACCESS                                       As String

Sub SearchStock(XX As String, XTYPE As String)
    If Not (RSPARTMAS.EOF Or RSPARTMAS.BOF) Then
        'RSPARTMAS.Find ("STOCKNO=" & N2Str2Null(xx) & " AND TYPE=" & N2Str2Null(XTYPE))
        RSPARTMAS.Filter = "STOCKNO=" & N2Str2Null(XX) & " AND TYPE=" & N2Str2Null(XTYPE)
        StoreMemVars
    End If
End Sub
Sub SETSTOCKTYPE(XXX As String)
    LOCAL_STOCKTYPE = XXX
    If XXX = "P" Then
        LOCAL_ACCESS = "PARTS MASTER FILE"
    ElseIf XXX = "A" Then
        LOCAL_ACCESS = "ACCESSORIES MASTER FILE"
    Else
        LOCAL_ACCESS = "MATERIALS MASTER FILE"
    End If
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
    txtTissqty.Text = 0
    txtTPOQty.Text = 0
    txtPhyCount.Text = 0
    txtAdjPhyCnt.Text = 0
    txtReceipts.Text = 0
    txtIssuances.Text = 0
    txtSStock.Text = 0
    txtResService.Text = 0
    txtPartType.Text = ""
End Sub

Sub StoreMemVars()
    Dim RSPARTS_DETAIL                                 As ADODB.Recordset



    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        Set RSPARTS_DETAIL = gconDMIS.Execute("select * from PMIS_STOCKMAS where ID= " & RSPARTMAS!ID)
        If (RSPARTS_DETAIL.EOF Or RSPARTS_DETAIL.BOF) Then
            ShowNoRecord
            cmdAdd.Value = True
            Exit Sub
        End If
        labID.Caption = RSPARTS_DETAIL!ID
        If Null2String(RSPARTS_DETAIL!Active) = "Y" Then
            TXT_ACTIVE = "ACTIVE"
            TXT_ACTIVE.ForeColor = &H8000&
            cmdActiveInactive.Caption = "Tag it Inactive"
            cmdActiveInactive.Tag = "Y"
        Else
            TXT_ACTIVE = "IN ACTIVE"
            TXT_ACTIVE.ForeColor = &HFF&
            cmdActiveInactive.Caption = "Tag it Active"
            cmdActiveInactive.Tag = "N"
        End If


        txtPartNo.Text = Null2String(RSPARTS_DETAIL!STOCKNO)
        txtPartDesc.Text = Null2String(RSPARTS_DETAIL!STOCKDESC)
        txtVehType.Text = Null2String(RSPARTS_DETAIL!vehtype)
        txtModelCode.Text = Null2String(RSPARTS_DETAIL!MODELCODE)
        txtLocation.Text = Null2String(RSPARTS_DETAIL!Location)
        txtWFP.Text = ToDoubleNumber(NumericVal(RSPARTS_DETAIL!WFP))
        txtDNP.Text = ToDoubleNumber(NumericVal(RSPARTS_DETAIL!dnp))
        txtMAC.Text = ToDoubleNumber(NumericVal(RSPARTS_DETAIL!MAC))
        txtLastM_Mac.Text = ToDoubleNumber(NumericVal(RSPARTS_DETAIL!LASTM_MAC))

        txtMAD.Text = NumericVal(RSPARTS_DETAIL!mad)
        txtOldNo.Text = Null2String(RSPARTS_DETAIL!oldno)
        txtSRP.Text = ToDoubleNumber(NumericVal(RSPARTS_DETAIL!SRP))
        txtLastM_Oh.Text = N2Str2IntZero(RSPARTS_DETAIL!LASTM_OH)
        txtOnHand.Text = N2Str2IntZero(RSPARTS_DETAIL!ONHAND)
        txtTissqty.Text = N2Str2IntZero(RSPARTS_DETAIL!TISSQTY)
        txtTPOQty.Text = N2Str2IntZero(RSPARTS_DETAIL!tpoqty)
        txtPhyCount.Text = N2Str2IntZero(RSPARTS_DETAIL!PHYCOUNT)
        txtAdjPhyCnt.Text = N2Str2IntZero(RSPARTS_DETAIL!ADJPHYCNT)
        txtReceipts.Text = N2Str2IntZero(RSPARTS_DETAIL!RECEIPTS)
        txtIssuances.Text = N2Str2IntZero(RSPARTS_DETAIL!ISSUANCES)
        txtSStock.Text = N2Str2IntZero(RSPARTS_DETAIL!SSTOCK)
        txtResService.Text = N2Str2IntZero(RSPARTS_DETAIL!RESSERVICE)
        If Null2String(RSPARTS_DETAIL!NON_HARI) = "Y" Then
            chkNonHARI.Value = 1
        Else
            chkNonHARI.Value = 0
        End If
        If LTrim(RTrim(Null2String(RSPARTS_DETAIL!StockType))) = "GJ" Then
            txtPartType.Text = "GJ"
        ElseIf LTrim(RTrim(Null2String(RSPARTS_DETAIL!StockType))) = "BP" Then
            txtPartType.Text = "BP"
        Else
            txtPartType.Text = "Others"
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set RSPARTMAS = New ADODB.Recordset
    'If OPT_ACTIVE.Value = True Then
    '    RSPARTMAS.Open "select stockno,ID from PMIS_STOCKMAS where TYPE = '" & LOCAL_STOCKTYPE & "' AND ACTIVE='Y' order by ID DESC", gconDMIS, adOpenDynamic
    'ElseIf OPT_INACTIVE.Value = True Then
    '    RSPARTMAS.Open "select stockno,ID from PMIS_STOCKMAS where TYPE = '" & LOCAL_STOCKTYPE & "' AND ISNULL(ACTIVE,'N')='N' order by ID DESC", gconDMIS, adOpenDynamic
    'Else
    RSPARTMAS.Open "select type,stockno,ID from PMIS_STOCKMAS where TYPE = '" & LOCAL_STOCKTYPE & "' order by ID DESC", gconDMIS, adOpenDynamic
    'End If
End Sub



Sub FillGrid()
    Dim rsParts                                        As ADODB.Recordset
    Dim SEARCH_STRING                                  As String
    lstParts.Sorted = False
    If optPartNo.Value = True Then
        If LTrim(RTrim(textSearch.Text)) <> "" Then
            SEARCH_STRING = " AND STOCKNO LIKE " & N2Str2Null(textSearch & "%") & " ORDER BY STOCKNO ASC"
        Else
            SEARCH_STRING = " ORDER BY STOCKNO ASC"
        End If
    Else
        If LTrim(RTrim(textSearch.Text)) <> "" Then
            SEARCH_STRING = " AND STOCKDESC LIKE " & N2Str2Null(textSearch & "%") & " ORDER BY STOCKDESC ASC"
        Else
            SEARCH_STRING = " ORDER BY STOCKDESC ASC"
        End If
    End If

    If OPT_ACTIVE.Value = True Then
        Set rsParts = gconDMIS.Execute("SELECT TOP 50 STOCKNO, ID FROM PMIS_STOCKMAS WHERE ACTIVE='Y' AND TYPE='" & LOCAL_STOCKTYPE & "'" & SEARCH_STRING)
    ElseIf OPT_INACTIVE.Value = True Then
        Set rsParts = gconDMIS.Execute("SELECT TOP 50 STOCKNO, ID FROM PMIS_STOCKMAS WHERE ISNULL(ACTIVE,'N')='N' AND TYPE='" & LOCAL_STOCKTYPE & "'" & SEARCH_STRING)
    Else
        Set rsParts = gconDMIS.Execute("SELECT TOP 50 STOCKNO, ID FROM PMIS_STOCKMAS WHERE TYPE='" & LOCAL_STOCKTYPE & "'" & SEARCH_STRING)
    End If
    Listview_Loadval Me.lstParts.ListItems, rsParts
    lstParts.Sorted = True
End Sub




Private Sub cmdActiveInactive_Click()

    If cmdActiveInactive.Tag = "Y" Then
        Dim rsCheck                                    As ADODB.Recordset
        Dim rsCount As ADODB.Recordset
        
        
        Set rsCount = gconDMIS.Execute("SELECT isnull(sum(onhand),0)  as onhand  from pmis_stockmas where id=" & labID)

        If rsCount!ONHAND > 0 Then
            MsgBox "Part Number has onhand quantity! " & vbCrLf & "Cannot make this Stock Inactive.", vbInformation
            Exit Sub
        End If
        
        Set rsCheck = gconDMIS.Execute("SELECT COUNT(*) from PMIS_ALLDAYTRAN WHERE STOCK_ORD =" & N2Str2Null(txtPartNo))
        
        If rsCheck.Fields(0).Value > 0 Then
            MsgBox "Part Number has been used in daily transaction file! ", vbInformation
             
        End If


        If MsgBox("Are You sure you want to make this Stock# Inactive", vbInformation + vbYesNo) = vbNo Then Exit Sub

        SQL_STATEMENT = "UPDATE pmis_stockmas set active='N' where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCAL_ACCESS, SQL_STATEMENT, labID, "", "TAGGED INACTIVE STOCK CODE: " & txtPartNo, "", "")
    Else
        If MsgBox("Are You sure you want to make this Stock# Active", vbInformation + vbYesNo) = vbNo Then Exit Sub
        SQL_STATEMENT = "UPDATE pmis_stockmas set active='Y' where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCAL_ACCESS, SQL_STATEMENT, labID, "", "TAGGED ACTIVE STOCK CODE: " & txtPartNo, "", "")
    End If

    ShowSuccessFullyUpdated
    FillGrid
    rsRefresh
    RSPARTMAS.Find ("ID=" & labID)
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LOCAL_ACCESS) = False Then Exit Sub
    On Error GoTo Errorcode:
    Screen.MousePointer = 11
    rptPrintParts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrintParts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If LOCAL_STOCKTYPE = "P" Then
        rptPrintParts.ReportTitle = "ALPHABETICAL LISTING OF PART NUMBERS"
        PrintSQLReport rptPrintParts, PMIS_REPORT_PATH & "printparts.rpt", "", DMIS_REPORT_Connection, 1
    ElseIf LOCAL_STOCKTYPE = "A" Then
        rptPrintParts.ReportTitle = "ALPHABETICAL LISTING OF ACCESSORIES NUMBERS"
        PrintSQLReport rptPrintParts, PMIS_REPORT_PATH & "printAccessories.rpt", "", DMIS_REPORT_Connection, 1
    Else
        rptPrintParts.ReportTitle = "ALPHABETICAL LISTING OF MATERIAL NUMBERS"
        PrintSQLReport rptPrintParts, PMIS_REPORT_PATH & "PrintMaterials.rpt", "", DMIS_REPORT_Connection, 1
    End If

    Call NEW_LogAudit("V", LOCAL_ACCESS, "", "", "", "", "", "")
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    Screen.MousePointer = 0
    ShowVBError

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", LOCAL_ACCESS) = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    lstParts.Enabled = False
    txtPartNo.Enabled = True
    textSearch.Enabled = False
    optPartNo.Enabled = False
    optDescription.Enabled = False
    'txtPartNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    txtPartNo.Enabled = False
    lstParts.Enabled = True
    textSearch.Enabled = True
    optPartNo.Enabled = True
    optDescription.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", LOCAL_ACCESS) = False Then Exit Sub
    'MsgBox
    On Error GoTo Errorcode

    Dim rsCheck                                        As ADODB.Recordset
    Set rsCheck = gconDMIS.Execute("SELECT COUNT(*) from PMIS_ALLDAYTRAN WHERE STOCK_ORD =" & N2Str2Null(txtPartNo))

    If rsCheck.Fields(0).Value > 0 Then
        MsgBox "Part Number has been used in daily transaction! " & vbCrLf & "Cannot Delete The Record.", vbInformation
    Else

        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from PMIS_STOCKMAS where id = " & labID.Caption
            gconDMIS.Execute SQL_STATEMENT
            Call NEW_LogAudit("X", LOCAL_ACCESS, SQL_STATEMENT, labID, "", "CODE: " & labID, "", "")
            ShowDeletedMsg
            FillGrid
        End If
        rsRefresh
        StoreMemVars
    End If
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LOCAL_ACCESS) = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    On Error Resume Next
    txtPartDesc.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    RSPARTMAS.MoveNext
    If RSPARTMAS.EOF Then
        RSPARTMAS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RSPARTMAS.MovePrevious
    If RSPARTMAS.BOF Then
        RSPARTMAS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim RSFINDDUP                                      As ADODB.Recordset
    Dim vtxtPARTNO                                     As String
    Dim vtxtPARTDESC                                   As String
    Dim VTXTVEHTYPE                                    As String
    Dim VTXTMODELCODE                                  As String
    Dim VTXTLocation                                   As String
    Dim vtxtMAC                                        As Double
    Dim VTXTDNP                                        As Double
    Dim VTXTMAD                                        As String
    Dim VTXTOLDNO                                      As String
    Dim VTXTWFP                                        As String
    Dim VTXTSRP                                        As Double
    Dim VTXTLASTM_OH                                   As Long
    Dim VTXTOnHand                                     As Long
    Dim VTXTTISSQTY                                    As Long
    Dim VTXTTPOQTY                                     As Long
    Dim VTXTPHYCOUNT                                   As Long
    Dim VTXTADJPHYCNT                                  As Long
    Dim VTXTRECEIPTS                                   As Long
    Dim VTXTISSUANCES                                  As Long
    Dim VTXTSSTOCK                                     As Long
    Dim VTXTRESSERVICE                                 As Long
    Dim VTXTLastM_Mac                                  As Double

    If IsNull(txtPartNo.Text) = True Then
        MsgSpeechBox "Stock Number must not be empty"
        On Error Resume Next
        txtPartNo.SetFocus
        Exit Sub
    End If
    If txtPartDesc.Text = "" Then
        ShowIsRequiredMsg "Description"
        On Error Resume Next
        txtPartDesc.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Set RSFINDDUP = New ADODB.Recordset
        RSFINDDUP.Open "select STOCKNO from PMIS_STOCKMAS where TYPE = '" & LOCAL_STOCKTYPE & "' AND STOCKNO = '" & txtPartNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
            MsgSpeechBox "Part Number already exist!"
            On Error Resume Next
            txtPartNo.SetFocus
            Exit Sub
        End If
    Else
        If LTrim(RTrim(txtPartNo)) <> Null2String(RSPARTMAS!STOCKNO) Then
            Set RSFINDDUP = New ADODB.Recordset
            RSFINDDUP.Open "select STOCKNO from PMIS_STOCKMAS where STOCKNO = '" & Repleys(LTrim(RTrim(txtPartNo))) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgSpeechBox "Stock Number already exist!"
                On Error Resume Next
                txtPartNo.SetFocus
                Exit Sub
            End If
        End If
    End If


    vtxtPARTNO = N2Str2Null(LTrim(RTrim(txtPartNo.Text)))
    vtxtPARTDESC = N2Str2Null(txtPartDesc.Text)
    VTXTVEHTYPE = N2Str2Null(txtVehType.Text)
    VTXTMODELCODE = N2Str2Null(txtModelCode.Text)
    VTXTLocation = N2Str2Null(txtLocation.Text)
    VTXTDNP = NumericVal(txtDNP.Text)
    vtxtMAC = NumericVal(txtMAC.Text)
    VTXTMAD = NumericVal(txtMAD.Text)
    VTXTOLDNO = N2Str2Null(txtOldNo.Text)
    VTXTWFP = NumericVal(txtWFP.Text)
    VTXTSRP = NumericVal(txtSRP.Text)
    VTXTLastM_Mac = NumericVal(txtLastM_Mac.Text)
    VTXTLASTM_OH = NumericVal(txtLastM_Oh.Text)
    VTXTOnHand = NumericVal(txtOnHand.Text)
    VTXTTISSQTY = NumericVal(txtTissqty.Text)
    VTXTTPOQTY = NumericVal(txtTPOQty.Text)
    VTXTPHYCOUNT = NumericVal(txtPhyCount.Text)
    VTXTADJPHYCNT = NumericVal(txtAdjPhyCnt.Text)
    VTXTRECEIPTS = NumericVal(txtReceipts.Text)
    VTXTISSUANCES = NumericVal(txtIssuances.Text)
    VTXTSSTOCK = NumericVal(txtSStock.Text)
    VTXTRESSERVICE = NumericVal(txtResService.Text)
    Dim NON_HARI_STR                                   As String
    If chkNonHARI.Value = 1 Then
        NON_HARI_STR = "'Y'"
    Else
        NON_HARI_STR = "'N'"
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT INTO PMIS_STOCKMAS " & _
                      " (TYPE,LASTM_MAC,STOCKNO,STOCKDESC,VEHTYPE,MODELCODE,LOCATION,MAC,OLDNO,WFP,SRP,TISSQTY,TPOQTY,PHYCOUNT,ADJPHYCNT,RECEIPTS,ISSUANCES,SSTOCK,RESSERVICE,LASTUPDATE,USERCODE,ACTIVE,NON_HARI)" & _
                      " VALUES (" & N2Str2Null(LOCAL_STOCKTYPE) & "," & VTXTLastM_Mac & "," & vtxtPARTNO & "," & vtxtPARTDESC & ", " & VTXTVEHTYPE & ", " & _
                      " " & VTXTMODELCODE & ", " & VTXTLocation & ", " & vtxtMAC & _
                        ", " & VTXTOLDNO & ", " & VTXTWFP & _
                        ", " & VTXTSRP & ", " & VTXTTISSQTY & ", " & VTXTTPOQTY & _
                        ", " & VTXTPHYCOUNT & ", " & VTXTADJPHYCNT & ", " & VTXTRECEIPTS & ", " & VTXTISSUANCES & ", " & VTXTSSTOCK & ", " & VTXTRESSERVICE & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ",'N'," & NON_HARI_STR & ")"

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("A", LOCAL_ACCESS, SQL_STATEMENT, FindTransactionID(vtxtPARTNO, "STOCKNO", "PMIS_PARTMAS"), "", "PART CODE: " & vtxtPARTNO, "", "")

        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                      " LASTM_MAC = " & VTXTLastM_Mac & "," & _
                      " STOCKNO = " & vtxtPARTNO & "," & _
                      " STOCKDESC = " & vtxtPARTDESC & "," & _
                      " VEHTYPE = " & VTXTVEHTYPE & "," & _
                      " MODELCODE = " & VTXTMODELCODE & "," & _
                      " LOCATION = " & VTXTLocation & "," & _
                      " MAC = " & vtxtMAC & "," & _
                      " OLDNO = " & VTXTOLDNO & "," & _
                      " WFP = " & VTXTWFP & "," & _
                      " SRP = " & VTXTSRP & "," & _
                      " TPOQTY = " & VTXTTPOQTY & "," & _
                      " PHYCOUNT = " & VTXTPHYCOUNT & "," & _
                      " ADJPHYCNT = " & VTXTADJPHYCNT & "," & _
                      " RECEIPTS = " & VTXTRECEIPTS & "," & _
                      " ISSUANCES = " & VTXTISSUANCES & "," & _
                      " SSTOCK = " & VTXTSSTOCK & ", RESSERVICE = " & VTXTRESSERVICE & ", " & _
                      " LASTUPDATE = '" & LOGDATE & "'," & _
                      " NON_HARI = " & NON_HARI_STR & "," & _
                      " USERCODE = " & N2Str2Null(LOGCODE) & _
                      " WHERE ID = " & labID.Caption

        gconDMIS.Execute SQL_STATEMENT

        Call NEW_LogAudit("E", LOCAL_ACCESS, SQL_STATEMENT, labID, "", "PART CODE: " & vtxtPARTNO, "", "")

        ShowSuccessFullyUpdated
    End If
    rsRefresh
    On Error Resume Next
    RSPARTMAS.Find "STOCKNO=" & vtxtPARTNO
    cmdCancel.Value = True
    FillGrid

    Exit Sub
Errorcode:
    ShowVBError
    cmdCancel.Value = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = LOCAL_ACCESS & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Frame1.Enabled = False
    txtWFP.Enabled = False: txtDNP.Enabled = False:   'txtMAC.Enabled = True:
    txtLastM_Mac.Enabled = False

    If LOGLEVEL = "SUPERVISOR" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
        fraSupervisor.Enabled = True
        If LOGLEVEL = "AUTHOR" Then
            txtWFP.Enabled = True: txtDNP.Enabled = True:    'txtMAC.Enabled = True:
            txtLastM_Mac.Enabled = True
        End If
    Else
        fraSupervisor.Enabled = False
    End If

    textSearch.Text = ""
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub




Private Sub lstParts_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    RSPARTMAS.MoveFirst
    RSPARTMAS.Find ("ID=" & ITEM.ListSubItems(1).Text)
    StoreMemVars
End Sub

Private Sub lstParts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstParts
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

Private Sub lstParts_DblClick()
    If lstParts.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lstParts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub OPT_ACTIVE_Click()
    Dim rsParts                                        As ADODB.Recordset
    lstParts.Sorted = False
    Set rsParts = gconDMIS.Execute("SELECT STOCKNO, ID FROM PMIS_STOCKMAS WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND ACTIVE='Y'")
    Listview_Loadval Me.lstParts.ListItems, rsParts
    lstParts.Sorted = True
    rsRefresh
    initMemvars
    If lstParts.ListItems.Count = 0 Then
        Frame1.Enabled = False
        Picture1.Visible = False
    Else
        Frame1.Enabled = True
        Picture1.Visible = True
        StoreMemVars
    End If

End Sub

Private Sub OPT_ALL_Click()
    Dim rsParts                                        As ADODB.Recordset
    lstParts.Sorted = False
    Set rsParts = gconDMIS.Execute("SELECT STOCKNO, ID FROM PMIS_STOCKMAS WHERE TYPE='" & LOCAL_STOCKTYPE & "'")
    Listview_Loadval Me.lstParts.ListItems, rsParts
    lstParts.Sorted = True
    rsRefresh
    initMemvars
    If lstParts.ListItems.Count = 0 Then
        Frame1.Enabled = False
        Picture1.Visible = False
    Else
        Frame1.Enabled = True
        Picture1.Visible = True
        StoreMemVars
    End If

End Sub

Private Sub OPT_INACTIVE_Click()
    Dim rsParts                                        As ADODB.Recordset
    lstParts.Sorted = False
    Set rsParts = gconDMIS.Execute("SELECT STOCKNO, ID FROM PMIS_STOCKMAS WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND ISNULL(ACTIVE,'N')='N'")
    Listview_Loadval Me.lstParts.ListItems, rsParts
    lstParts.Sorted = True
    rsRefresh
    initMemvars
    If lstParts.ListItems.Count = 0 Then
        Frame1.Enabled = False
        Picture1.Visible = False
    Else
        Frame1.Enabled = True
        Picture1.Visible = True
        StoreMemVars
    End If


End Sub

Private Sub textSearch_Change()
    FillGrid
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstParts.ListItems.Count > 0 And lstParts.Enabled = True Then: lstParts.SetFocus
    End If
End Sub

Private Sub optDescription_Click()
    lstParts.ColumnHeaders(1).Text = "DESCRIPTION"
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optPARTNO_Click()
    lstParts.ColumnHeaders(1).Text = "PART NO."
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picAdds.Visible = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (" & LOCAL_ACCESS & " )"
            Call frmALL_AuditInquiry.DisplayHistory(labID, LOCAL_ACCESS, "")
    End Select
End Sub




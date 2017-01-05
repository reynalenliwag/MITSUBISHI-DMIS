VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPMISMaterials 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Master File"
   ClientHeight    =   6405
   ClientLeft      =   900
   ClientTop       =   315
   ClientWidth     =   8925
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Materials.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   8925
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3015
      ScaleHeight     =   855
      ScaleWidth      =   6015
      TabIndex        =   55
      Top             =   5460
      Width           =   6015
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
         MouseIcon       =   "Materials.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   56
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
         MouseIcon       =   "Materials.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   57
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
         MouseIcon       =   "Materials.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   58
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
         MouseIcon       =   "Materials.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   59
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
         MouseIcon       =   "Materials.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   60
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
         MouseIcon       =   "Materials.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   61
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
         MouseIcon       =   "Materials.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   62
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
         MouseIcon       =   "Materials.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
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
            Name            =   "Verdana"
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
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
            Text            =   "Text1"
            ToolTipText     =   "Input the number of the part in the safety stock (e.g. 1, 5, 3) "
            Top             =   2010
            Width           =   1185
         End
         Begin VB.TextBox txtReceipts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
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
            Text            =   "Text1"
            ToolTipText     =   "Type the number of receipts of the particular part, if there's any."
            Top             =   1620
            Width           =   1185
         End
         Begin VB.TextBox txtPhyCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
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
            Text            =   "Text1"
            ToolTipText     =   "Type the physical count of the particular part (e.g. 58, 60)"
            Top             =   1230
            Width           =   1185
         End
         Begin VB.TextBox txtTPOQty 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
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
            Text            =   "Text1"
            ToolTipText     =   "Type part's temporary PO."
            Top             =   840
            Width           =   1185
         End
         Begin VB.TextBox txtOnHand 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
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
            Text            =   "Text1"
            ToolTipText     =   "Type how many of that part are on hand (e.g. 5, 10)"
            Top             =   450
            Width           =   1185
         End
         Begin VB.TextBox txtLastM_Mac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
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
            Text            =   "Text1"
            Top             =   60
            Width           =   1185
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
               Name            =   "Verdana"
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
            Name            =   "Verdana"
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
         Width           =   3075
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
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
         Text            =   "Text1"
         ToolTipText     =   "Type the location where the part can be found (e.g. Q-1)"
         Top             =   1800
         Width           =   1755
      End
      Begin VB.TextBox txtModelCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
         Width           =   210
      End
      Begin VB.TextBox txtPartDesc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
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
         Top             =   2190
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
         Left            =   4710
         TabIndex        =   48
         Top             =   270
         Width           =   225
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "DNP"
         BeginProperty Font 
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
         Caption         =   "Material Code"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
         Caption         =   "&Material Code [Alt + M]"
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
         Left            =   180
         TabIndex        =   53
         Top             =   510
         Value           =   -1  'True
         Width           =   2385
      End
      Begin VB.OptionButton optDescription 
         Caption         =   "D&escription [Alt + E]"
         BeginProperty Font 
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
         Text            =   "Text1"
         Top             =   1080
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstParts 
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Materials.frx":2D71
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
            Name            =   "Verdana"
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
         MouseIcon       =   "Materials.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   65
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
         MouseIcon       =   "Materials.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   735
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
Attribute VB_Name = "frmPMISMaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPartMas                                          As ADODB.Recordset
Dim rsTDAYTRAN                                         As ADODB.Recordset
Dim AddorEdit                                          As String

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "MATERIALS MASTER FILE") = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Screen.MousePointer = 11
    rptPrintParts.ReportTitle = "ALPHABETICAL LISTING OF MATERIALS"
    rptPrintParts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrintParts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    
    PrintSQLReport rptPrintParts, PMIS_REPORT_PATH & "PrintMaterials.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "MATERIALS MASTER FILE") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    fraDetails.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    txtPartNo.Enabled = True
    
    '    txtPartNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    txtPartNo.Enabled = False
    fraDetails.Enabled = True
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "MATERIALS MASTER FILE") = False Then Exit Sub
    On Error GoTo Errorcode
    If Not rsPartMas.BOF Or Not rsPartMas.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from CSMS_MATMAS where id = " & labID.Caption
            LogAudit "X", "Materials Info: ", txtPartNo & "-" & txtPartDesc
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemvars
    FillSearchGrid textSearch
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "MATERIALS MASTER FILE") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    fraDetails.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
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
    rsPartMas.MoveNext
    If rsPartMas.EOF Then
        rsPartMas.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsPartMas.MovePrevious
    If rsPartMas.BOF Then
        rsPartMas.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim rsfindDup                                      As ADODB.Recordset
    
    
    If LTrim(RTrim(txtPartNo.Text)) = "" Then
        MsgSpeechBox "MATERIAL CODE MUST NOT BE EMPTY"
        On Error Resume Next
        txtModelCode.SetFocus
        Exit Sub
        End If
        

    Dim vtxtPARTNO, vtxtPARTDESC, VTXTVehType          As String
    Dim VTXTModelCode, VTXTLocation                    As String
    Dim vtxtMAC, VTXTDNP                               As Double
    Dim VTXTMAD, VTXTOldNo, VTXTWFP                    As String
    Dim VTXTSRP, VTXTLastM_Mac                         As Double
    Dim VTXTLastM_Oh, VTXTOnHand, VTXTTISSQty          As Integer
    Dim VTXTTPOQty, VTXTTPRQty, VTXTPhyCount           As Integer
    Dim VTXTAdjPhyCnt, VTXTReceipts, VTXTIssuances     As Integer
    Dim VTXTSStock, VTXTResService                     As Integer

    If IsNull(txtPartNo.Text) = True Then
        MsgSpeechBox "Part Number must not be empty"
        On Error Resume Next
        txtPartNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select STOCKNO from CSMS_MATMAS where STOCKNO = '" & txtPartNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    VTXTModelCode = N2Str2Null(txtModelCode.Text)
    VTXTLocation = N2Str2Null(txtLocation.Text)
    VTXTDNP = NumericVal(txtDNP.Text)
    vtxtMAC = NumericVal(txtMAC.Text)
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
            labID.Caption = NumericVal(rsPartMas!ID) + 1
        End If
        gconDMIS.Execute "Insert into CSMS_MATMAS" & _
                       " (TYPE,LastM_Mac,STOCKNO,STOCKDESC,vehtype,modelcode,location,oldno,wfp,srp,tissqty,tpoqty,phycount,adjphycnt,receipts,issuances,sstock,resservice,lastupdate,usercode,ACTIVE)" & _
                       " values ('M'," & VTXTLastM_Mac & "," & vtxtPARTNO & "," & vtxtPARTDESC & ", " & VTXTVehType & ", " & _
                       " " & VTXTModelCode & ", " & VTXTLocation & _
                         ", " & VTXTOldNo & ", " & VTXTWFP & _
                         ", " & VTXTSRP & ", " & VTXTTISSQty & ", " & VTXTTPOQty & _
                         ", " & VTXTPhyCount & ", " & VTXTAdjPhyCnt & ", " & VTXTReceipts & ", " & VTXTIssuances & ", " & VTXTSStock & ", " & VTXTResService & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ",'Y')"
        LogAudit "A", "Parts Info"
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "update CSMS_MATMAS set" & _
                       " LastM_Mac = " & VTXTLastM_Mac & "," & _
                       " STOCKNO = " & vtxtPARTNO & "," & _
                       " STOCKDESC = " & vtxtPARTDESC & "," & _
                       " vehtype = " & VTXTVehType & "," & _
                       " modelcode = " & VTXTModelCode & "," & _
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
                       " where id = " & labID.Caption
        ShowSuccessFullyUpdated
        LogAudit "E", "Materials Info: ", txtPartNo & "-" & txtPartDesc
    End If
    rsRefresh
    On Error Resume Next
    If AddorEdit = "EDIT" Then
        rsPartMas.Find "id =" & labID.Caption
    End If
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
    textSearch.Text = "":    'Picture3.ZOrder 0
    initMemvars
    StoreMemvars
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

Sub StoreMemvars()
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        labID.Caption = rsPartMas!ID
        txtPartNo.Text = Null2String(rsPartMas!STOCKNO)
        txtPartDesc.Text = Null2String(rsPartMas!STOCKDESC)
        txtVehType.Text = Null2String(rsPartMas!vehtype)
        txtModelCode.Text = Null2String(rsPartMas!modelcode)
        txtLocation.Text = Null2String(rsPartMas!Location)
        If frmMain.wizEnc1.DecryptAccess("524956t^yw9|kk") = LOGLEVEL Then
            txtWFP.Text = "Classified"
            txtDNP.Text = "Classified"
            txtMAC.Text = "Classified"
            txtLastM_Mac.Text = "Classified"
        Else
            txtWFP.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!WFP))
            txtDNP.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!dnp))
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
    rsPartMas.Open "select * from CSMS_MATMAS order by STOCKNO asc", gconDMIS, adOpenForwardOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISParts = Nothing
    UnloadForm Me
End Sub

Private Sub lstParts_GotFocus()
    If optPartNo.Value = True Then
        rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "STOCKNO", lstParts.SelectedItem).Bookmark
    Else
        rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "STOCKNO", lstParts.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemvars
End Sub

Private Sub lstParts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    If optPartNo.Value = True Then
        rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "STOCKNO", Item).Bookmark
    Else
        rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "STOCKNO", lstParts.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemvars
End Sub

Private Sub lstParts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstParts
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstParts_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstParts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If optPartNo.Value = True Then
        If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    Else
        If Trim(textSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    End If
End Sub

Sub FillGrid()
    Dim rsParts                                        As ADODB.Recordset
    lstParts.Enabled = False
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select STOCKNO,ID from CSMS_MATMAS  order by STOCKNO asc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstParts.Enabled = True: Listview_Loadval Me.lstParts.ListItems, rsParts: lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsParts                                        As ADODB.Recordset
    lstParts.Enabled = False
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsParts = gconDMIS.Execute("select STOCKNO, ID from CSMS_MATMAS where STOCKNO like'" & XXX & "%'")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstParts.Enabled = True: Listview_Loadval Me.lstParts.ListItems, rsParts: lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsParts                                        As ADODB.Recordset
    lstParts.Enabled = False
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select STOCKDESC, STOCKNO from CSMS_MATMAS  order by STOCKDESC asc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstParts.Enabled = True: Listview_Loadval Me.lstParts.ListItems, rsParts: lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsParts                                        As ADODB.Recordset
    lstParts.Enabled = False
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsParts = gconDMIS.Execute("select STOCKDESC, STOCKNO from CSMS_MATMAS where STOCKDESC like '" & ReplaceQuote(XXX) & "%' order by STOCKNO asc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstParts.Enabled = True: Listview_Loadval Me.lstParts.ListItems, rsParts: lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstParts.ListItems.Count > 0 And lstParts.Enabled = True Then: lstParts.SetFocus
    End If
End Sub

Private Sub optDescription_Click()
    lstParts.ColumnHeaders(1).Text = "DESCRIPTION"
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optPARTNO_Click()
    lstParts.ColumnHeaders(1).Text = "PART NO."
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

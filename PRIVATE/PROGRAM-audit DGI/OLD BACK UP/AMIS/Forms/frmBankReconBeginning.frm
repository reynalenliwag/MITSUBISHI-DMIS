VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBankReconBeginning 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beginning Bank Recon"
   ClientHeight    =   8175
   ClientLeft      =   10170
   ClientTop       =   3885
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "frmBankReconBeginning.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   9795
   Begin VB.TextBox txtAdjustedBook 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0.00"
      Top             =   7740
      Width           =   1815
   End
   Begin VB.TextBox txtAdjustedBank 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7890
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "0.00"
      Top             =   7740
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   195
      Left            =   150
      TabIndex        =   67
      Top             =   6960
      Width           =   4785
   End
   Begin VB.TextBox txtCredit 
      Alignment       =   1  'Right Justify
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
      Left            =   7875
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   25
      Text            =   "0.00"
      Top             =   6900
      Width           =   1785
   End
   Begin VB.TextBox txtDebit 
      Alignment       =   1  'Right Justify
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
      Left            =   6015
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   6900
      Width           =   1815
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   90
      ScaleHeight     =   2325
      ScaleWidth      =   9735
      TabIndex        =   60
      Top             =   4560
      Width           =   9735
      Begin MSComctlLib.ListView lvTransactions 
         Height          =   2295
         Left            =   30
         TabIndex        =   23
         Top             =   0
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   4048
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
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6086
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Reference"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Deposits"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Withdrawals"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2325
      Left            =   150
      ScaleHeight     =   2265
      ScaleWidth      =   9435
      TabIndex        =   52
      Top             =   60
      Width           =   9495
      Begin VB.Frame Frame1 
         Height          =   165
         Left            =   60
         TabIndex        =   63
         Top             =   1440
         Width           =   9375
      End
      Begin VB.TextBox txtJDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7830
         MaxLength       =   10
         TabIndex        =   62
         Text            =   "88/88/8888"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.OptionButton optOutstanding 
         BackColor       =   &H80000018&
         Caption         =   "&Outstanding Checks"
         Height          =   465
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1710
         Width           =   2895
      End
      Begin VB.OptionButton optDeposits 
         BackColor       =   &H80000018&
         Caption         =   "&Deposits in Transit"
         Height          =   465
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1710
         Width           =   2895
      End
      Begin VB.TextBox txtBank 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7590
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1110
         Width           =   1815
      End
      Begin VB.TextBox txtBook 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5730
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   1110
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtCurrent 
         Height          =   315
         Left            =   7680
         TabIndex        =   1
         Top             =   150
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52232193
         CurrentDate     =   40002
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Index           =   3
         Left            =   7170
         TabIndex        =   65
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblBankName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   150
         Width           =   7005
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   60
         Top             =   60
         Width           =   9345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Journal Date"
         ForeColor       =   &H00701E2A&
         Height          =   210
         Left            =   6600
         TabIndex        =   61
         Top             =   2400
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   0
         Left            =   7560
         TabIndex        =   55
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   5730
         TabIndex        =   54
         Top             =   870
         Width           =   525
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unadjusted Balance:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   1
         Left            =   3630
         TabIndex        =   53
         Top             =   1230
         Width           =   2040
      End
   End
   Begin VB.PictureBox picAdd 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   -4950
      ScaleHeight     =   900
      ScaleWidth      =   9735
      TabIndex        =   58
      Top             =   7260
      Width           =   9735
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
         Left            =   8820
         MouseIcon       =   "frmBankReconBeginning.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmBankReconBeginning.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   765
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
         Left            =   8070
         MouseIcon       =   "frmBankReconBeginning.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "frmBankReconBeginning.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdUnPost 
         Enabled         =   0   'False
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
         Left            =   7320
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmBankReconBeginning.frx":123A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   30
         Width           =   765
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
         Left            =   6570
         MouseIcon       =   "frmBankReconBeginning.frx":138C
         MousePointer    =   99  'Custom
         Picture         =   "frmBankReconBeginning.frx":14DE
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
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
         Left            =   5820
         MouseIcon       =   "frmBankReconBeginning.frx":1809
         MousePointer    =   99  'Custom
         Picture         =   "frmBankReconBeginning.frx":195B
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   765
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
         Left            =   5070
         MouseIcon       =   "frmBankReconBeginning.frx":1CB7
         MousePointer    =   99  'Custom
         Picture         =   "frmBankReconBeginning.frx":1E09
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   765
      End
      Begin VB.Label lblID 
         Height          =   345
         Left            =   720
         TabIndex        =   59
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.PictureBox picSave 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   3120
      ScaleHeight     =   885
      ScaleWidth      =   1590
      TabIndex        =   57
      Top             =   7260
      Width           =   1590
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
         Left            =   765
         MouseIcon       =   "frmBankReconBeginning.frx":211C
         MousePointer    =   99  'Custom
         Picture         =   "frmBankReconBeginning.frx":226E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   765
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
         Left            =   10
         MouseIcon       =   "frmBankReconBeginning.frx":25AC
         MousePointer    =   99  'Custom
         Picture         =   "frmBankReconBeginning.frx":26FE
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox picPayables 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   150
      ScaleHeight     =   2085
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   2430
      Width           =   9495
      Begin VB.CommandButton cmdVendor 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9000
         TabIndex        =   17
         Top             =   60
         Width           =   405
      End
      Begin VB.TextBox txtVendorName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   16
         Top             =   75
         Width           =   6435
      End
      Begin VB.TextBox txtVendorCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   15
         Text            =   "000226"
         Top             =   75
         Width           =   1005
      End
      Begin VB.TextBox txtCheckNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   19
         Text            =   "000226"
         Top             =   870
         Width           =   1485
      End
      Begin VB.ComboBox cboBankName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4530
         TabIndex        =   36
         Top             =   480
         Width           =   4890
      End
      Begin VB.TextBox txtBankCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   18
         Text            =   "000226"
         Top             =   480
         Width           =   1485
      End
      Begin VB.TextBox txtCheckAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   1650
         Width           =   1485
      End
      Begin VB.TextBox txtCheckDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "88/88/8888"
         Top             =   1260
         Width           =   1485
      End
      Begin RichTextLib.RichTextBox txtParticulars 
         Height          =   705
         Left            =   4530
         TabIndex        =   22
         Top             =   900
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   1244
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmBankReconBeginning.frx":2A4E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Code"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   56
         Top             =   90
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3390
         TabIndex        =   42
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3390
         TabIndex        =   41
         Top             =   900
         Width           =   990
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check No."
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   40
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Code"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   39
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amount"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   38
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   37
         Top             =   1290
         Width           =   1080
      End
   End
   Begin VB.PictureBox picReceivable 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   150
      ScaleHeight     =   2085
      ScaleWidth      =   9465
      TabIndex        =   43
      Top             =   2430
      Width           =   9495
      Begin VB.CommandButton cmdCustomer 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9000
         TabIndex        =   8
         Top             =   60
         Width           =   405
      End
      Begin VB.TextBox txtCustomerName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   7
         Top             =   75
         Width           =   6435
      End
      Begin VB.ComboBox cboBankName2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4530
         TabIndex        =   10
         Top             =   480
         Width           =   4890
      End
      Begin VB.ComboBox cboPayType 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   360
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   1500
      End
      Begin VB.TextBox txtORAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   1680
         Width           =   1485
      End
      Begin VB.TextBox txtORDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "88/88/8888"
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox txtCustCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "000226"
         Top             =   75
         Width           =   1005
      End
      Begin VB.TextBox txtORNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "000000"
         Top             =   900
         Width           =   1485
      End
      Begin RichTextLib.RichTextBox txtParticulars2 
         Height          =   705
         Left            =   4530
         TabIndex        =   14
         Top             =   900
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   1244
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmBankReconBeginning.frx":2AE2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkNonVat 
         Caption         =   "Non-Vat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         TabIndex        =   44
         Top             =   900
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label labBankName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3390
         TabIndex        =   51
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label labType 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   50
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Code"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   49
         Top             =   90
         Width           =   1050
      End
      Begin VB.Label labParticulars 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3390
         TabIndex        =   48
         Top             =   900
         Width           =   990
      End
      Begin VB.Label labAmt 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "O.R. Amount"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   47
         Top             =   1710
         Width           =   1170
      End
      Begin VB.Label labDate 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "O.R. Date"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   46
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label LabNo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "O.R. No."
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   45
         Top             =   930
         Width           =   765
      End
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjusted Balance:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   480
      Index           =   3
      Left            =   5070
      TabIndex        =   70
      Top             =   7500
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   6030
      TabIndex        =   69
      Top             =   7500
      Width           =   525
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   2
      Left            =   7920
      TabIndex        =   68
      Top             =   7500
      Width           =   495
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5040
      TabIndex        =   66
      Top             =   6990
      Width           =   705
   End
   Begin VB.Line Line1 
      X1              =   7200
      X2              =   7200
      Y1              =   60
      Y2              =   600
   End
End
Attribute VB_Name = "frmBankReconBeginning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddorEdit                                     As String
Dim WithEvents frmSearchCustomerVendor            As frmAMISCustomerVendor
Attribute frmSearchCustomerVendor.VB_VarHelpID = -1

Private Sub cboBankName_Click()
    txtBankCode.Text = SetBankCode(cboBankName)
End Sub

Private Sub cboBankName_LostFocus()
    Dim rsBankCode                                As ADODB.Recordset
    Set rsBankCode = New ADODB.Recordset
    rsBankCode.Open "Select BankName from All_Banks where BankName = '" & cboBankName.Text & "'", gconDMIS, adOpenKeyset
    If Not rsBankCode.EOF And Not rsBankCode.BOF Then
    Else
        MsgBox "Please select Bank Name from the list.", vbInformation, "Message"
        cboBankName.SetFocus
        Exit Sub
    End If
    Set rsBankCode = Nothing
End Sub

Private Sub cboBankName2_Click()
    txtBankCode.Text = SetBankCode(cboBankName2)
End Sub

Private Sub cboBankName2_LostFocus()
    Dim rsBankCode                                As ADODB.Recordset
    Set rsBankCode = New ADODB.Recordset
    rsBankCode.Open "Select BankName from All_Banks where BankName = '" & cboBankName2.Text & "'", gconDMIS, adOpenKeyset
    If Not rsBankCode.EOF And Not rsBankCode.BOF Then
    Else
        MsgBox "Please select Bank Name from the list.", vbInformation, "Message"
        cboBankName2.SetFocus
        Exit Sub
    End If
    Set rsBankCode = Nothing
End Sub

Private Sub cmdAdd_Click()
    initMemvars
    picAdd.Visible = False
    picSave.Visible = True
    Picture1.Enabled = True
    optDeposits.Value = False
    optOutstanding.Value = False
    txtBook.SetFocus
    AddorEdit = "ADD"
    xSELECTED = ""
End Sub

Private Sub cmdCancel_Click()
    picSave.Visible = False
    picAdd.Visible = True
    Picture1.Enabled = False
    picReceivable.Enabled = False
    picPayables.Enabled = False
End Sub

Private Sub cmdCustomer_Click()
    Set frmSearchCustomerVendor = New frmAMISCustomerVendor
    frmSearchCustomerVendor.Show 1
End Sub

Private Sub cmdDelete_Click()
    If lblID.Caption = "" Then
        'MsgBox "No record to delete. Please select from the list.", vbInformation, "Message"
        MessagePop RecNotFound, "INFORMATION", "No record to delete. Select from the list."
        Exit Sub
    Else
        If MsgBox("Date: " & lvTransactions.SelectedItem.Text & "  Reference No.: " & lvTransactions.SelectedItem.SubItems(2) & vbCrLf & "Are you sure you want to delete this transaction?", vbQuestion + vbYesNo, "Delete") = vbYes Then
            SQL_STATEMENT = "Delete from AMIS_RECONBEGINNING where ID = '" & lblID.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT
            Call FillGrid(lblBankName)
            Call GetBankName(lblBankName)
            MessagePop Delete, "INFORMATION", "Record deleted."
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    initMemvars
    picAdd.Visible = False
    picSave.Visible = True
    Picture1.Enabled = True
    optDeposits.Value = False
    optOutstanding.Value = False
    txtBook.SetFocus
    AddorEdit = "EDIT"
    xSELECTED = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim xlApplication                             As Excel.Application
    Dim xlBook                                    As Excel.Workbook
    Dim xlSheet                                   As Excel.Worksheet
    Dim xlRange                                   As Excel.Range
    Dim rsBankOpening                             As ADODB.Recordset
    Dim xCounter                                  As Integer
    Dim xListCounter                              As Integer
    Dim xdebit                                    As Double
    Dim xcredit                                   As Double
    Set xlApplication = CreateObject("Excel.Application")
    Set xlBook = xlApplication.Workbooks.Open(AMIS_REPORT_PATH & "JOURNALS\BankRecon.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    xlSheet.Cells(1, "A") = COMPANY_NAME
    xlSheet.Cells(2, "A") = COMPANY_ADDRESS
    xlSheet.Cells(3, "A") = "BANK RECONCILIATION REPORT"
    xlSheet.Cells(5, "A") = "As of: " & dtCurrent.Value
    xlSheet.Cells(8, "A") = "Unadjusted Balance"
    xlSheet.Cells(8, "A").Font.Bold = True
    xlSheet.Cells(8, "C") = txtBook
    xlSheet.Cells(8, "C").Font.Bold = True
    xlSheet.Cells(8, "D") = txtBank
    xlSheet.Cells(8, "D").Font.Bold = True
    xCounter = 9
    xlSheet.Cells(xCounter, "A") = "Deposits in Transit"
    xlSheet.Cells(xCounter, "A").Font.Bold = True
    For xListCounter = 1 To lvTransactions.ListItems.Count
        If Left(lvTransactions.ListItems.Item(xListCounter).SubItems(2), 2) = "OR" Then
            xlSheet.Cells(xCounter, "D") = lvTransactions.ListItems.Item(xListCounter).Text
            xlSheet.Cells(xCounter, "E") = lvTransactions.ListItems.Item(xListCounter).SubItems(2)
            xlSheet.Cells(xCounter, "F") = lvTransactions.ListItems.Item(xListCounter).SubItems(3)
            xdebit = xdebit + NumericVal(lvTransactions.ListItems.Item(xListCounter).SubItems(3))
            xCounter = xCounter + 1
        End If
    Next xListCounter
    xCounter = xCounter + 1
    xlSheet.Cells(xCounter, "A") = "Outstanding Checks"
    xlSheet.Cells(xCounter, "A").Font.Bold = True
    For xListCounter = 1 To lvTransactions.ListItems.Count
        If Left(lvTransactions.ListItems.Item(xListCounter).SubItems(2), 3) = "CHK" Then
            xlSheet.Cells(xCounter, "D") = lvTransactions.ListItems.Item(xListCounter).Text
            xlSheet.Cells(xCounter, "E") = lvTransactions.ListItems.Item(xListCounter).SubItems(2)
            xlSheet.Cells(xCounter, "F") = lvTransactions.ListItems.Item(xListCounter).SubItems(4)
            xcredit = xcredit + NumericVal(lvTransactions.ListItems.Item(xListCounter).SubItems(4))
            xCounter = xCounter + 1
        End If
    Next xListCounter

    xlSheet.Cells(xCounter + 1, "A") = "Adjustments"
    xlSheet.Cells(xCounter + 1, "A").Font.Bold = True
    xlSheet.Cells(xCounter + 2, "B") = "Interest"
    xlSheet.Cells(xCounter + 3, "B") = "Bank Charges"
    xlSheet.Cells(xCounter + 4, "D") = "Unidentified Deposit"
    xlSheet.Cells(xCounter + 5, "D") = "Unidentified Bank Charges"
    xlSheet.Cells(xCounter + 7, "A") = "Adjusted Book Balance"
    xlSheet.Cells(xCounter + 7, "A").Font.Bold = True
    xlSheet.Cells(xCounter + 7, "C") = txtBook.Text
    xlSheet.Cells(xCounter + 7, "C").Font.Bold = True
    xlSheet.Cells(xCounter + 7, "D") = "Adjusted Bank Balance"
    xlSheet.Cells(xCounter + 7, "D").Font.Bold = True
    xlSheet.Cells(xCounter + 7, "F") = NumericVal(txtBank.Text) + (xdebit + xcredit)
    xlSheet.Cells(xCounter + 7, "F").Font.Bold = True

    xlApplication.Visible = True
    Set xlApplication = Nothing
End Sub

Private Sub cmdSave_Click()
    If optDeposits.Value = False And optOutstanding.Value = False Then
        Dim rsBegRecon                            As ADODB.Recordset
        Set rsBegRecon = New ADODB.Recordset
        rsBegRecon.Open "select * from All_Banks where BankName ='" & lblBankName.Caption & "'", gconDMIS, adOpenForwardOnly
        If Not rsBegRecon.EOF And Not rsBegRecon.BOF Then
            '        If MsgBox("Are all entries correct?", vbQuestion + vbYesNo, "Question") = vbYes Then
            If txtBank.Text = 0 Or txtBook.Text = 0 Then
                MsgBox "Please enter Beginnning Balance", vbExclamation, "Beginning Balance"
                Exit Sub
            Else
                gconDMIS.Execute "Update All_Banks Set Beginning_Bank='" & NumericVal(txtBank.Text) & "',Beginning_Book='" & NumericVal(txtBook.Text) & "',LastDate_Recon='" & dtCurrent.Value & "' where BankName='" & lblBankName.Caption & "'"
                MessagePop RecSave, "INFORMATION", "Unadjusted Balance updated."
            End If
            '        Else
            '            Exit Sub
            '        End If
        End If
    Else
        Dim vJDate                                As String
        Dim vVendorCode                           As String
        Dim vBankCode                             As String
        Dim vCheckNo                              As String
        Dim vCheck_Date                           As String
        Dim vCheck_Amt                            As Double
        Dim vRemarks                              As String
        Dim vCustCode                             As String
        Dim vPayType                              As String
        Dim vOR_Num                               As String
        Dim vOR_Date                              As String
        Dim vOR_Amt

        If txtJDate.Text = "" Or IsDate(txtJDate.Text) = False Then
            MsgBox "Invalid Date", vbInformation, "Message"
            Exit Sub
        End If
        If xSELECTED = "Customer" Then
            If txtCustCode.Text = "" Then
                MsgBox "Please select customer name", vbInformation, "Missing entry"
                Exit Sub
            ElseIf cboPayType.Text = "" Then
                MsgBox "Please select pay type", vbInformation, "Missing entry"
                Exit Sub
            ElseIf txtBankCode.Text = "" Then
                MsgBox "Please select Bank Name", vbInformation, "Missing entry"
                Exit Sub
            ElseIf txtORNo.Text = "" Then
                MsgBox "Please enter OR No.", vbInformation, "Missing entry"
                Exit Sub
            ElseIf txtORDate.Text = "" Or IsDate(txtORDate.Text) = False Then
                MsgBox "Invalid Date", vbInformation, "Message"
                Exit Sub
            ElseIf txtORAmount.Text = "" Or NumericVal(txtORAmount.Text) = 0 Then
                MsgBox "Please enter OR Amount", vbInformation, "Missing entry"
                Exit Sub
            End If
        ElseIf xSELECTED = "Vendor" Then
            If txtVendorCode.Text = "" Then
                MsgBox "Please select vendor name", vbInformation, "Missing entry"
                Exit Sub
            ElseIf txtBankCode.Text = "" Then
                MsgBox "Please select Bank Name", vbInformation, "Missing entry"
                Exit Sub
            ElseIf txtCheckNo.Text = "" Then
                MsgBox "Please enter Check No.", vbInformation, "Missing entry"
                Exit Sub
            ElseIf txtCheckDate.Text = "" Or IsDate(txtCheckDate.Text) = False Then
                MsgBox "Invalid Date", vbInformation, "Message"
                Exit Sub
            ElseIf txtCheckAmt.Text = "" Or NumericVal(txtCheckAmt.Text) = 0 Then
                MsgBox "Please enter Check Amount", vbInformation, "Missing entry"
                Exit Sub
            End If
        Else
            MsgBox "Please select from Deposits in Transit or Outstanding Checks", vbInformation, "Select"
            optDeposits.SetFocus
            Exit Sub
        End If
        vCustCode = N2Str2Null(txtCustCode.Text)
        vPayType = N2Str2Null(cboPayType.Text)
        vOR_Num = N2Str2Null(txtORNo.Text)
        vOR_Date = N2Date2Null(txtORDate.Text)
        vOR_Amt = NumericVal(txtORAmount.Text)
        vVendorCode = N2Str2Null(txtVendorCode.Text)
        vCheckNo = N2Str2Null(txtCheckNo.Text)
        vCheck_Date = N2Date2Null(txtCheckDate.Text)
        vCheck_Amt = NumericVal(txtCheckAmt.Text)
        vBankCode = N2Str2Null(txtBankCode.Text)
        vJDate = N2Date2Null(txtJDate.Text)
        If xSELECTED = "Customer" Then
            If Trim(txtParticulars2.Text) = "Type Your Message Here!" Then vRemarks = "NULL" Else vRemarks = N2Str2Null(Trim(txtParticulars2.Text))
        Else
            If Trim(txtParticulars.Text) = "Type Your Message Here!" Then vRemarks = "NULL" Else vRemarks = N2Str2Null(Trim(txtParticulars.Text))
        End If

        If AddorEdit = "ADD" Then
            SQL_STATEMENT = "Insert into AMIS_RECONBEGINNING(JDATE,CUSTOMERCODE,VENDORCODE,PAYTYPE,BANKCODE,OR_NUM,OR_DATE,OR_AMT,CHECKNO,CHECK_DATE,CHECK_AMT,REMARKS) Values (" & vJDate & "," & vCustCode & "," & vVendorCode & "," & vPayType & "," & vBankCode & "," & vOR_Num & "," & vOR_Date & "," & vOR_Amt & "," & vCheckNo & "," & vCheck_Date & "," & vCheck_Amt & "," & vRemarks & ")"
            gconDMIS.Execute SQL_STATEMENT
            MessagePop RecSave, "INFORMATION", "Record saved"
        Else
            SQL_STATEMENT = "Update AMIS_RECONBEGINNING Set " & _
                            "JDATE= " & vJDate & "," & _
                            "CUSTOMERCODE =" & vCustCode & "," & _
                            "VENDORCODE =" & vVendorCode & "," & _
                            "PAYTYPE = " & vPayType & "," & _
                            "BANKCODE = " & vBankCode & "," & _
                            "OR_NUM = " & vOR_Num & "," & _
                            "OR_DATE = " & vOR_Date & "," & _
                            "OR_AMT = " & vOR_Amt & "," & _
                            "CHECKNO = " & vCheckNo & "," & _
                            "CHECK_DATE = " & vCheck_Date & "," & _
                            "CHECK_AMT = " & vCheck_Amt & ", " & _
                            "REMARKS =" & vRemarks & _
                            "WHERE ID = '" & lblID.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT
            MessagePop RecSave, "INFORMATION", "Record updated"
        End If
    End If
    Call FillGrid(lblBankName)
    Call GetBankName(lblBankName)
    picSave.Visible = False
    picAdd.Visible = True
    Picture1.Enabled = False
    picReceivable.Enabled = False
    picPayables.Enabled = False
    'picList.Enabled = False
End Sub

Private Sub cmdVendor_Click()
    Set frmSearchCustomerVendor = New frmAMISCustomerVendor
    frmSearchCustomerVendor.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtCurrent.Value = frmReconcileAccount.dtCurrent.Value - Format(frmReconcileAccount.dtCurrent.Value, "dd")
    lblBankName.Caption = frmReconcileAccount.lblBankName
    Call FillGrid(lblBankName)
    initMemvars
    Call GetBankName(lblBankName)
    Dim frmSsearchCustomerVendor                  As New frmAMISCustomerVendor
End Sub

Private Sub lblAccount_Change()
'lblBankName.Caption = BankName(lblAccount.Caption)
End Sub

Private Sub lvTransactions_DblClick()
    Dim rsTransactions                            As ADODB.Recordset
    Set rsTransactions = New ADODB.Recordset
    rsTransactions.Open "Select * from AMIS_RECONBEGINNING where ID = '" & lblID.Caption & "'", gconDMIS, adOpenKeyset
    If Not rsTransactions.EOF And Not rsTransactions.BOF Then
        txtBankCode.Text = Null2String(rsTransactions!bankcode)
        If xSELECTED = "Customer" Then
            txtCustCode.Text = Null2String(rsTransactions!CustomerCode)
            cboBankName2.Text = SetBankName(txtBankCode.Text)
            txtORNo.Text = Format(NumericVal(rsTransactions!OR_NUM), "000000")
            txtORDate.Text = Null2String(rsTransactions!OR_DATE)
            txtORAmount.Text = ToDoubleNumber(rsTransactions!OR_AMT)
            txtVendorCode.Text = ""
            txtCheckNo.Text = ""
            txtCheckAmt.Text = ""
            txtParticulars2.Text = Null2String(rsTransactions!remarks)
            If NumericVal(rsTransactions!OR_AMT) > 0 Then
                cboPayType.Text = Null2String(rsTransactions!paytype)
            End If
        Else
            txtVendorCode.Text = Null2String(rsTransactions!VendorCode)
            cboBankName.Text = SetBankName(txtBankCode.Text)
            txtCheckNo.Text = Null2String(rsTransactions!CheckNo)
            txtCheckDate.Text = Null2String(rsTransactions!Check_Date)
            txtCheckAmt.Text = ToDoubleNumber(rsTransactions!CHECK_AMT)
            txtParticulars.Text = Null2String(rsTransactions!remarks)
            txtCustCode.Text = ""
            txtORNo.Text = ""
            txtORAmount.Text = ""
        End If
    End If
End Sub

Private Sub lvTransactions_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblID.Caption = lvTransactions.SelectedItem.SubItems(5)
End Sub

Private Sub optDeposits_Click()
    picReceivable.Visible = True
    picPayables.Visible = False
    picReceivable.Enabled = True
    picList.Enabled = True
    InitCombo
    xSELECTED = "Customer"
    cmdCustomer.SetFocus
End Sub

Private Sub optOutstanding_Click()
    picReceivable.Visible = False
    picPayables.Visible = True
    picPayables.Enabled = True
    picList.Enabled = True
    InitCombo
    xSELECTED = "Vendor"
    cmdVendor.SetFocus
End Sub

Private Sub txtBank_GotFocus()
    txtBank.Text = NumericVal(txtBank.Text)
End Sub

Private Sub txtBank_KeyPress(KeyAscii As Integer)
'KeyAscii = OnlyNumeric(KeyAscii)
    Select Case KeyAscii
    Case 8
    Case 45
    Case 46
    Case 48 To 57
    Case Else
        KeyAscii = 0
        MsgBox "Invalid character!", vbExclamation, "Message"
    End Select

    If KeyAscii = 13 Then
        '        If NumericVal(txtBank.Text) <= 0 Then
        '            MsgBox "Beginning Balance must be greater than 0", vbExclamation, "Beginnning Balance"
        '            txtBank.SetFocus
        '        Else
        txtBook.SetFocus
        '        End If
    End If
End Sub

Private Sub txtBank_LostFocus()
    txtBank.Text = ToDoubleNumber(txtBank.Text)
End Sub

Private Sub txtBook_GotFocus()
    txtBook.Text = NumericVal(txtBook.Text)
End Sub

Private Sub txtBook_KeyPress(KeyAscii As Integer)
'KeyAscii = OnlyNumeric(KeyAscii)
    Select Case KeyAscii
    Case 8
    Case 45
    Case 46
    Case 48 To 57
    Case Else
        KeyAscii = 0
        MsgBox "Invalid character!", vbExclamation, "Message"
    End Select

    If KeyAscii = 13 Then
        '        If NumericVal(txtBook.Text) <= 0 Then
        '            MsgBox "Beginning Balance must be greater than 0", vbExclamation, "Beginnning Balance"
        '            txtBook.SetFocus
        '        Else
        '        End If
    End If
End Sub

Private Sub txtBook_LostFocus()
    txtBook.Text = ToDoubleNumber(txtBook.Text)
End Sub

Sub GetBankName(xBankName As String)
    Dim rsBankName                                As ADODB.Recordset
    Set rsBankName = New ADODB.Recordset
    rsBankName.Open "Select * from All_Banks where BankName = '" & xBankName & "'", gconDMIS, adOpenForwardOnly
    If Not rsBankName.EOF And Not rsBankName.BOF Then
        BankName = rsBankName!BankName
        txtBook.Text = ToDoubleNumber(NumericVal(rsBankName!Beginning_Book))
        txtBank.Text = ToDoubleNumber(NumericVal(rsBankName!Beginning_Bank))
        txtAdjustedBook.Text = ToDoubleNumber(NumericVal(rsBankName!Beginning_Book))
        txtAdjustedBank.Text = ToDoubleNumber(NumericVal(rsBankName!Beginning_Bank) + (NumericVal(txtDebit.Text) + NumericVal(txtCredit.Text)))
    End If
    Set rsBankName = Nothing
End Sub

Sub initMemvars()
    Dim txt                                       As Control
    Dim cbo                                       As Control
    '    For Each txt In Me.ControlS
    '        If TypeOf txt Is TextBox Then
    '            txt.Text = ""
    '        End If
    '    Next
    txtCustCode.Text = ""
    txtCustomerName.Text = ""
    cboPayType.ListIndex = -1
    txtBankCode.Text = ""
    cboBankName2.Text = ""
    txtORNo.Text = ""
    txtORDate.Text = ""
    txtORAmount.Text = ""
    txtParticulars2.Text = "Type Your Message Here!"

    txtVendorCode.Text = ""
    txtVendorName.Text = ""
    txtBankCode.Text = ""
    cboBankName.Text = ""
    txtCheckNo.Text = ""
    txtCheckDate.Text = ""
    txtCheckAmt.Text = ""
    txtParticulars.Text = "Type Your Message Here!"

    For Each cbo In Me.ControlS
        If TypeOf cbo Is ComboBox Then
            cbo.ListIndex = -1
        End If
    Next

    txtJDate.Text = LOGDATE
    txtORDate.Text = LOGDATE
    txtCheckDate.Text = LOGDATE
    txtORAmount.Text = "0.00"
    txtCheckAmt.Text = "0.00"
    cboPayType.Clear
    cboPayType.AddItem "CASH"
    cboPayType.AddItem "CARD"
    cboPayType.AddItem "CHECK"
    picReceivable.Enabled = False
    picReceivable.ZOrder 0
    picPayables.Enabled = False
    Picture1.Enabled = False
End Sub

Sub InitCombo()
    If optDeposits.Value = True Then
        Dim rsBanks2                              As ADODB.Recordset
        Set rsBanks2 = New ADODB.Recordset
        rsBanks2.Open "Select BankName from All_Banks Order by BankName", gconDMIS, adOpenKeyset
        If Not rsBanks2.EOF And Not rsBanks2.BOF Then
            Combo_Loadval cboBankName2, rsBanks2
        End If
    End If
    If optOutstanding.Value = True Then
        Dim rsBanks                               As ADODB.Recordset
        Set rsBanks = New ADODB.Recordset
        rsBanks.Open "Select BankName from All_Banks Order by BankName", gconDMIS, adOpenKeyset
        If Not rsBanks.EOF And Not rsBanks.BOF Then
            Combo_Loadval cboBankName, rsBanks
        End If
    End If
    Set rsBanks2 = Nothing
    Set rsBanks = Nothing
End Sub

Function SetCustomerName(XCustomerCode As String) As String
    Dim rsCustomer                                As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "Select CustName from All_CustMaster_AMIS where CustCode = '" & XCustomerCode & "'", gconDMIS, adOpenKeyset
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = Null2String(rsCustomer!CUSTNAME)
    End If
    Set rsCustomer = Nothing
End Function

Function SetVendorName(xVENDORCODE As String) As String
    Dim rsVENDOR                                  As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select NameofVendor from All_Vendor where Code = '" & xVENDORCODE & "' Order by NameofVendor", gconDMIS, adOpenKeyset
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    End If
    Set rsVENDOR = Nothing
End Function

Function SetBankCode(xBankName As String) As String
    Dim rsBank                                    As ADODB.Recordset
    Set rsBank = New ADODB.Recordset
    rsBank.Open "Select BankCode from All_Banks where BankName = '" & xBankName & "'", gconDMIS, adOpenKeyset
    If Not rsBank.EOF And Not rsBank.BOF Then
        SetBankCode = Null2String(rsBank!bankcode)
    End If
    Set rsBank = Nothing
End Function

Function SetBankName(xBankCode As String) As String
    Dim rsBank                                    As ADODB.Recordset
    Set rsBank = New ADODB.Recordset
    rsBank.Open "Select BankName from All_Banks where BankCode ='" & xBankCode & "'", gconDMIS, adOpenKeyset
    If Not rsBank.EOF And Not rsBank.BOF Then
        SetBankName = Null2String(rsBank!BankName)
    End If
    Set rsBank = Nothing
End Function

Private Sub txtCheckAmt_GotFocus()
    txtCheckAmt.Text = NumericVal(txtCheckAmt.Text)
End Sub

Private Sub txtCheckAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 8
    Case 45
    Case 46
    Case 48 To 57
    Case Else
        KeyAscii = 0
        MsgBox "Invalid character", vbInformation, "Message"
    End Select
End Sub

Private Sub txtCheckAmt_LostFocus()
    txtCheckAmt.Text = ToDoubleNumber(txtCheckAmt.Text)
End Sub

Private Sub txtCheckDate_GotFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtCheckDate_LostFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "DD-MMM-YY")
End Sub

Private Sub txtCheckNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCustCode_Change()
    txtCustomerName.Text = SetCustomerName(txtCustCode.Text)
End Sub

Private Sub txtJDate_GotFocus()
    txtJDate.Text = Format(txtJDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtJDate_LostFocus()
    txtJDate.Text = Format(txtJDate.Text, "DD-MMM-YY")
    If optDeposits.Value = 1 Then
        cboCustName.SetFocus
    End If
    If optOutstanding.Value = 1 Then
        cboVendor.SetFocus
    End If
End Sub

Private Sub txtORAmount_GotFocus()
    txtORAmount.Text = NumericVal(txtORAmount.Text)
End Sub

Private Sub txtORAmount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 8
    Case 45
    Case 46
    Case 48 To 57
    Case Else
        KeyAscii = 0
        MsgBox "Invalid character!", vbExclamation, "Message"
    End Select
End Sub

Private Sub txtORAmount_LostFocus()
    txtORAmount.Text = ToDoubleNumber(txtORAmount.Text)
End Sub

Private Sub txtORDate_GotFocus()
    txtORDate.Text = Format(txtORDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtORDate_LostFocus()
    txtORDate.Text = Format(txtORDate.Text, "DD-MMM-YY")
End Sub

Private Sub frmSearchCustomerVendor_SelectedInfo(strSelected As String)
    If xSELECTED = "Customer" Then
        txtCustCode.Text = strSelected
    ElseIf xSELECTED = "Vendor" Then
        txtVendorCode.Text = strSelected
    End If
End Sub

Private Sub txtORNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtParticulars_GotFocus()
    If txtParticulars.Text = "Type Your Message Here!" Then txtParticulars.Text = ""
End Sub

Private Sub txtParticulars2_GotFocus()
    If txtParticulars2.Text = "Type Your Message Here!" Then txtParticulars2.Text = ""
End Sub

Private Sub txtVendorCode_Change()
    txtVendorName.Text = SetVendorName(txtVendorCode.Text)
End Sub

Private Sub FillGrid(xBankCode As String)
    Dim xList                                     As ListItem
    Dim rsAMIS_Recon                              As ADODB.Recordset
    Set rsAMIS_Recon = New ADODB.Recordset
    Dim xOR_Amt                                   As Double
    Dim xCheck_Amt                                As Double
    lvTransactions.ListItems.Clear
    rsAMIS_Recon.Open "Select * from AMIS_RECONBEGINNING where BankCode ='" & SetBankCode(xBankCode) & "'", gconDMIS, adOpenKeyset
    Do While Not rsAMIS_Recon.EOF
        If Null2String(rsAMIS_Recon!OR_NUM) <> "" Or Null2String(rsAMIS_Recon!OR_NUM) <> Null Then
            Set xList = lvTransactions.ListItems.Add(, , Null2String(rsAMIS_Recon!OR_DATE))
            xList.SubItems(2) = "OR# " & Null2String(rsAMIS_Recon!OR_NUM)
        Else
            Set xList = lvTransactions.ListItems.Add(, , Null2String(rsAMIS_Recon!Check_Date))
            xList.SubItems(2) = "CHK# " & Null2String(rsAMIS_Recon!CheckNo)
        End If
        xList.SubItems(1) = Null2String(rsAMIS_Recon!remarks)
        xOR_Amt = xOR_Amt + NumericVal(rsAMIS_Recon!OR_AMT)
        xList.SubItems(3) = ToDoubleNumber(rsAMIS_Recon!OR_AMT)
        xCheck_Amt = xCheck_Amt + rsAMIS_Recon!CHECK_AMT
        xList.SubItems(4) = ToDoubleNumber(rsAMIS_Recon!CHECK_AMT)
        xList.SubItems(5) = NumericVal(rsAMIS_Recon!ID)
        rsAMIS_Recon.MoveNext
    Loop
    txtDebit.Text = ToDoubleNumber(xOR_Amt)
    txtCredit.Text = ToDoubleNumber(xCheck_Amt)
    Set rsAMIS_Recon = Nothing
End Sub

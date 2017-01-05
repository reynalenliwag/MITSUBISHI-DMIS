VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmPMISProcess_MonthEndProc 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Month-End Processing"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13440
   ControlBox      =   0   'False
   ForeColor       =   &H00FF8080&
   Icon            =   "MonthEndProc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13440
   Begin FlexCell.Grid grdTransactions 
      Height          =   6375
      Left            =   7560
      TabIndex        =   41
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   11245
      BackColor2      =   12648384
      Cols            =   6
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   2
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
      Left            =   6630
      MouseIcon       =   "MonthEndProc.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "MonthEndProc.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit Window"
      Top             =   3330
      Width           =   795
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "[ Inventory Balances ]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   7560
      TabIndex        =   11
      Top             =   6720
      Width           =   5775
      Begin VB.TextBox txtTPA 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox txtOHA 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtTRA 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtTIA 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1830
         Width           =   1215
      End
      Begin VB.TextBox txtTPM 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox txtOHM 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtTRM 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtTIM 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1830
         Width           =   1215
      End
      Begin VB.TextBox txtTI 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1830
         Width           =   1215
      End
      Begin VB.TextBox txtTR 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtOH 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtTP 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Accessories"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4380
         TabIndex        =   30
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Materials"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3090
         TabIndex        =   29
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Parts"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         TabIndex        =   28
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Issuance"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         TabIndex        =   19
         Top             =   1860
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Receipts"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         TabIndex        =   18
         Top             =   1530
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "On-Hand"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         TabIndex        =   17
         Top             =   780
         Width           =   2175
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Puchase Orders"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         TabIndex        =   16
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Frame fraCurrentActivity 
      BackColor       =   &H00FF8080&
      Caption         =   "Current Activity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   90
      TabIndex        =   9
      Top             =   4170
      Width           =   7365
      Begin RichTextLib.RichTextBox txtCurrentActivity 
         Height          =   4515
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   7964
         _Version        =   393217
         BackColor       =   8454143
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"MonthEndProc.frx":08FA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Month End Activity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      TabIndex        =   8
      Top             =   150
      Width           =   7335
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1875
         Left            =   180
         ScaleHeight     =   1875
         ScaleWidth      =   7035
         TabIndex        =   31
         Top             =   330
         Width           =   7035
         Begin VB.CheckBox chkBatchPosting 
            BackColor       =   &H00FF8080&
            Caption         =   "Batch Posting"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            TabIndex        =   36
            Top             =   390
            Width           =   3735
         End
         Begin VB.CheckBox chkMonthEnd 
            BackColor       =   &H00FF8080&
            Caption         =   "Forwarding Ending Balance as Beginning Balance for Next Cut-Off"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            TabIndex        =   35
            Top             =   750
            Width           =   6975
         End
         Begin VB.CheckBox chkGenerateRankFile 
            BackColor       =   &H00FF8080&
            Caption         =   "Generating Rank File"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            TabIndex        =   34
            Top             =   1110
            Width           =   6975
         End
         Begin VB.CheckBox chkCreateStockStatus 
            BackColor       =   &H00FF8080&
            Caption         =   "Creating Data for Stock Status Report"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            TabIndex        =   33
            Top             =   1440
            Width           =   6975
         End
         Begin VB.CheckBox chkUpdateMaster 
            BackColor       =   &H00FF8080&
            Caption         =   "Updating Master File"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            TabIndex        =   32
            Top             =   30
            Width           =   3735
         End
      End
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Process"
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
      Left            =   5850
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "MonthEndProc.frx":0978
      MousePointer    =   99  'Custom
      Picture         =   "MonthEndProc.frx":0ACA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Process Month-End Processing"
      Top             =   3330
      Width           =   795
   End
   Begin VB.PictureBox picCPB 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1335
      Left            =   60
      ScaleHeight     =   1335
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   2580
      Width           =   7455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   90
         ScaleHeight     =   405
         ScaleWidth      =   5505
         TabIndex        =   1
         Top             =   720
         Width           =   5505
         Begin VB.Label labProcessing 
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   30
            TabIndex        =   2
            ToolTipText     =   "Process progress"
            Top             =   30
            Width           =   5415
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   300
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         Picture         =   "MonthEndProc.frx":0DEF
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "MonthEndProc.frx":0E0B
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   5625
         TabIndex        =   3
         Top             =   660
         Width           =   5625
      End
      Begin VB.Label labCPB 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   7395
      End
   End
   Begin VB.PictureBox PIC_UNPOSTED 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   3630
      ScaleHeight     =   5145
      ScaleWidth      =   6405
      TabIndex        =   37
      Top             =   1290
      Visible         =   0   'False
      Width           =   6435
      Begin XtremeReportControl.ReportControl LST_UNPOSTED 
         Height          =   4425
         Left            =   90
         TabIndex        =   38
         Top             =   600
         Width           =   6225
         _Version        =   655364
         _ExtentX        =   10980
         _ExtentY        =   7805
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.CommandButton cmdT_Close 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1350
         TabIndex        =   40
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton cmdT_Print 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   39
         Top             =   90
         Width           =   1245
      End
      Begin VB.Label Label 
         Height          =   375
         Left            =   3120
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPMISProcess_MonthEndProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSTDAYTRAN, RSPARTMAS, RSSHIPPING, RSSHIPPING_COST As ADODB.Recordset
Dim rsRR_HD, RSORD_HD, RSORD_HIST                      As ADODB.Recordset
Dim RSREC_HIST, RSPO_HD, RSPO_HIST                     As ADODB.Recordset
Dim RSPO_STAT, RSDAYTRAN, rsNOHeader                   As ADODB.Recordset
Dim rsNODetail, rsNO_Mstr, rsSupplier                  As ADODB.Recordset
Dim VSUPPLIER, vVatAmt, ADDSQL, UPSQL                  As String
Dim vTDTranno, VTDPARTORD, vTDTranType, vTDType        As String
Dim VTDINOUT, VTDSTATUS                                As String
Dim vTotTranCost, vMAC                                 As Double
Dim VTDRECNO, VPMRECNO                                 As Long
Dim vPMOnhand, VPMTRECQTY, VPMTISSQTY, VPMTPOQTY       As Double
Dim VPMLAST_RECD, vTDTranDate                          As String
Dim VPMRECEIPTS, VPMISSUANCES, vTDTranQTY, VTDRRNETCOST As Double
Dim VTDNETPRICE, VTDNETCOST, VTDTRANUCOST, VTDRRINVAMT As Double
Dim VORDTOTPRICE, VTDTRANUPRICE, VTDTRANINVAMT         As Double
Dim VNETPRICE, VNETCOST                                As Double
Dim vOrdHDRecNo, VRRHDRECNO, VPOHDRECNO                As Long
Dim ictr_parts, ictr_acc, ictr_mat                     As Integer

Function GET_LESS_ERROR_PERCENTAGE_DEVIATION(aaa As Double, XXX As Double, YYY As Double, zzz As Double) As Double
    Dim DEVIATION_ERROR1                               As Double
    Dim DEVIATION_ERROR2                               As Double
    Dim DEVIATION_ERROR3                               As Double

    DEVIATION_ERROR1 = Abs(aaa - XXX)
    DEVIATION_ERROR2 = Abs(aaa - YYY)
    DEVIATION_ERROR3 = Abs(aaa - zzz)

    If DEVIATION_ERROR1 < DEVIATION_ERROR2 Then
        If DEVIATION_ERROR1 < DEVIATION_ERROR3 Then
            GET_LESS_ERROR_PERCENTAGE_DEVIATION = XXX
        Else
            If DEVIATION_ERROR2 < DEVIATION_ERROR3 Then
                GET_LESS_ERROR_PERCENTAGE_DEVIATION = zzz
            Else
                GET_LESS_ERROR_PERCENTAGE_DEVIATION = zzz
            End If
        End If
    Else
        If DEVIATION_ERROR2 < DEVIATION_ERROR3 Then
            GET_LESS_ERROR_PERCENTAGE_DEVIATION = YYY
        Else
            If DEVIATION_ERROR1 < DEVIATION_ERROR3 Then
                GET_LESS_ERROR_PERCENTAGE_DEVIATION = XXX
            Else
                GET_LESS_ERROR_PERCENTAGE_DEVIATION = zzz
            End If
        End If
    End If
End Function

Function CheckIfNonVatSup(SupplierCode As String) As Boolean
    Dim rsSupplierMaster                               As ADODB.Recordset
    Set rsSupplierMaster = New ADODB.Recordset
    rsSupplierMaster.Open "Select supcode,supname,NONVAT from PMIS_vw_Supplier where supcode = '" & SupplierCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplierMaster.EOF And Not rsSupplierMaster.BOF Then
        If Null2String(rsSupplierMaster!NONVAT) = "Y" Then CheckIfNonVatSup = True Else CheckIfNonVatSup = False
    Else
        CheckIfNonVatSup = False
    End If
End Function

Sub SetAsActiveAllStocksWithOnhand()
    Screen.MousePointer = 11
    gconDMIS.Execute ("Update PMIS_STOCKMAS SET ACTIVE = 'Y' WHERE ONHAND > 0")
    Screen.MousePointer = 0
End Sub

Function UpdateMaster() As Boolean
    On Error GoTo ErrorCode
    
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim rsCURPartmas                                   As ADODB.Recordset
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    Dim RSPO_HD                                        As ADODB.Recordset
    Dim rsRR_HD                                        As ADODB.Recordset
    Dim RSORD_HD                                       As ADODB.Recordset
    Dim i                                              As Integer
    Dim vTotTranCost                                   As Long
    Dim vTotTranInvAmt                                 As Double
    Dim vTDTranQTY                                     As Double
    Dim vTDTranDate                                    As String
    Dim vTDTranType                                    As String
    Dim vTDTranno                                      As String
    Dim vTDType                                        As String
    Dim vVatAmt, vMAC                                  As Double
    Dim vPMOnhand                                      As Long
    Dim vSTOCKDESC                                     As String
    Dim vTotalQty                                      As Long
    Dim vOrdHDRecNo                                    As Long


    '==========================================================
    'UPDATING PMIS_PO_HD VALUE FROM DETAILS
    Set RSPO_HD = New ADODB.Recordset
    RSPO_HD.Open "select id,pono,status,TYPE from PMIS_PO_Hd order by pono asc", gconDMIS
    If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
        RSPO_HD.MoveFirst: i = 0: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Computing Total Quantity of Purchases......": DoEvents
        Do While Not RSPO_HD.EOF
            vOrdHDRecNo = RSPO_HD!ID
            labProcessing.Caption = "Processing: PO #" & Null2String(RSPO_HD!Type) & "-" & Null2String(RSPO_HD!PONO)
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,status,itemno from PMIS_TdayTran where [TYPE] = '" & Null2String(RSPO_HD!Type) & "' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst: vTotalQty = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    gconDMIS.Execute "Update PMIS_TdayTran SET " & _
                        " STATUS = '" & RSPO_HD!Status & _
                        "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
                
                gconDMIS.Execute "update PMIS_PO_Hd set " & _
                    " TotalQty = " & vTotalQty & _
                    " where id = " & vOrdHDRecNo
            End If
            i = i + 1
            progCPB.Value = (i / RSPO_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            RSPO_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    Set RSPO_HD = Nothing
    Set RSTDAYTRAN = Nothing

    
    '==========================================================
    'UPDATING PMIS_RR_HD VALUE FROM DETAILS
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,rrno,status,TYPE from PMIS_RR_Hd order by rrno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst: i = 0: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Computing Total Quantity of Receiving......": DoEvents
        Do While Not rsRR_HD.EOF
            vOrdHDRecNo = rsRR_HD!ID
            labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!Type) & "-" & Null2String(rsRR_HD!RRNO)
            
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,status,itemno from PMIS_TdayTran where [TYPE] = '" & Null2String(rsRR_HD!Type) & "' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    gconDMIS.Execute "Update PMIS_TdayTran SET " & _
                        " STATUS = '" & Null2String(rsRR_HD!Status) & _
                        "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
                
                gconDMIS.Execute "update PMIS_RR_Hd set " & _
                    " TotalQty = " & vTotalQty & _
                    " where id = " & vOrdHDRecNo
            End If
            i = i + 1
            progCPB.Value = (i / rsRR_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsRR_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    Set rsRR_HD = Nothing
    Set RSTDAYTRAN = Nothing

    '====================================================
    'UPDATING PMIS_ORD_HD VALUE FROM DETAILS
    Dim vTotalTranCost                                 As Double
    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select id,trantype,tranno,status,TYPE from PMIS_Ord_Hd order by trantype,tranno asc", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveFirst: i = 0: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Computing Total Quantity of Request and Issuance......": DoEvents
        Do While Not RSORD_HD.EOF
            vOrdHDRecNo = RSORD_HD!ID
            labProcessing.Caption = "Processing: " & Null2String(RSORD_HD!Type) & "-" & Null2String(RSORD_HD!TranType) & " #" & Null2String(RSORD_HD!TRANNO): DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,tranucost,status,itemno from PMIS_TdayTran where [TYPE] = '" & Null2String(RSORD_HD!Type) & "' AND trantype = " & N2Str2Null(RSORD_HD!TranType) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0: vTotalTranCost = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    vTotalTranCost = vTotalTranCost + (N2Str2Zero(RSTDAYTRAN!TRANUCOST) * N2Str2Zero(RSTDAYTRAN!TRANQTY))
                    gconDMIS.Execute "Update PMIS_TdayTran SET " & _
                        " STATUS = '" & Null2String(RSORD_HD!Status) & _
                        "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
                
                gconDMIS.Execute "update PMIS_Ord_Hd set " & _
                    " NETCOST = " & vTotalTranCost & _
                    ", TotalQty = " & vTotalQty & _
                    " where id = " & vOrdHDRecNo
            End If
            i = i + 1
            progCPB.Value = (i / RSORD_HD.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            RSORD_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    Set RSORD_HD = Nothing
    Set RSTDAYTRAN = Nothing
    

    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,ItemNo,trantype,tranno,TYPE,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_TdayTran where trantype not in('ADB') and isnull(status,'N') in('P','B') and isnull(status,'N') not in('N','C') order by TYPE,id asc", gconDMIS
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst: Screen.MousePointer = 11
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Storing Stocks Master Beginning Balances...": DoEvents
                
        If Month(LOGDATE) = 1 Then
            gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                " lasty_oh = ISNULL(lastm_oh,0)," & _
                " lasty_mac = ISNULL(lastm_mac,0)," & _
                " lasty_oo = ISNULL(lastm_oo,0)," & _
                " onorder = ISNULL(lastm_oo,0)," & _
                " tpoqty = 0," & _
                " tissqty = 0," & _
                " trecqty = 0," & _
                " purchases = 0," & _
                " receipts = 0," & _
                " issuances = 0 " & _
                " WHERE ACTIVE = 'Y'"
                
'            gconDMIS.Execute "update PMIS_STOCKMAS set" & _
'                " lasty_oh = ISNULL(lastm_oh,0)," & _
'                " lasty_mac = ISNULL(lastm_mac,0)," & _
'                " lasty_oo = ISNULL(lastm_oo,0)," & _
'                " onhand = ISNULL(lastm_oh,0)," & _
'                " mac = ISNULL(lastm_mac,0)," & _
'                " onorder = ISNULL(lastm_oo,0)," & _
'                " tpoqty = 0," & _
'                " tissqty = 0," & _
'                " trecqty = 0," & _
'                " purchases = 0," & _
'                " receipts = 0," & _
'                " issuances = 0 " & _
'                " WHERE ACTIVE = 'Y'"
        Else
            gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                " purchases = ISNULL(purchases,0) - ISNULL(tpoqty,0)," & _
                " receipts = ISNULL(receipts,0) - ISNULL(trecqty,0)," & _
                " issuances = ISNULL(issuances,0) - ISNULL(tissqty,0)"
            gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                " onorder = ISNULL(lastm_oo,0)," & _
                " tpoqty = 0," & _
                " tissqty = 0," & _
                " trecqty = 0 WHERE ACTIVE = 'Y'"
'
'            gconDMIS.Execute "update PMIS_STOCKMAS set" & _
'                " purchases = ISNULL(purchases,0) - ISNULL(tpoqty,0)," & _
'                " receipts = ISNULL(receipts,0) - ISNULL(trecqty,0)," & _
'                " issuances = ISNULL(issuances,0) - ISNULL(tissqty,0)"
'            gconDMIS.Execute "update PMIS_STOCKMAS set" & _
'                " onhand = ISNULL(lastm_oh,0)," & _
'                " mac = ISNULL(lastm_mac,0)," & _
'                " onorder = ISNULL(lastm_oo,0)," & _
'                " tpoqty = 0," & _
'                " tissqty = 0," & _
'                " trecqty = 0 WHERE ACTIVE = 'Y'"
        End If
        
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
            " ONREQUEST = 0," & _
            " REQSERVED = 0," & _
            " REQUNSERVED = 0," & _
            " REQFILLRATE = 0," & _
            " S_ONREQUEST = 0," & _
            " S_REQSERVED = 0," & _
            " S_REQUNSERVED = 0," & _
            " S_REQFILLRATE = 0," & _
            " ORDERED = 0," & _
            " ONORDER = 0," & _
            " SERVED = 0," & _
            " UNSERVED = 0," & _
            " BACKORDER = 0," & _
            " FILLRATE = 0," & _
            " EMERGENCY_PO = 0 WHERE ACTIVE = 'Y'"
            
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
            " STOCKTYPE = 'GJ' " & _
            " WHERE (STOCKTYPE <> 'BP' AND LEFT(STOCKNO,2) <> '08')"
        
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
            " NON_HARI = 'N' " & _
            " WHERE NON_HARI IS NULL"
        
        Call CheckInventoryBalances
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Updating Stock Master from Transactions Made...": DoEvents
        i = 0
        Do While Not RSTDAYTRAN.EOF
'            If vTDType = "P" Then
'                grdTransactions.AddItem 1 & Chr(9) & "PARTS" & Chr(9) & vTDTranType & Chr(9) & vTDTranno & Chr(9) & VTDSTATUS
'            ElseIf vTDType = "M" Then
'                grdTransactions.AddItem 1 & Chr(9) & "MATERIALS" & Chr(9) & vTDTranType & Chr(9) & vTDTranno & Chr(9) & VTDSTATUS
'            Else
'                grdTransactions.AddItem 1 & Chr(9) & "ACCESSORIES" & Chr(9) & vTDTranType & Chr(9) & vTDTranno & Chr(9) & VTDSTATUS
'            End If

            gconDMIS.Execute "update PMIS_TdayTran set " & _
                " ItemNo = '" & Format(Null2String(RSTDAYTRAN!itemno), "0000") & _
                "' where ID = " & RSTDAYTRAN!ID
                
            VTDSTATUS = Null2String(RSTDAYTRAN!Status)
            vTDType = Null2String(RSTDAYTRAN![Type])
            vTDTranDate = N2Date2Null(RSTDAYTRAN!trandate)
            vTDTranType = Null2String(RSTDAYTRAN!TranType)
            vTDTranno = Null2String(RSTDAYTRAN!TRANNO)
            vTDTranQTY = N2Str2Zero(RSTDAYTRAN!TRANQTY)
            vTotTranCost = N2Str2Zero(RSTDAYTRAN!TRANUCOST) * vTDTranQTY
            vTotTranInvAmt = N2Str2Zero(RSTDAYTRAN!TRANINVAMT) * vTDTranQTY

            labProcessing.Caption = "Processing: " & Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO): DoEvents

            Set RSPARTMAS = New ADODB.Recordset
            Set RSPARTMAS = gconDMIS.Execute("select STOCKNO from PMIS_STOCKMAS where TYPE = '" & vTDType & "' AND STOCKNO = " & N2Str2Null(RSTDAYTRAN!STOCK_ORD))
            If RSPARTMAS.EOF And RSPARTMAS.BOF Then
                If vTDType = "P" Then
                    Set rsCURPartmas = New ADODB.Recordset
                    Set rsCURPartmas = gconDMIS.Execute("Select PARTNUMBER,DESCRIPTIO from PMIS_DNPP where PARTNUMBER = " & N2Str2Null(RSTDAYTRAN!STOCK_ORD))
                    If Not rsCURPartmas.EOF And Not rsCURPartmas.BOF Then
                        vSTOCKDESC = N2Str2Null(rsCURPartmas!DESCRIPTIO)
                    Else
                        vSTOCKDESC = "'NO DESCRIPTION'"
                    End If
                Else
                    vSTOCKDESC = "'NO DESCRIPTION'"
                End If
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Inserting Found New Stock No. (" & Null2String(RSTDAYTRAN!STOCK_ORD) & ")": DoEvents
                
                gconDMIS.Execute ("Insert into PMIS_STOCKMAS (TYPE,STOCKNO,STOCKDESC,date_entered) " & _
                    " values ('" & vTDType & _
                    "', " & N2Str2Null(RSTDAYTRAN!STOCK_ORD) & _
                    ", " & vSTOCKDESC & _
                    ", " & N2Str2Null(RSTDAYTRAN!trandate) & ")")
                txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Continued: Updating Stock Master from Transactions Made...": DoEvents
            Else
                gconDMIS.Execute ("Update PMIS_STOCKMAS SET " & _
                    " ACTIVE = 'Y' " & _
                    ", TYPE = '" & vTDType & _
                    "' WHERE STOCKNO = " & N2Str2Null(RSTDAYTRAN!STOCK_ORD))
            End If
            
            Set RSPARTMAS = New ADODB.Recordset
            Set RSPARTMAS = gconDMIS.Execute("select id, STOCKNO, mac, Onhand, NON_HARI from PMIS_STOCKMAS where " & _
                " TYPE = '" & vTDType & _
                "' AND STOCKNO = " & N2Str2Null(RSTDAYTRAN!STOCK_ORD))
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                vMAC = N2Str2Zero(RSPARTMAS!MAC)
                vPMOnhand = N2Str2IntZero(RSPARTMAS!ONHAND)
                
                If Null2String(RSTDAYTRAN!IN_OUT) = "R" And vTDTranQTY <> 0 Then
                    Set RSORD_HD = New ADODB.Recordset
                    Set RSORD_HD = gconDMIS.Execute("Select sales_origin from PMIS_Ord_HD where [TYPE] = '" & vTDType & "' AND trantype = '" & vTDTranType & "' and tranno = '" & Format(vTDTranno, "000000") & "'")
                    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                " ONREQUEST = ISNULL(ONREQUEST,0) + " & vTDTranQTY & _
                                " where id = " & RSPARTMAS!ID
                        Else
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                " S_ONREQUEST = ISNULL(S_ONREQUEST,0) + " & vTDTranQTY & _
                                " where id = " & RSPARTMAS!ID
                        End If
                    End If
                End If
                
                If Null2String(RSTDAYTRAN!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    Set RSORD_HD = New ADODB.Recordset
                    Set RSORD_HD = gconDMIS.Execute("Select sales_origin from PMIS_Ord_HD where [TYPE] = '" & vTDType & "' AND trantype = '" & vTDTranType & "' and tranno = '" & Format(vTDTranno, "000000") & "' and isnull(status,'N') in('P','B') and isnull(status,'N') not in ('C','N')")
                    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                " REQSERVED = ISNULL(REQSERVED,0) + " & vTDTranQTY & ", " & _
                                " tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                " issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                " where id = " & RSPARTMAS!ID
'                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
'                                " onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
'                                " REQSERVED = ISNULL(REQSERVED,0) + " & vTDTranQTY & ", " & _
'                                " tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
'                                " issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
'                                " where id = " & RSPARTMAS!ID
                        Else
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                " S_REQSERVED = ISNULL(S_REQSERVED,0) + " & vTDTranQTY & ", " & _
                                " tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                " issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                " where id = " & RSPARTMAS!ID
'
'                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
'                                " onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
'                                " S_REQSERVED = ISNULL(S_REQSERVED,0) + " & vTDTranQTY & ", " & _
'                                " tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
'                                " issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
'                                " where id = " & RSPARTMAS!ID
                        End If
                    Else
                        If vTDTranType = "ADJ" Then
                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                " tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
                                " issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
                                " where id = " & RSPARTMAS!ID
'                            Execute "update PMIS_STOCKMAS set " & _
'                                " onhand = ISNULL(onhand,0) - " & vTDTranQTY & ", " & _
'                                " tissqty = ISNULL(tissqty,0) + " & vTDTranQTY & ", " & _
'                                " issuances = ISNULL(issuances,0) + " & vTDTranQTY & _
'                                " where id = " & RSPARTMAS!ID
                        End If
                    End If
                    
                    gconDMIS.Execute "update PMIS_TdayTran set " & _
                        " NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & _
                        " where ID = " & RSTDAYTRAN!ID
                        
                    'COMMENT : REMOVE THE TRANUCOST UPDATE
                    'gconDMIS.Execute "update PMIS_TdayTran set " & _
                    '    " NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & _
                    '    ", tranucost = " & vMAC & _
                    '    " where ID = " & RSTDAYTRAN!ID
                    
                    If vTDType = "P" Then
                        txtOH.Text = Format(NumericVal(txtOH.Text) - vTDTranQTY, DIGIT_FORMAT)
                        txtTI.Text = Format(NumericVal(txtTI.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "M" Then
                        txtOHM.Text = Format(NumericVal(txtOHM.Text) - vTDTranQTY, DIGIT_FORMAT)
                        txtTIM.Text = Format(NumericVal(txtTIM.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "A" Then
                        txtOHA.Text = Format(NumericVal(txtOHA.Text) - vTDTranQTY, DIGIT_FORMAT)
                        txtTIA.Text = Format(NumericVal(txtTIA.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                End If

                If Null2String(RSTDAYTRAN!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from PMIS_RR_HD where [TYPE] = '" & vTDType & "' AND rrno = '" & Format(vTDTranno, "000000") & "' and isnull(status,'N') in ('P','B') and isnull(status,'N') not in ('C','N') ", gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        vVatAmt = N2Str2Zero(rsRR_HD!ds1)
                        If Null2String(rsRR_HD!classcode) = "PCG" Or Null2String(rsRR_HD!classcode) = "PCS" Then
                            If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = False Then
                                vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                            Else
                                vTotTranCost = vTotTranInvAmt
                                gconDMIS.Execute ("update PMIS_TdayTran Set " & _
                                    " NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & _
                                    " Where id = " & RSTDAYTRAN!ID)
                                    
                                'COMMENT BY  : MJP 07302010 0112PM
                                'DESCRIPTION : REMOVE THE TRANUCOST FIELD
                                'gconDMIS.Execute ("update PMIS_TdayTran Set " & _
                                '    " NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & _
                                '    ", tranucost = " & N2Str2Zero(RSTDAYTRAN!TRANINVAMT) & _
                                '    " Where id = " & RSTDAYTRAN!ID)
                                
                                gconDMIS.Execute ("update PMIS_RR_HD Set " & _
                                    " DS1 = 0 " & _
                                    ", DS_AMT1 = 0 " & _
                                    " WHERE ID = " & rsRR_HD!ID)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Or vTDTranType = "BEG" Then
                            vTotTranCost = vMAC * vTDTranQTY
                        End If
                    End If
                    
                    If vPMOnhand <= 0 Then
                        'UPDATE BY   : MJP 07202010 0418PM
                        'DESCRIPTION : VTOTTRANCOST IS ALREADY WITH OUT VAT, IS CHANGES THE MAC COMPUTATION
                            vMAC = (vTotTranInvAmt / 1.12) / vTDTranQTY
                        'UPDATE BY   : MJP 07202010 0418PM
                        
                        'COMMENT BY  : MJP 07202010 0418PM
                        'DESCRIPTION : VTOTTRANCOST IS ALREADY WITH OUT VAT, IS CHANGES THE MAC COMPUTATION
                            'vMAC = vTotTranCost / vTDTranQTY
                        'COMMENT BY  : MJP 07202010 0418PM
                        
                        gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                            " ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
                            " SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
                            " last_recd = " & vTDTranDate & ", " & _
                            " trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
                            " receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
                            " where id =" & RSPARTMAS!ID
                            
'                        gconDMIS.Execute "update PMIS_STOCKMAS set " & _
'                            " Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & ", " & _
'                            " ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
'                            " SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
'                            " last_recd = " & vTDTranDate & ", " & _
'                            " trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
'                            " receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
'                            " where id =" & RSPARTMAS!ID
                    Else
                        'UPDATE BY   : MJP 07202010 0418PM
                        'DESCRIPTION : VTOTTRANCOST IS ALREADY WITH OUT VAT, IS CHANGES THE MAC COMPUTATION
                            vMAC = ((vMAC * vPMOnhand) + (vTotTranInvAmt / 1.12)) / (vTDTranQTY + vPMOnhand)
                        'UPDATE BY   : MJP 07202010 0418PM
                        
                        'COMMENT BY  : MJP 07202010 0418PM
                        'DESCRIPTION : VTOTTRANCOST IS ALREADY WITH OUT VAT, IS CHANGES THE MAC COMPUTATION
                            'vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                        'COMMENT BY  : MJP 07202010 0418PM
                        
                        gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                            " ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
                            " SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
                            " last_recd = " & vTDTranDate & ", " & _
                            " trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
                            " receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
                            " where id =" & RSPARTMAS!ID
'                        gconDMIS.Execute "update PMIS_STOCKMAS set " & _
'                            " ONORDER = ISNULL(ONORDER,0) - " & vTDTranQTY & ", " & _
'                            " SERVED = ISNULL(SERVED,0) + " & vTDTranQTY & ", " & _
'                            " Onhand = ISNULL(Onhand,0) + " & vTDTranQTY & ", " & _
'                            " last_recd = " & vTDTranDate & ", " & _
'                            " trecqty = ISNULL(trecqty,0) + " & vTDTranQTY & ", " & _
'                            " receipts = ISNULL(receipts,0) + " & vTDTranQTY & _
'                            " wher                    "
                    End If
                    
                    'gconDMIS.Execute "update PMIS_TdayTran set " & _
                    '    " NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & _
                    '    ", MAC = " & vMAC & _
                    '    " where id = " & RSTDAYTRAN!ID
                    gconDMIS.Execute "update PMIS_TdayTran set " & _
                        " NON_HARI = " & N2Str2Null(RSPARTMAS!NON_HARI) & _
                        " where id = " & RSTDAYTRAN!ID
                    
                    If vTDType = "P" Then
                        txtOH.Text = Format(NumericVal(txtOH.Text) + vTDTranQTY, DIGIT_FORMAT)
                        txtTR.Text = Format(NumericVal(txtTR.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "M" Then
                        txtOHM.Text = Format(NumericVal(txtOHM.Text) + vTDTranQTY, DIGIT_FORMAT)
                        txtTRM.Text = Format(NumericVal(txtTRM.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "A" Then
                        txtOHA.Text = Format(NumericVal(txtOHA.Text) + vTDTranQTY, DIGIT_FORMAT)
                        txtTRA.Text = Format(NumericVal(txtTRA.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                End If
                
                If Null2String(RSTDAYTRAN!TranType) = "PO" Then
                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                        " purchases = ISNULL(purchases,0) + " & vTDTranQTY & "," & _
                        " tpoqty = ISNULL(tpoqty,0) + " & vTDTranQTY & "," & _
                        " ONORDER = ISNULL(ONORDER,0) + " & vTDTranQTY & "," & _
                        " ORDERED = ISNULL(ORDERED,0) + " & vTDTranQTY & _
                        " where id = " & RSPARTMAS!ID
                    If vTDType = "P" Then
                        txtTP.Text = Format(NumericVal(txtTP.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "M" Then
                        txtTPM.Text = Format(NumericVal(txtTPM.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "A" Then
                        txtTPA.Text = Format(NumericVal(txtTPA.Text) + vTDTranQTY, DIGIT_FORMAT)
                    End If
                End If
            End If

            If vTDType = "P" Then
                grdTransactions.AddItem 1 & Chr(9) & "PARTS" & Chr(9) & vTDTranType & Chr(9) & vTDTranno & Chr(9) & VTDSTATUS
            ElseIf vTDType = "M" Then
                grdTransactions.AddItem 1 & Chr(9) & "MATERIALS" & Chr(9) & vTDTranType & Chr(9) & vTDTranno & Chr(9) & VTDSTATUS
            Else
                grdTransactions.AddItem 1 & Chr(9) & "ACCESSORIES" & Chr(9) & vTDTranType & Chr(9) & vTDTranno & Chr(9) & VTDSTATUS
            End If
            
            i = i + 1
            progCPB.Value = (i / RSTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            RSTDAYTRAN.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents
        Screen.MousePointer = 0
    End If
    
    UpdateMaster = True
    Exit Function
    
ErrorCode:
    UpdateMaster = False
End Function
Private Sub RANKING_REPORT()
    '***************11032010***************************************************
    'For Ranking Reports of Parts
    'Updating code:     jbuzted
    Dim jCommand             As ADODB.Command
    Dim rsSM                 As ADODB.Recordset
    Set rsSM = New ADODB.Recordset
    Set rsSM = gconDMIS.Execute("SELECT STOCKNO FROM PMIS_PARTMAS WHERE " & _
               "STOCKNO IN (SELECT STOCK_ORD FROM PMIS_TDAYTRAN WHERE TYPE = 'P' AND " & _
               "TRANTYPE IN ('CSH','CHG','DR','RIV') AND (STATUS = 'P' OR STATUS = 'B'))")
    
    If Not rsSM.EOF And Not rsSM.BOF Then
        Do While Not rsSM.EOF
            Set jCommand = New ADODB.Command
            jCommand.NamedParameters = True
            jCommand.CommandType = adCmdStoredProc
            jCommand.CommandText = "RANKING_REPORT"
            jCommand.ActiveConnection = gconDMIS
               With jCommand
                   .Parameters.Append jCommand.CreateParameter("@STOCKNO", adVarChar, adParamInput, 30, rsSM!STOCKNO)
                   .Parameters.Append jCommand.CreateParameter("@DATE_GEN", adDBDate, adParamInput, , LOGDATE)
                   .Execute
               End With
            rsSM.MoveNext
        Loop
    End If
    '**************************************************************************
    Set rsSM = Nothing
    End Sub
Function BatchPosting() As Boolean
    
    'RANKING_REPORT
    Dim i                                              As Long
    Set RSTDAYTRAN = New ADODB.Recordset
'    RSTDAYTRAN.Open "Select TYPE,id,in_out,trantype,tranno,STOCK_ORD,status,tranqty,netcost,tranucost,trandate,tranuprice,traninvamt from PMIS_TdayTran where (TRANTYPE <> 'PRS' OR TRANTYPE <> 'MRS' OR TRANTYPE <> 'ARS') AND (status ='P' OR STATUS='B') order by TYPE,id asc", gconDMIS
    RSTDAYTRAN.Open "Select TYPE,id,in_out,trantype,tranno,STOCK_ORD,status,tranqty,netcost,tranucost,trandate,tranuprice,traninvamt from PMIS_TdayTran where (TRANTYPE <> 'PRS' OR TRANTYPE <> 'MRS' OR TRANTYPE <> 'ARS') AND (isnull(status,'N') ='P' OR isnull(status,'N')='B') and (isnull(status,'N') <>'C' OR isnull(status,'N')<>'N') order by TYPE,id asc", gconDMIS
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst: i = 0: Screen.MousePointer = 11:
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Posting Parts Transactions from Daily Transactions File...": DoEvents
        Do While Not RSTDAYTRAN.EOF
            '================================================================================================================================================================
            'STORE PMIS_TDAYTRAN VALUES
            VTDRECNO = RSTDAYTRAN!ID
            vTDType = Null2String(RSTDAYTRAN!Type)
            VTDINOUT = Null2String(RSTDAYTRAN!IN_OUT)
            vTDTranType = Null2String(RSTDAYTRAN!TranType)
            vTDTranno = Null2String(RSTDAYTRAN!TRANNO)
            VTDPARTORD = Null2String(RSTDAYTRAN!STOCK_ORD)
            VTDSTATUS = Null2String(RSTDAYTRAN!Status)
            vTDTranQTY = N2Str2Zero(RSTDAYTRAN!TRANQTY)
            vTDTranDate = Null2Date(RSTDAYTRAN!trandate)

            VTDNETCOST = Round(N2Str2Zero(RSTDAYTRAN!netcost), 2)
            VTDTRANUCOST = Round(N2Str2Zero(RSTDAYTRAN!TRANUCOST), 2)
            VTDTRANINVAMT = Round(N2Str2Zero(RSTDAYTRAN!TRANINVAMT), 2)
            vTotTranCost = Round(VTDTRANUCOST * vTDTranQTY, 2)
            VTDTRANUPRICE = Round(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), 2)
            '================================================================================================================================================================

            labProcessing.Caption = "Processing: " & vTDTranType & " #" & vTDTranno: DoEvents

            '================================================================================================================================================================
            'PROCESS DATA IF STOCK ORDER EXIST IN STOCK MASTER FILE
            Set RSPARTMAS = New ADODB.Recordset
            Set RSPARTMAS = gconDMIS.Execute("Select id, onhand, trecqty, last_recd, receipts, tissqty, issuances, tpoqty, lastm_MAC, MAC from PMIS_STOCKMAS where [TYPE] = '" & vTDType & "' AND STOCKNO = '" & VTDPARTORD & "'")
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                VPMRECNO = RSPARTMAS!ID
                vMAC = Round(N2Str2Zero(RSPARTMAS!MAC), 2)

                '================================================================================================================================================================
                'CHECK IF TRANSACTIONS ARE RECEIVING OR ISSUANCE
                If vTDTranType <> "ADJ" And vTDTranType <> "PO" And (VTDINOUT = "I" Or VTDINOUT = "O") And vTDTranQTY <> 0 Then
                    '================================================================================================================================================================
                    'IF TRANSACTIONS IS PARTS RECEIVING

                    If vTDTranType = "RR" Then
                        Set rsRR_HD = New ADODB.Recordset
                        Set rsRR_HD = gconDMIS.Execute("Select ID,recvd_code,ds1,status,classcode,rrno from PMIS_RR_Hd where [TYPE] = '" & vTDType & "' AND rrno = '" & Format(RSTDAYTRAN!TRANNO, "000000") & "'")
                        If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                            VSUPPLIER = Null2String(rsRR_HD!recvd_code)
                            If rsRR_HD!classcode = "PCG" Or rsRR_HD!classcode = "PCS" Then
                                If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = False Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                                Else
                                    vTotTranCost = VTDTRANINVAMT * vTDTranQTY
                                End If
                            End If

                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                " trecqty = ISNULL(trecqty,0) - " & vTDTranQTY & ", " & _
                                " last_recd = " & N2Str2Null(vTDTranDate) & _
                                " where id =" & VPMRECNO

                            If vTDType = "P" Then
                                txtTR.Text = Format(NumericVal(txtTR.Text) - vTDTranQTY, DIGIT_FORMAT)
                            End If
                            If vTDType = "M" Then
                                txtTRM.Text = Format(NumericVal(txtTRM.Text) - vTDTranQTY, DIGIT_FORMAT)
                            End If
                            If vTDType = "A" Then
                                txtTRA.Text = Format(NumericVal(txtTRA.Text) - vTDTranQTY, DIGIT_FORMAT)
                            End If

                            gconDMIS.Execute "update PMIS_TdayTran set " & _
                                " status = 'P' " & _
                                " where id = " & VTDRECNO
                            
                            gconDMIS.Execute "Update PMIS_RR_Hd Set " & _
                                " Status = 'P' " & _
                                " where id = " & rsRR_HD!ID
                        Else
                            gconDMIS.Execute "insert into PMIS_NoHeader " & _
                                " (TYPE, trantype, tranno, recno, stat_h)" & _
                                " values ('" & vTDType & "','RR', '" & vTDTranno & "', " & VTDRECNO & ", '" & VTDSTATUS & "')"
                        End If
                        Set rsRR_HD = Nothing

                        Set RSSHIPPING = New ADODB.Recordset
                        Set RSSHIPPING = gconDMIS.Execute("select * from PMIS_Shipping WHERE [TYPE] = '" & vTDType & "' and Partno = '" & VTDPARTORD & "'")
                        If RSSHIPPING.EOF And RSSHIPPING.BOF Then
                            gconDMIS.Execute "insert into PMIS_Shipping ([TYPE],partno) values ('" & vTDType & "','" & VTDPARTORD & "')"
                        End If
                        Set RSSHIPPING = Nothing

                        Set RSSHIPPING = New ADODB.Recordset
                        Set RSSHIPPING = gconDMIS.Execute("select * from PMIS_Shipping_Cost where Partno = '" & VTDPARTORD & "'")
                        If RSSHIPPING.EOF And RSSHIPPING.BOF Then
                            gconDMIS.Execute "insert into PMIS_Shipping_Cost ([TYPE],partno) values ('P','" & VTDPARTORD & "')"
                        End If
                        Set RSSHIPPING = Nothing
                    End If
                    '================================================================================================================================================================

                    '================================================================================================================================================================
                    'IF TRANSACTIONS IS PARTS ISSUANCE
                    If VTDINOUT = "O" Then
                        Set RSORD_HD = New ADODB.Recordset
                        Set RSORD_HD = gconDMIS.Execute("Select trantype,tranno from PMIS_Ord_Hd where [TYPE] = '" & vTDType & "' AND trantype = '" & vTDTranType & "' and tranno = '" & Format(RSTDAYTRAN!TRANNO, "000000") & "'")
                        If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                            If vTDTranType = "CHG" Or vTDTranType = "CSH" Or vTDTranType = "RIV" Then
                                VORDTOTPRICE = (VTDTRANUPRICE * vTDTranQTY) / ConvertToBIRDecimalFormat(VAT_RATE)
                            Else
                                VORDTOTPRICE = (VTDTRANUPRICE * vTDTranQTY)
                            End If

                            gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                                " tissqty = ISNULL(tissqty,0) - " & vTDTranQTY & _
                                " where id = " & VPMRECNO

                            If vTDType = "P" Then
                                txtTI.Text = Format(NumericVal(txtTI.Text) - vTDTranQTY, DIGIT_FORMAT)
                            End If
                            If vTDType = "M" Then
                                txtTIM.Text = Format(NumericVal(txtTIM.Text) - vTDTranQTY, DIGIT_FORMAT)
                            End If
                            If vTDType = "A" Then
                                txtTIA.Text = Format(NumericVal(txtTIA.Text) - vTDTranQTY, DIGIT_FORMAT)
                            End If

                            gconDMIS.Execute "update PMIS_TdayTran set " & _
                                " netcost = " & vTotTranCost & _
                                ", netprice = " & VORDTOTPRICE & _
                                ", status = 'P' " & _
                                " where id = " & VTDRECNO

                            Set RSSHIPPING = New ADODB.Recordset
                            Set RSSHIPPING = gconDMIS.Execute("select * from PMIS_Shipping WHERE [TYPE] = '" & vTDType & "' and Partno = '" & VTDPARTORD & "'")
                            If Not RSSHIPPING.EOF And Not RSSHIPPING.BOF Then
                                gconDMIS.Execute "update PMIS_Shipping set " & _
                                    " curr_month = ISNULL(curr_month,0) + " & vTDTranQTY & ", " & _
                                    " freq_curr = 1 " & _
                                    " where id = " & RSSHIPPING!ID
                            Else
                                gconDMIS.Execute "insert into PMIS_Shipping ([TYPE],partno,curr_month,freq_curr)" & _
                                    " values ('" & vTDType & _
                                    "', '" & VTDPARTORD & _
                                    "', " & vTDTranQTY & _
                                    ", 1)"
                            End If
                            Set RSSHIPPING = Nothing

                            Set RSSHIPPING = New ADODB.Recordset
                            Set RSSHIPPING = gconDMIS.Execute("select * from PMIS_Shipping_Cost where Partno = '" & VTDPARTORD & "'")
                            If Not RSSHIPPING.EOF And Not RSSHIPPING.BOF Then
                                gconDMIS.Execute "update PMIS_Shipping_Cost set " & _
                                    " curr_month = ISNULL(curr_month,0) + " & vTotTranCost & _
                                    ", freq_curr = 1 " & _
                                    " where id = " & RSSHIPPING!ID
                            Else
                                gconDMIS.Execute "insert into PMIS_Shipping_Cost ([TYPE],partno,curr_month,freq_curr) values " & _
                                    " ('P' " & _
                                    ", '" & VTDPARTORD & _
                                    "', " & vTotTranCost & _
                                    ", 1)"
                            End If
                        Else
                            gconDMIS.Execute "insert into PMIS_NoHeader " & _
                                "(TYPE, trantype, tranno, recno, stat_h) values " & _
                                "('" & vTDType & _
                                "', '" & vTDTranType & _
                                "', '" & vTDTranno & _
                                "', " & VTDRECNO & _
                                ", '" & VTDSTATUS & "')"
                        End If
                    End If
                    '================================================================================================================================================================
                End If
                '================================================================================================================================================================

                '================================================================================================================================================================
                'PROCESS IF TRANSACTION IS ADJUSTMENT IN
                If vTDTranType = "ADJ" And VTDINOUT = "I" And vTDTranQTY <> 0 Then
                    vTotTranCost = N2Str2Zero(RSPARTMAS!MAC) * vTDTranQTY
                    
                    'UPDATE BY   : MJP07202010 0520PM
                    'DESCRIPTION : REMOVE THE FIELD TRANUCOST
                        gconDMIS.Execute "update PMIS_TdayTran set " & _
                             " netcost = " & vTotTranCost & _
                            " where id = " & VTDRECNO
                    'UPDATE BY   : MJP07202010 0520PM
                    
                    'UPDATE BY   : MJP07202010 0520PM
                    'DESCRIPTION : REMOVE THE FIELD TRANUCOST
                    'KUTO?? - BAKIT NAG UPUPDATE DITO NG TRANUCOST
'                        gconDMIS.Execute "update PMIS_TdayTran set " & _
'                            " tranucost = " & vMAC & "," & _
'                            " netcost = " & vTotTranCost & _
'                            " where id = " & VTDRECNO
                    'UPDATE BY   : MJP07202010 0520PM
                    
                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                        " trecqty = ISNULL(trecqty,0) - " & vTDTranQTY & ", " & _
                        " last_recd = " & N2Str2Null(vTDTranDate) & _
                        " where id =" & VPMRECNO

                    If vTDType = "P" Then
                        txtTR.Text = Format(NumericVal(txtTR.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "M" Then
                        txtTRM.Text = Format(NumericVal(txtTRM.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "A" Then
                        txtTRA.Text = Format(NumericVal(txtTRA.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                End If
                '================================================================================================================================================================

                '================================================================================================================================================================
                'PROCESS IF TRANSACTION IS ADJUSTMENT OUT
                If vTDTranType = "ADJ" And VTDINOUT = "O" And vTDTranQTY <> 0 Then
                    vTotTranCost = vMAC * vTDTranQTY
                    VORDTOTPRICE = (vMAC * vTDTranQTY)
                    VPMTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
                    
                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                        " tissqty = ISNULL(tissqty,0) - " & vTDTranQTY & _
                        " where id = " & VPMRECNO
                        
                    'UPDATE BY   : MJP07202010 0520PM
                    'DESCRIPTION : REMOVE THE FIELD TRANUCOST
                        gconDMIS.Execute "update PMIS_TdayTran set " & _
                            " netcost = " & vTotTranCost & _
                            ", netprice = " & VORDTOTPRICE & _
                            ", status = 'P' where id = " & VTDRECNO
                    'UPDATE BY   : MJP07202010 0520PM
                    
                    'COMMENT BY  : MJP07202010 0520PM
                    'DESCRIPTION : CHANGE THE COST IN TRANSACTION COST
                        'KUTO?? - BAKIT NAG UPUPDATE DITO NG TRANUCOST
                        'gconDMIS.Execute "update PMIS_TdayTran set " & _
                        '    " tranucost = " & vMAC & _
                        '    ", netcost = " & vTotTranCost & _
                        '    ", netprice = " & VORDTOTPRICE & _
                        '    ", status = 'P' where id = " & VTDRECNO
                    'COMMENT BY  : MJP07202010 0520PM

                        
                    If vTDType = "P" Then
                        txtTI.Text = Format(NumericVal(txtTI.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "M" Then
                        txtTIM.Text = Format(NumericVal(txtTIM.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "A" Then
                        txtTIA.Text = Format(NumericVal(txtTIA.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                End If
                '================================================================================================================================================================

                '================================================================================================================================================================
                'PROCESS IF TRANSACTION IS PURCHASING
                If vTDTranType = "PO" And vTDTranQTY <> 0 And (VTDSTATUS <> "C" Or VTDSTATUS <> "N") Then
                    gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                        "tpoqty = ISNULL(tpoqty,0) - " & vTDTranQTY & _
                        " where id = " & VPMRECNO
                    If vTDType = "P" Then
                        txtTP.Text = Format(NumericVal(txtTP.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "M" Then
                        txtTPM.Text = Format(NumericVal(txtTPM.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                    If vTDType = "A" Then
                        txtTPA.Text = Format(NumericVal(txtTPA.Text) - vTDTranQTY, DIGIT_FORMAT)
                    End If
                End If
                '================================================================================================================================================================

            Else
                If vTDTranType <> "ADB" Then
                    gconDMIS.Execute "insert into PMIS_No_Mstr " & _
                        "(TYPE, trantype, tranno, recno)" & _
                        " values ('" & vTDType & _
                        "', '" & VTDINOUT & _
                        "', '" & vTDTranno & _
                        "', " & VTDRECNO & ")"
                End If
            End If
            
            i = i + 1
            progCPB.Value = (i / RSTDAYTRAN.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            RSTDAYTRAN.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents: Screen.MousePointer = 0
    End If
    Set RSTDAYTRAN = Nothing

    '======================================================
    'UPDATING PMIS_ORD_HD VALUE FROM DETAIL VALUE
    Set RSORD_HD = New ADODB.Recordset
    Set RSORD_HD = gconDMIS.Execute("select id,[TYPE],trantype,tranno,status from PMIS_Ord_Hd order by trantype,tranno asc")
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveFirst: i = 0
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Transfering Issuance Transactions in History File...": DoEvents
        Screen.MousePointer = 11: DoEvents
        
        Do While Not RSORD_HD.EOF
            vOrdHDRecNo = RSORD_HD!ID
            labProcessing.Caption = "Processing: " & Null2String(RSORD_HD!Type) & Null2String(RSORD_HD!TranType) & " #" & Null2String(RSORD_HD!TRANNO): DoEvents

            Set RSTDAYTRAN = New ADODB.Recordset
            Set RSTDAYTRAN = gconDMIS.Execute("select id,trantype,tranno,netprice,netcost,status,itemno from PMIS_TdayTran where [TYPE] = '" & Null2String(RSORD_HD![Type]) & "' AND trantype = " & N2Str2Null(RSORD_HD!TranType) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc")
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst: VNETPRICE = 0: VNETCOST = 0
                Do While Not RSTDAYTRAN.EOF
                    VTDNETPRICE = N2Str2Zero(RSTDAYTRAN!NETprice)
                    VTDNETCOST = N2Str2Zero(RSTDAYTRAN!netcost)
                    VTDSTATUS = Null2String(RSTDAYTRAN!Status)

                    If VTDSTATUS <> "C" And VTDSTATUS <> "N" Then
                        VNETPRICE = VNETPRICE + VTDNETPRICE
                        VNETCOST = VNETCOST + VTDNETCOST
                    Else
                        VNETPRICE = 0: VNETCOST = 0
                    End If

                    'SYSTEM WILL CANCEL IF NOT POSTED
                    If Null2String(RSTDAYTRAN!Status) = "N" Then
                        gconDMIS.Execute "update PMIS_TDAYTRAN set " & _
                            " status = 'C' " & _
                            " where id = " & RSTDAYTRAN!ID
                    Else
                        gconDMIS.Execute "update PMIS_TDAYTRAN set " & _
                            " status = " & N2Str2Null(RSTDAYTRAN!Status) & _
                            " where id = " & RSTDAYTRAN!ID
                    End If
                    RSTDAYTRAN.MoveNext
                Loop
                
                'SYSTEM WILL CANCEL IF NOT POSTED
                If Null2String(RSORD_HD!Status) = "N" Then
                    gconDMIS.Execute "update PMIS_Ord_Hd set " & _
                        " netcost2 = " & VNETCOST & _
                        ", netinvamt2 = " & VNETPRICE & _
                        ", status = 'C' " & _
                        " where id = " & vOrdHDRecNo
                Else
                    gconDMIS.Execute "update PMIS_Ord_Hd set " & _
                        " netcost2 = " & VNETCOST & _
                        ", netinvamt2 = " & VNETPRICE & _
                        ", status = " & N2Str2Null(RSORD_HD!Status) & _
                        " where id = " & vOrdHDRecNo
                End If
            End If
            MoveOrdHd (vOrdHDRecNo)
            i = i + 1: progCPB.Value = (i / RSORD_HD.RecordCount) * 100: labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents: RSORD_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents: Screen.MousePointer = 0
    End If

    '============================================================
    'UPDATING PMIS_RR_HD VALUE FROM DETAILS VALUE
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select [TYPE],id,rrno,status,classcode,recvd_code,ds1 from PMIS_RR_Hd order by rrno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst: i = 0: Screen.MousePointer = 11:
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Transfering Receiving Transactions in History File...": DoEvents
        Do While Not rsRR_HD.EOF
            VRRHDRECNO = rsRR_HD!ID: labProcessing.Caption = "Processing: RR #" & Null2String(rsRR_HD!RRNO): DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            Set RSTDAYTRAN = gconDMIS.Execute("select id,status,tranqty,trantype,tranno,itemno,tranucost,mac,traninvamt from PMIS_TdayTran where [TYPE] = '" & Null2String(rsRR_HD![Type]) & "' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc")
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst: VTDRRINVAMT = 0
                Do While Not RSTDAYTRAN.EOF
                    VTDRECNO = RSTDAYTRAN!ID: VTDSTATUS = Null2String(RSTDAYTRAN!Status)
                    VTDRRINVAMT = VTDRRINVAMT + (N2Str2Zero(RSTDAYTRAN!TRANINVAMT) * N2Str2Zero(RSTDAYTRAN!TRANQTY))

                    'SYSTEM WILL CANCEL IF NOT POSTED
                    If Null2String(RSTDAYTRAN!Status) = "N" Then
                        gconDMIS.Execute "update PMIS_TDAYTRAN set " & _
                            " status = 'C' " & _
                            " where id = " & RSTDAYTRAN!ID
                    Else
                        gconDMIS.Execute "update PMIS_TDAYTRAN set " & _
                            " status = " & N2Str2Null(RSTDAYTRAN!Status) & _
                            " where id = " & RSTDAYTRAN!ID
                    End If
                    RSTDAYTRAN.MoveNext
                Loop
                
                'SYSTEM WILL CANCEL IF NOT POSTED
                If Null2String(rsRR_HD!Status) = "N" And (rsRR_HD!classcode = "PCG" Or rsRR_HD!classcode = "PCS") Then
                    If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = True Then
                        gconDMIS.Execute "update PMIS_RR_Hd set " & _
                            " ttlrramt = " & VTDRRINVAMT & _
                            ", netcost = " & VTDRRNETCOST & _
                            ", status = 'C' " & _
                            " where id = " & VRRHDRECNO
                    Else
                        gconDMIS.Execute "update PMIS_RR_Hd set " & _
                            " ttlrramt = " & VTDRRINVAMT / ConvertToBIRDecimalFormat(N2Str2Zero(rsRR_HD!ds1)) & _
                            ", ds_amt1 = " & VTDRRINVAMT - (VTDRRINVAMT / ConvertToBIRDecimalFormat(N2Str2Zero(rsRR_HD!ds1))) & _
                            ", netrramt = " & VTDRRINVAMT & _
                            ", netcost = " & VTDRRNETCOST & _
                            ", status = 'C' " & _
                            " where id = " & VRRHDRECNO
                    End If
                Else
                    If CheckIfNonVatSup(Null2String(rsRR_HD!recvd_code)) = True Then
                        gconDMIS.Execute "update PMIS_RR_Hd set " & _
                            " ttlrramt = " & VTDRRINVAMT & _
                            ", netcost = " & VTDRRNETCOST & _
                            ", status = '" & Null2String(rsRR_HD!Status) & _
                            "' where id = " & VRRHDRECNO
                    Else
                        gconDMIS.Execute "update PMIS_RR_Hd set " & _
                            " ttlrramt = " & VTDRRINVAMT / ConvertToBIRDecimalFormat(N2Str2Zero(rsRR_HD!ds1)) & _
                            ", ds_amt1 = " & VTDRRINVAMT - (VTDRRINVAMT / ConvertToBIRDecimalFormat(N2Str2Zero(rsRR_HD!ds1))) & _
                            ", netrramt = " & VTDRRINVAMT & _
                            ", netcost = " & VTDRRNETCOST & _
                            ", status = '" & Null2String(rsRR_HD!Status) & _
                            "' where id = " & VRRHDRECNO
                    End If
                End If
            End If
            MoveRRhd (VRRHDRECNO)
            i = i + 1: progCPB.Value = (i / rsRR_HD.RecordCount) * 100: labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents: rsRR_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents: Screen.MousePointer = 0
    End If

    '=====================================================
    'UPDATE PMIS_PO_HD VALUE FROM DETAILS VALUE
    Set RSPO_HD = New ADODB.Recordset
    RSPO_HD.Open "select [TYPE],id,pono,status from PMIS_PO_Hd order by pono asc", gconDMIS
    If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
        RSPO_HD.MoveFirst: i = 0: Screen.MousePointer = 11:
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Transfering Purchase Transactions in History File...": DoEvents
        Do While Not RSPO_HD.EOF
            VPOHDRECNO = RSPO_HD!ID: labProcessing.Caption = "Processing: PO #" & Null2String(RSPO_HD!PONO): DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            Set RSTDAYTRAN = gconDMIS.Execute("select id,status,trantype,tranno,itemno from PMIS_TdayTran where [TYPE] = '" & Null2String(RSPO_HD![Type]) & "' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc")
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                Do While Not RSTDAYTRAN.EOF
                    VTDRECNO = RSTDAYTRAN!ID: VTDSTATUS = Null2String(RSTDAYTRAN!Status)
                    'SYSTEM WILL CANCEL IF NOT POSTED
                    If Null2String(RSPO_HD!Status) = "N" Then
                        gconDMIS.Execute "update PMIS_PO_Hd set " & _
                            " status = 'C' " & _
                            " where id = " & VPOHDRECNO
                    Else
                        gconDMIS.Execute "update PMIS_PO_Hd set " & _
                            " status = '" & Null2String(RSPO_HD!Status) & _
                            "' where id = " & VPOHDRECNO
                    End If
                    RSTDAYTRAN.MoveNext
                Loop
            End If
            MovePOhd (VPOHDRECNO)
            i = i + 1: progCPB.Value = (i / RSPO_HD.RecordCount) * 100: labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents: RSPO_HD.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents: Screen.MousePointer = 0
    End If


    'MOVING TRANSACTION FILE TO HISTORY
    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Transfering Transactions Details in History File...": DoEvents
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,status,trantype,tranno,itemno from PMIS_TdayTran where (trantype <> 'ADJ' AND trantype <> 'BEG') order by id asc", gconDMIS
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst: labProcessing.Caption = "Processing: " & Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO): DoEvents
        Do While Not RSTDAYTRAN.EOF
            VTDRECNO = RSTDAYTRAN!ID: VTDSTATUS = Null2String(RSTDAYTRAN!Status)
            'SYSTEM WILL CANCEL IF NOT POSTED
            If VTDSTATUS = "N" Then
                gconDMIS.Execute "update PMIS_TdayTran set " & _
                    " status = 'C' " & _
                    " where id = " & VTDRECNO
            Else
                gconDMIS.Execute "update PMIS_TdayTran set " & _
                    " status = '" & VTDSTATUS & _
                    "' where id = " & VTDRECNO
            End If
            
            MoveTdaytran (RSTDAYTRAN!ID)
            RSTDAYTRAN.MoveNext
        Loop
    End If
    
    'MOVING ADJUSTMENT TRANSACTION TO HISTORY
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,status,trantype,tranno,itemno from PMIS_TdayTran where trantype = 'ADJ' order by id asc", gconDMIS
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst: labProcessing.Caption = "Processing: ADJ #" & Null2String(RSTDAYTRAN!TRANNO): DoEvents
        Do While Not RSTDAYTRAN.EOF
            VTDRECNO = RSTDAYTRAN!ID: VTDSTATUS = Null2String(RSTDAYTRAN!Status)
            'SYSTEM WILL CANCEL IF NOT POSTED
            If VTDSTATUS = "N" Then
                gconDMIS.Execute "update PMIS_TdayTran set " & _
                    " status = 'C' " & _
                    " where id =" & VTDRECNO
            Else
                gconDMIS.Execute "update PMIS_TdayTran set " & _
                    " status = '" & VTDSTATUS & _
                    "' where id =" & VTDRECNO
            End If
            
            MoveTdaytran (RSTDAYTRAN!ID)
            RSTDAYTRAN.MoveNext
        Loop
    End If

    'MOVING BEG TO HISTORY FILE
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,status,trantype,tranno,itemno from PMIS_TdayTran where trantype = 'BEG' order by id asc", gconDMIS
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst: labProcessing.Caption = "Processing: BEG #" & Null2String(RSTDAYTRAN!TRANNO): DoEvents
        Do While Not RSTDAYTRAN.EOF
            VTDRECNO = RSTDAYTRAN!ID: VTDSTATUS = Null2String(RSTDAYTRAN!Status)
            'SYSTEM WILL CANCEL IF NOT POSTED
            If VTDSTATUS = "N" Then
                gconDMIS.Execute "update PMIS_TdayTran set " & _
                    " status = 'C' " & _
                    " where id =" & VTDRECNO
            Else
                gconDMIS.Execute "update PMIS_TdayTran set " & _
                    " status = '" & VTDSTATUS & _
                    "' where id =" & VTDRECNO
            End If
            
            MoveTdaytran (RSTDAYTRAN!ID)
            RSTDAYTRAN.MoveNext
        Loop
    End If
    
    MsgSpeech "Posting of Parts Transactions Completed...": cmdPost.Enabled = False
    Set RSTDAYTRAN = Nothing: Set RSPARTMAS = Nothing: Set RSSHIPPING = Nothing: Set RSORD_HD = Nothing: Set rsRR_HD = Nothing: Set RSPO_HD = Nothing
    BatchPosting = True
    
    Exit Function
ErrorCode:
    BatchPosting = False
End Function

Function MonthEndUpdate() As Boolean
    On Error GoTo errocode:
    Screen.MousePointer = 11
    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Forwarding Ending Balance as Beginning Balance for Next Cut-Off...": DoEvents
    labCPB.Caption = "Updating Part Master File... Please Wait..."

    If Month(LOGDATE) = 1 Then
        'UPDATE PARTS
        gconDMIS.Execute ("Update PMIS_PARTMAS SET PMIS_PARTMAS.NOSHIP = 0 FROM PMIS_SHIPPING WHERE ISNULL(PMIS_Shipping.Curr_Month,0) > 0 and PMIS_Shipping.PARTNO = PMIS_PARTMAS.STOCKNO")
        gconDMIS.Execute ("Update PMIS_PARTMAS SET PMIS_PARTMAS.NOSHIP = ISNULL(PMIS_PARTMAS.NOSHIP,0) + 1 FROM PMIS_SHIPPING WHERE (ISNULL(PMIS_Shipping.Curr_Month,0) <= 0 OR PMIS_Shipping.Curr_Month IS NULL) and PMIS_Shipping.PARTNO = PMIS_PARTMAS.STOCKNO")
        gconDMIS.Execute "update PMIS_PARTMAS set" & _
                       " PMIS_PARTMAS.lastm_oh = ISNULL(PMIS_PARTMAS.onhand,0)," & _
                       " PMIS_PARTMAS.lastm_mac = ISNULL(PMIS_PARTMAS.Mac,0)," & _
                       " PMIS_PARTMAS.lastm_mad = ISNULL(PMIS_PARTMAS.Mad,0)," & _
                       " PMIS_PARTMAS.lastm_oo = ISNULL(PMIS_PARTMAS.onorder,0)" & _
                       " where PMIS_PARTMAS.ACTIVE = 'Y'"

        'UPDATE MATERIALS
        gconDMIS.Execute ("Update PMIS_STOCKMAS SET PMIS_STOCKMAS.NOSHIP = 0 FROM PMIS_SHIPPING WHERE ISNULL(PMIS_Shipping.Curr_Month,0) > 0 and PMIS_Shipping.PARTNO = PMIS_STOCKMAS.STOCKNO AND PMIS_STOCKMAS.TYPE = 'M'")
        gconDMIS.Execute ("Update PMIS_STOCKMAS SET PMIS_STOCKMAS.NOSHIP = ISNULL(PMIS_STOCKMAS.NOSHIP,0) + 1 FROM PMIS_SHIPPING WHERE (ISNULL(PMIS_Shipping.Curr_Month,0) <= 0 OR PMIS_Shipping.Curr_Month IS NULL) and PMIS_Shipping.PARTNO = PMIS_STOCKMAS.STOCKNO AND PMIS_STOCKMAS.TYPE = 'M'")
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " PMIS_STOCKMAS.lastm_oh = ISNULL(PMIS_STOCKMAS.onhand,0)," & _
                       " PMIS_STOCKMAS.lastm_mac = ISNULL(PMIS_STOCKMAS.Mac,0)," & _
                       " PMIS_STOCKMAS.lastm_mad = ISNULL(PMIS_STOCKMAS.Mad,0)" & _
                       " where PMIS_STOCKMAS.TYPE = 'M' AND PMIS_STOCKMAS.ACTIVE = 'Y'"

        'UPDATE ACCESSORIES
        gconDMIS.Execute ("Update PMIS_STOCKMAS SET PMIS_STOCKMAS.NOSHIP = 0 FROM PMIS_SHIPPING WHERE ISNULL(PMIS_Shipping.Curr_Month,0) > 0 and PMIS_Shipping.PARTNO = PMIS_STOCKMAS.STOCKNO AND PMIS_STOCKMAS.TYPE = 'A'")
        gconDMIS.Execute ("Update PMIS_STOCKMAS SET PMIS_STOCKMAS.NOSHIP = ISNULL(PMIS_STOCKMAS.NOSHIP,0) + 1 FROM PMIS_SHIPPING WHERE (ISNULL(PMIS_Shipping.Curr_Month,0) <= 0 OR PMIS_Shipping.CURR_MONTH IS NULL) and PMIS_Shipping.PARTNO = PMIS_STOCKMAS.STOCKNO AND PMIS_STOCKMAS.TYPE = 'A'")
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " PMIS_STOCKMAS.lastm_oh = ISNULL(PMIS_STOCKMAS.onhand,0)," & _
                       " PMIS_STOCKMAS.lastm_mac = ISNULL(PMIS_STOCKMAS.Mac,0)," & _
                       " PMIS_STOCKMAS.lastm_mad = ISNULL(PMIS_STOCKMAS.Mad,0)," & _
                       " PMIS_STOCKMAS.lastm_oo = ISNULL(PMIS_STOCKMAS.onorder,0)," & _
                       " PMIS_STOCKMAS.lasty_mac = ISNULL(PMIS_STOCKMAS.Mac,0)," & _
                       " PMIS_STOCKMAS.lasty_oo = ISNULL(PMIS_STOCKMAS.onorder,0)" & _
                       " where PMIS_STOCKMAS.TYPE = 'A' AND PMIS_STOCKMAS.ACTIVE = 'Y'"
    Else
        'UPDATE PARTS
        gconDMIS.Execute ("Update PMIS_PARTMAS SET PMIS_PARTMAS.NOSHIP = 0 FROM PMIS_SHIPPING WHERE ISNULL(PMIS_Shipping.Curr_Month,0) > 0 and PMIS_Shipping.PARTNO = PMIS_PARTMAS.STOCKNO")
        gconDMIS.Execute ("Update PMIS_PARTMAS SET PMIS_PARTMAS.NOSHIP = ISNULL(PMIS_PARTMAS.NOSHIP,0) + 1 FROM PMIS_SHIPPING WHERE (ISNULL(PMIS_Shipping.Curr_Month,0) <= 0 OR PMIS_Shipping.CURR_MONTH IS NULL) and PMIS_Shipping.PARTNO = PMIS_PARTMAS.STOCKNO")
        gconDMIS.Execute "update PMIS_PARTMAS set" & _
                       " PMIS_PARTMAS.lastm_oh = ISNULL(PMIS_PARTMAS.onhand,0)," & _
                       " PMIS_PARTMAS.lastm_mac = ISNULL(PMIS_PARTMAS.Mac,0)," & _
                       " PMIS_PARTMAS.lastm_mad = ISNULL(PMIS_PARTMAS.Mad,0)," & _
                       " PMIS_PARTMAS.lastm_oo = ISNULL(PMIS_PARTMAS.onorder,0)" & _
                       " where PMIS_PARTMAS.ACTIVE = 'Y'"

        'UPDATE MATERIALS
        gconDMIS.Execute ("Update PMIS_STOCKMAS SET PMIS_STOCKMAS.NOSHIP = 0 FROM PMIS_SHIPPING WHERE PMIS_Shipping.Curr_Month > 0 and PMIS_Shipping.PARTNO = PMIS_STOCKMAS.STOCKNO AND PMIS_STOCKMAS.TYPE = 'M'")
        gconDMIS.Execute ("Update PMIS_STOCKMAS SET PMIS_STOCKMAS.NOSHIP = ISNULL(PMIS_STOCKMAS.NOSHIP,0) + 1 FROM PMIS_SHIPPING WHERE (ISNULL(PMIS_Shipping.Curr_Month,0) <= 0 OR PMIS_Shipping.CURR_MONTH IS NULL) and PMIS_Shipping.PARTNO = PMIS_STOCKMAS.STOCKNO AND PMIS_STOCKMAS.TYPE = 'M'")
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " PMIS_STOCKMAS.lastm_oh = ISNULL(PMIS_STOCKMAS.onhand,0)," & _
                       " PMIS_STOCKMAS.lastm_mac = ISNULL(PMIS_STOCKMAS.Mac,0)," & _
                       " PMIS_STOCKMAS.lastm_mad = ISNULL(PMIS_STOCKMAS.Mad,0)," & _
                       " PMIS_STOCKMAS.lastm_oo = ISNULL(PMIS_STOCKMAS.onorder,0)" & _
                       " where PMIS_STOCKMAS.TYPE = 'M' AND PMIS_STOCKMAS.ACTIVE = 'Y'"

        'UPDATE ACCESSORIES
        gconDMIS.Execute ("Update PMIS_STOCKMAS SET PMIS_STOCKMAS.NOSHIP = 0 FROM PMIS_SHIPPING WHERE ISNULL(PMIS_Shipping.Curr_Month,0) > 0 and PMIS_Shipping.PARTNO = PMIS_STOCKMAS.STOCKNO AND PMIS_STOCKMAS.TYPE = 'A'")
        gconDMIS.Execute ("Update PMIS_STOCKMAS SET PMIS_STOCKMAS.NOSHIP = ISNULL(PMIS_STOCKMAS.NOSHIP,0) + 1 FROM PMIS_SHIPPING WHERE (ISNULL(PMIS_Shipping.Curr_Month,0) <= 0 OR PMIS_Shipping.CURR_MONTH IS NULL) and PMIS_Shipping.PARTNO = PMIS_STOCKMAS.STOCKNO AND PMIS_STOCKMAS.TYPE = 'A'")
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " PMIS_STOCKMAS.lastm_oh = ISNULL(PMIS_STOCKMAS.onhand,0)," & _
                       " PMIS_STOCKMAS.lastm_mac = ISNULL(PMIS_STOCKMAS.Mac,0)," & _
                       " PMIS_STOCKMAS.lastm_mad = ISNULL(PMIS_STOCKMAS.Mad,0)," & _
                       " PMIS_STOCKMAS.lastm_oo = ISNULL(PMIS_STOCKMAS.onorder,0)" & _
                       " where PMIS_STOCKMAS.TYPE = 'A' AND PMIS_STOCKMAS.ACTIVE = 'Y'"

    End If
    Screen.MousePointer = 11

    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Forwarding Shipping Quantity as Beginning Balance for Next Cut-Off...": DoEvents
    labCPB.Caption = "Updating Shipping File... Please Wait...": DoEvents
    'UPDATE PARTS
    gconDMIS.Execute "update PMIS_Shipping set" & _
        " months_60 = ISNULL(Months_59,0), months_59 = ISNULL(Months_58,0), months_58 = ISNULL(Months_57,0), months_57 = ISNULL(Months_56,0)," & _
        " months_56 = ISNULL(Months_55,0), months_55 = ISNULL(Months_54,0), months_54 = ISNULL(Months_53,0), months_53 = ISNULL(Months_52,0)," & _
        " months_52 = ISNULL(Months_51,0), months_51 = ISNULL(Months_50,0), months_50 = ISNULL(Months_49,0), months_49 = ISNULL(Months_48,0)," & _
        " months_48 = ISNULL(Months_47,0), months_47 = ISNULL(Months_46,0), months_46 = ISNULL(Months_45,0), months_45 = ISNULL(Months_44,0)," & _
        " months_44 = ISNULL(Months_43,0), months_43 = ISNULL(Months_42,0), months_42 = ISNULL(Months_41,0), months_41 = ISNULL(Months_40,0)," & _
        " months_40 = ISNULL(Months_39,0), months_39 = ISNULL(Months_38,0), months_38 = ISNULL(Months_37,0), months_37 = ISNULL(Months_36,0)," & _
        " months_36 = ISNULL(Months_35,0), months_35 = ISNULL(Months_34,0), months_34 = ISNULL(Months_33,0), months_33 = ISNULL(Months_32,0)," & _
        " months_32 = ISNULL(Months_31,0), months_31 = ISNULL(Months_30,0), months_30 = ISNULL(Months_29,0), months_29 = ISNULL(Months_28,0)," & _
        " months_28 = ISNULL(Months_27,0), months_27 = ISNULL(Months_26,0), months_26 = ISNULL(Months_25,0), months_25 = ISNULL(Months_24,0)," & _
        " months_24 = ISNULL(Months_23,0), months_23 = ISNULL(Months_22,0), months_22 = ISNULL(Months_21,0), months_21 = ISNULL(Months_20,0)," & _
        " months_20 = ISNULL(Months_19,0), months_19 = ISNULL(Months_18,0), months_18 = ISNULL(Months_17,0), months_17 = ISNULL(Months_16,0)," & _
        " months_16 = ISNULL(Months_15,0), months_15 = ISNULL(Months_14,0), months_14 = ISNULL(Months_13,0), months_13 = ISNULL(Months_12,0)," & _
        " months_12 = ISNULL(Months_11,0), months_11 = ISNULL(Months_10,0), months_10 = ISNULL(Months_9,0), months_9 = ISNULL(Months_8,0)," & _
        " months_8 = ISNULL(Months_7,0), months_7 = ISNULL(Months_6,0), months_6 = ISNULL(Months_5,0), months_5 = ISNULL(Months_4,0)," & _
        " months_4 = ISNULL(Months_3,0), months_3 = ISNULL(Months_2,0), months_2 = ISNULL(Prev_Month,0), prev_month = ISNULL(Curr_Month,0)," & _
        " curr_month = 0 WHERE [TYPE] = 'P'"

    'UPDATE MATERIALS
    gconDMIS.Execute "update PMIS_Shipping set" & _
        " months_60 = ISNULL(Months_59,0), months_59 = ISNULL(Months_58,0), months_58 = ISNULL(Months_57,0), months_57 = ISNULL(Months_56,0)," & _
        " months_56 = ISNULL(Months_55,0), months_55 = ISNULL(Months_54,0), months_54 = ISNULL(Months_53,0), months_53 = ISNULL(Months_52,0)," & _
        " months_52 = ISNULL(Months_51,0), months_51 = ISNULL(Months_50,0), months_50 = ISNULL(Months_49,0), months_49 = ISNULL(Months_48,0)," & _
        " months_48 = ISNULL(Months_47,0), months_47 = ISNULL(Months_46,0), months_46 = ISNULL(Months_45,0), months_45 = ISNULL(Months_44,0)," & _
        " months_44 = ISNULL(Months_43,0), months_43 = ISNULL(Months_42,0), months_42 = ISNULL(Months_41,0), months_41 = ISNULL(Months_40,0)," & _
        " months_40 = ISNULL(Months_39,0), months_39 = ISNULL(Months_38,0), months_38 = ISNULL(Months_37,0), months_37 = ISNULL(Months_36,0)," & _
        " months_36 = ISNULL(Months_35,0), months_35 = ISNULL(Months_34,0), months_34 = ISNULL(Months_33,0), months_33 = ISNULL(Months_32,0)," & _
        " months_32 = ISNULL(Months_31,0), months_31 = ISNULL(Months_30,0), months_30 = ISNULL(Months_29,0), months_29 = ISNULL(Months_28,0)," & _
        " months_28 = ISNULL(Months_27,0), months_27 = ISNULL(Months_26,0), months_26 = ISNULL(Months_25,0), months_25 = ISNULL(Months_24,0)," & _
        " months_24 = ISNULL(Months_23,0), months_23 = ISNULL(Months_22,0), months_22 = ISNULL(Months_21,0), months_21 = ISNULL(Months_20,0)," & _
        " months_20 = ISNULL(Months_19,0), months_19 = ISNULL(Months_18,0), months_18 = ISNULL(Months_17,0), months_17 = ISNULL(Months_16,0)," & _
        " months_16 = ISNULL(Months_15,0), months_15 = ISNULL(Months_14,0), months_14 = ISNULL(Months_13,0), months_13 = ISNULL(Months_12,0)," & _
        " months_12 = ISNULL(Months_11,0), months_11 = ISNULL(Months_10,0), months_10 = ISNULL(Months_9,0), months_9 = ISNULL(Months_8,0)," & _
        " months_8 = ISNULL(Months_7,0), months_7 = ISNULL(Months_6,0), months_6 = ISNULL(Months_5,0), months_5 = ISNULL(Months_4,0)," & _
        " months_4 = ISNULL(Months_3,0), months_3 = ISNULL(Months_2,0), months_2 = ISNULL(Prev_Month,0), prev_month = ISNULL(Curr_Month,0)," & _
        " curr_month = 0 WHERE [TYPE] = 'M'"

    'UPDATE ACCESSORIES
    gconDMIS.Execute "update PMIS_Shipping set" & _
        " months_60 = ISNULL(Months_59,0), months_59 = ISNULL(Months_58,0), months_58 = ISNULL(Months_57,0), months_57 = ISNULL(Months_56,0)," & _
        " months_56 = ISNULL(Months_55,0), months_55 = ISNULL(Months_54,0), months_54 = ISNULL(Months_53,0), months_53 = ISNULL(Months_52,0)," & _
        " months_52 = ISNULL(Months_51,0), months_51 = ISNULL(Months_50,0), months_50 = ISNULL(Months_49,0), months_49 = ISNULL(Months_48,0)," & _
        " months_48 = ISNULL(Months_47,0), months_47 = ISNULL(Months_46,0), months_46 = ISNULL(Months_45,0), months_45 = ISNULL(Months_44,0)," & _
        " months_44 = ISNULL(Months_43,0), months_43 = ISNULL(Months_42,0), months_42 = ISNULL(Months_41,0), months_41 = ISNULL(Months_40,0)," & _
        " months_40 = ISNULL(Months_39,0), months_39 = ISNULL(Months_38,0), months_38 = ISNULL(Months_37,0), months_37 = ISNULL(Months_36,0)," & _
        " months_36 = ISNULL(Months_35,0), months_35 = ISNULL(Months_34,0), months_34 = ISNULL(Months_33,0), months_33 = ISNULL(Months_32,0)," & _
        " months_32 = ISNULL(Months_31,0), months_31 = ISNULL(Months_30,0), months_30 = ISNULL(Months_29,0), months_29 = ISNULL(Months_28,0)," & _
        " months_28 = ISNULL(Months_27,0), months_27 = ISNULL(Months_26,0), months_26 = ISNULL(Months_25,0), months_25 = ISNULL(Months_24,0)," & _
        " months_24 = ISNULL(Months_23,0), months_23 = ISNULL(Months_22,0), months_22 = ISNULL(Months_21,0), months_21 = ISNULL(Months_20,0)," & _
        " months_20 = ISNULL(Months_19,0), months_19 = ISNULL(Months_18,0), months_18 = ISNULL(Months_17,0), months_17 = ISNULL(Months_16,0)," & _
        " months_16 = ISNULL(Months_15,0), months_15 = ISNULL(Months_14,0), months_14 = ISNULL(Months_13,0), months_13 = ISNULL(Months_12,0)," & _
        " months_12 = ISNULL(Months_11,0), months_11 = ISNULL(Months_10,0), months_10 = ISNULL(Months_9,0), months_9 = ISNULL(Months_8,0)," & _
        " months_8 = ISNULL(Months_7,0), months_7 = ISNULL(Months_6,0), months_6 = ISNULL(Months_5,0), months_5 = ISNULL(Months_4,0)," & _
        " months_4 = ISNULL(Months_3,0), months_3 = ISNULL(Months_2,0), months_2 = ISNULL(Prev_Month,0), prev_month = ISNULL(Curr_Month,0)," & _
        " curr_month = 0 WHERE [TYPE] = 'A'"


    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Forwarding Shipping Cost as Beginning Balance for Next Cut-Off...": DoEvents
    labCPB.Caption = "Updating Shipping Cost File... Please Wait...": DoEvents
    gconDMIS.Execute "update PMIS_Shipping_Cost set" & _
        " months_60 = ISNULL(Months_59,0), months_59 = ISNULL(Months_58,0), months_58 = ISNULL(Months_57,0), months_57 = ISNULL(Months_56,0)," & _
        " months_56 = ISNULL(Months_55,0), months_55 = ISNULL(Months_54,0), months_54 = ISNULL(Months_53,0), months_53 = ISNULL(Months_52,0)," & _
        " months_52 = ISNULL(Months_51,0), months_51 = ISNULL(Months_50,0), months_50 = ISNULL(Months_49,0), months_49 = ISNULL(Months_48,0)," & _
        " months_48 = ISNULL(Months_47,0), months_47 = ISNULL(Months_46,0), months_46 = ISNULL(Months_45,0), months_45 = ISNULL(Months_44,0)," & _
        " months_44 = ISNULL(Months_43,0), months_43 = ISNULL(Months_42,0), months_42 = ISNULL(Months_41,0), months_41 = ISNULL(Months_40,0)," & _
        " months_40 = ISNULL(Months_39,0), months_39 = ISNULL(Months_38,0), months_38 = ISNULL(Months_37,0), months_37 = ISNULL(Months_36,0)," & _
        " months_36 = ISNULL(Months_35,0), months_35 = ISNULL(Months_34,0), months_34 = ISNULL(Months_33,0), months_33 = ISNULL(Months_32,0)," & _
        " months_32 = ISNULL(Months_31,0), months_31 = ISNULL(Months_30,0), months_30 = ISNULL(Months_29,0), months_29 = ISNULL(Months_28,0)," & _
        " months_28 = ISNULL(Months_27,0), months_27 = ISNULL(Months_26,0), months_26 = ISNULL(Months_25,0), months_25 = ISNULL(Months_24,0)," & _
        " months_24 = ISNULL(Months_23,0), months_23 = ISNULL(Months_22,0), months_22 = ISNULL(Months_21,0), months_21 = ISNULL(Months_20,0)," & _
        " months_20 = ISNULL(Months_19,0), months_19 = ISNULL(Months_18,0), months_18 = ISNULL(Months_17,0), months_17 = ISNULL(Months_16,0)," & _
        " months_16 = ISNULL(Months_15,0), months_15 = ISNULL(Months_14,0), months_14 = ISNULL(Months_13,0), months_13 = ISNULL(Months_12,0)," & _
        " months_12 = ISNULL(Months_11,0), months_11 = ISNULL(Months_10,0), months_10 = ISNULL(Months_9,0), months_9 = ISNULL(Months_8,0)," & _
        " months_8 = ISNULL(Months_7,0), months_7 = ISNULL(Months_6,0), months_6 = ISNULL(Months_5,0), months_5 = ISNULL(Months_4,0)," & _
        " months_4 = ISNULL(Months_3,0), months_3 = ISNULL(Months_2,0), months_2 = ISNULL(Prev_Month,0), prev_month = ISNULL(Curr_Month,0)," & _
        " curr_month = 0 WHERE [TYPE] = 'P'"

    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Forwarding Forecasted Quantity as Beginning Balance for Next Cut-Off...": DoEvents
    labCPB.Caption = "Updating Forecasting File... Please Wait...": DoEvents
    gconDMIS.Execute "update PMIS_Forecast_Qty set" & _
        " months_60 = ISNULL(Months_59,0), months_59 = ISNULL(Months_58,0), months_58 = ISNULL(Months_57,0), months_57 = ISNULL(Months_56,0)," & _
        " months_56 = ISNULL(Months_55,0), months_55 = ISNULL(Months_54,0), months_54 = ISNULL(Months_53,0), months_53 = ISNULL(Months_52,0)," & _
        " months_52 = ISNULL(Months_51,0), months_51 = ISNULL(Months_50,0), months_50 = ISNULL(Months_49,0), months_49 = ISNULL(Months_48,0)," & _
        " months_48 = ISNULL(Months_47,0), months_47 = ISNULL(Months_46,0), months_46 = ISNULL(Months_45,0), months_45 = ISNULL(Months_44,0)," & _
        " months_44 = ISNULL(Months_43,0), months_43 = ISNULL(Months_42,0), months_42 = ISNULL(Months_41,0), months_41 = ISNULL(Months_40,0)," & _
        " months_40 = ISNULL(Months_39,0), months_39 = ISNULL(Months_38,0), months_38 = ISNULL(Months_37,0), months_37 = ISNULL(Months_36,0)," & _
        " months_36 = ISNULL(Months_35,0), months_35 = ISNULL(Months_34,0), months_34 = ISNULL(Months_33,0), months_33 = ISNULL(Months_32,0)," & _
        " months_32 = ISNULL(Months_31,0), months_31 = ISNULL(Months_30,0), months_30 = ISNULL(Months_29,0), months_29 = ISNULL(Months_28,0)," & _
        " months_28 = ISNULL(Months_27,0), months_27 = ISNULL(Months_26,0), months_26 = ISNULL(Months_25,0), months_25 = ISNULL(Months_24,0)," & _
        " months_24 = ISNULL(Months_23,0), months_23 = ISNULL(Months_22,0), months_22 = ISNULL(Months_21,0), months_21 = ISNULL(Months_20,0)," & _
        " months_20 = ISNULL(Months_19,0), months_19 = ISNULL(Months_18,0), months_18 = ISNULL(Months_17,0), months_17 = ISNULL(Months_16,0)," & _
        " months_16 = ISNULL(Months_15,0), months_15 = ISNULL(Months_14,0), months_14 = ISNULL(Months_13,0), months_13 = ISNULL(Months_12,0)," & _
        " months_12 = ISNULL(Months_11,0), months_11 = ISNULL(Months_10,0), months_10 = ISNULL(Months_9,0), months_9 = ISNULL(Months_8,0)," & _
        " months_8 = ISNULL(Months_7,0), months_7 = ISNULL(Months_6,0), months_6 = ISNULL(Months_5,0), months_5 = ISNULL(Months_4,0)," & _
        " months_4 = ISNULL(Months_3,0), months_3 = ISNULL(Months_2,0), months_2 = ISNULL(Prev_Month,0), prev_month = ISNULL(Curr_Month,0)," & _
        " curr_month = 0"

    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Forwarding Level of Service Quantity as Beginning Balance for Next Cut-Off...": DoEvents
    labCPB.Caption = "Updating Level of Service File... Please Wait...": DoEvents
    gconDMIS.Execute "update PMIS_Level_Of_Service set" & _
        " months_60 = ISNULL(Months_59,0), months_59 = ISNULL(Months_58,0), months_58 = ISNULL(Months_57,0), months_57 = ISNULL(Months_56,0)," & _
        " months_56 = ISNULL(Months_55,0), months_55 = ISNULL(Months_54,0), months_54 = ISNULL(Months_53,0), months_53 = ISNULL(Months_52,0)," & _
        " months_52 = ISNULL(Months_51,0), months_51 = ISNULL(Months_50,0), months_50 = ISNULL(Months_49,0), months_49 = ISNULL(Months_48,0)," & _
        " months_48 = ISNULL(Months_47,0), months_47 = ISNULL(Months_46,0), months_46 = ISNULL(Months_45,0), months_45 = ISNULL(Months_44,0)," & _
        " months_44 = ISNULL(Months_43,0), months_43 = ISNULL(Months_42,0), months_42 = ISNULL(Months_41,0), months_41 = ISNULL(Months_40,0)," & _
        " months_40 = ISNULL(Months_39,0), months_39 = ISNULL(Months_38,0), months_38 = ISNULL(Months_37,0), months_37 = ISNULL(Months_36,0)," & _
        " months_36 = ISNULL(Months_35,0), months_35 = ISNULL(Months_34,0), months_34 = ISNULL(Months_33,0), months_33 = ISNULL(Months_32,0)," & _
        " months_32 = ISNULL(Months_31,0), months_31 = ISNULL(Months_30,0), months_30 = ISNULL(Months_29,0), months_29 = ISNULL(Months_28,0)," & _
        " months_28 = ISNULL(Months_27,0), months_27 = ISNULL(Months_26,0), months_26 = ISNULL(Months_25,0), months_25 = ISNULL(Months_24,0)," & _
        " months_24 = ISNULL(Months_23,0), months_23 = ISNULL(Months_22,0), months_22 = ISNULL(Months_21,0), months_21 = ISNULL(Months_20,0)," & _
        " months_20 = ISNULL(Months_19,0), months_19 = ISNULL(Months_18,0), months_18 = ISNULL(Months_17,0), months_17 = ISNULL(Months_16,0)," & _
        " months_16 = ISNULL(Months_15,0), months_15 = ISNULL(Months_14,0), months_14 = ISNULL(Months_13,0), months_13 = ISNULL(Months_12,0)," & _
        " months_12 = ISNULL(Months_11,0), months_11 = ISNULL(Months_10,0), months_10 = ISNULL(Months_9,0), months_9 = ISNULL(Months_8,0)," & _
        " months_8 = ISNULL(Months_7,0), months_7 = ISNULL(Months_6,0), months_6 = ISNULL(Months_5,0), months_5 = ISNULL(Months_4,0)," & _
        " months_4 = ISNULL(Months_3,0), months_3 = ISNULL(Months_2,0), months_2 = ISNULL(Prev_Month,0), prev_month = ISNULL(Curr_Month,0)," & _
        " curr_month = 0": DoEvents

    Screen.MousePointer = 0: Me.Caption = "Updating Complete!": labCPB.Caption = "Updating Complete!": MsgSpeech "Month End Processing Completed!"
    MonthEndUpdate = True
    Exit Function
errocode:
    MonthEndUpdate = False
End Function

Function GenRankFile() As Boolean
    On Error GoTo ErrorCode
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim rsForeCast_Qty                                 As ADODB.Recordset
    Dim rsLevel_Of_Service                             As ADODB.Recordset
    Dim i                                              As Long

    Dim SMonths_12, SMonths_11, SMonths_10             As Long
    Dim SMonths_9, SMonths_8, SMonths_7                As Long
    Dim SMonths_6, SMonths_5, SMonths_4                As Long
    Dim SMonths_3, SMonths_2, SPrev_Month              As Long
    Dim vTotSales, vMAD12                              As Double

    Dim SMonths_12_Qty, SMonths_11_Qty, SMonths_10_Qty As Long
    Dim SMonths_9_Qty, SMonths_8_Qty, SMonths_7_Qty    As Long
    Dim SMonths_6_Qty, SMonths_5_Qty, SMonths_4_Qty    As Long
    Dim SMonths_3_Qty, SMonths_2_Qty, SPrev_Month_Qty  As Long

    Dim SMonths_12_Cost, SMonths_11_Cost, SMonths_10_Cost As Double
    Dim SMonths_9_Cost, SMonths_8_Cost, SMonths_7_Cost As Double
    Dim SMonths_6_Cost, SMonths_5_Cost, SMonths_4_Cost As Double
    Dim SMonths_3_Cost, SMonths_2_Cost, SPrev_Month_Cost As Double
    Dim vTotSales_Cost, vMAD12_Cost                    As Double

    Dim vRankType, vSubClass                           As String
    Dim vPrevClass, vPrevSClass                        As String
    Dim Number_Of_Years_No_Sale                        As Long
    Dim S_year1, S_year2, S_year3, S_year4, S_year5    As Long
    Dim P_Onhand                                       As Long
    Dim P_MAC                                          As Double
    Dim P_Last_recd, P_STOCKDESC                       As String

    Dim A_SALES, B_SALES, C_SALES                      As Double
    Dim RANK_COST                                      As Double
    Dim LOS                                            As Double

    Dim SIX_MONTHS_SALE                                As Double
    Dim MIDDLE_OF_SIX_MONTHS                           As Double

    Dim MOVING_AVERAGE_SIX_MONTHS                      As Double
    Dim MOVING_MEDIAN                                  As Double
    Dim LINEAR_REGRESSION                              As Double

    Dim Y_REGRESSION                                   As Double
    Dim X_REGRESSION                                   As Double
    Dim MEAN_OF_X                                      As Double
    Dim MEAN_OF_Y                                      As Double

    Const N_VALUE = 6
    Dim SUMMATION_OF_X                                 As Double
    Dim SUMMATION_OF_Y                                 As Double
    Dim SLOPE_OF_THE_LINE                              As Double
    Dim INTERCEPT_VALUE_AT_ZERO                        As Double
    Dim CURRENT_DEMAND                                 As Double
    Dim NEARES_DEMAND_FORECAST                         As Double
    Dim MEAN_ABSOULTE_DEVIATION                        As Double
    Dim SUGGESTED_ORDER_QTY                            As Double
    Dim PIYA                                           As Long

    Dim ELEVEN_MONTHS_FORECAST                         As Double

    Dim TOTAL_DEMAND                                   As Double
    Dim EMERGENCY_PURCHASES                            As Double
    Dim LOST_SALES                                     As Double
    Dim LEVEL_OF_SERVICE                               As Double

    Dim SAFETY_FACTOR                                  As Double
    Dim LEAD_TIME                                      As Double
    Dim ORDER_FREQUENCY                                As Double

    Dim ONHAND                                         As Double
    Dim ON_ORDER                                       As Double
    Dim BACK_ORDER                                     As Double

    Dim SAFETY_STOCK                                   As Double

    Dim VarPartNo                                      As String
    Dim varRECORDDATE                                  As String
    Dim varTRANYPE                                     As String
    Dim varTRANNO                                      As String

    Dim varC_REQUESTED                                 As Double
    Dim varC_SERVED                                    As Double
    Dim varC_UNSERVED                                  As Double
    Dim varC_BACKORDER                                 As Double
    Dim varC_FILLRATE                                  As Double

    Dim varS_REQUESTED                                 As Double
    Dim varS_SERVED                                    As Double
    Dim varS_UNSERVED                                  As Double
    Dim varS_BACKORDER                                 As Double
    Dim varS_FILLRATE                                  As Double

    Dim varD_ORDERED                                   As Double
    Dim varD_SERVED                                    As Double
    Dim varD_UNSERVED                                  As Double
    Dim varD_BACKORDER                                 As Double
    Dim varD_FILLRATE                                  As Double
    Dim varD_ONORDER                                   As Double
    Dim varD_EMERGENCY_PO                              As Double
    Dim LVAL()                                         As Double

    Dim rsDemand_Monitoring                            As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "select STOCKNO,STOCKDESC,TYPE,onhand,mac,last_recd,invclass,subinvclas,DATE_ENTERED,EMERGENCY_PO,LOST_SALES,ONORDER,BACKORDER from PMIS_STOCKMAS WHERE ACTIVE = 'Y'  order by STOCKNO asc", gconDMIS
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        RSPARTMAS.MoveFirst
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Generating Ranking For Stocks...": DoEvents
        txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "This will take few minutes...": DoEvents
        Me.Caption = "Generating Rank File": DoEvents: i = 0
        Do While Not RSPARTMAS.EOF
            labProcessing.Caption = "TYPE = " & Null2String(RSPARTMAS!Type) & " - Stock No: " & Null2String(RSPARTMAS!STOCKNO)
            DoEvents
            SMonths_12 = 0: SMonths_11 = 0: SMonths_10 = 0: SMonths_9 = 0
            SMonths_8 = 0: SMonths_7 = 0: SMonths_6 = 0: SMonths_5 = 0
            SMonths_4 = 0: SMonths_3 = 0: SMonths_2 = 0: SPrev_Month = 0
            vTotSales = 0: vMAD12 = 0

            SMonths_12_Cost = 0: SMonths_11_Cost = 0: SMonths_10_Cost = 0: SMonths_9_Cost = 0
            SMonths_8_Cost = 0: SMonths_7_Cost = 0: SMonths_6_Cost = 0: SMonths_5_Cost = 0
            SMonths_4_Cost = 0: SMonths_3_Cost = 0: SMonths_2_Cost = 0: SPrev_Month_Cost = 0
            vTotSales_Cost = 0: vMAD12_Cost = 0

            S_year1 = 0: S_year2 = 0: S_year3 = 0: S_year4 = 0: S_year5 = 0
            Number_Of_Years_No_Sale = 0
            P_Onhand = N2Str2Zero(RSPARTMAS!ONHAND)
            P_MAC = N2Str2Zero(RSPARTMAS!MAC)
            P_Last_recd = N2Date2Null(RSPARTMAS!LAST_RECD)
            P_STOCKDESC = N2Str2Null(RSPARTMAS!STOCKDESC)
            vPrevClass = N2Str2Null(RSPARTMAS!InvClass)
            vPrevSClass = N2Str2Null(RSPARTMAS!SubInvClas)
            Set RSSHIPPING = New ADODB.Recordset
            RSSHIPPING.Open "Select * from PMIS_Shipping where [TYPE] = " & N2Str2Null(RSPARTMAS!Type) & " AND PARTNO = " & N2Str2Null(RSPARTMAS!STOCKNO), gconDMIS
            If Not RSSHIPPING.EOF And Not RSSHIPPING.BOF Then
  
  'updated by: IEBV    1252011_05pm
  'description: new formula for ranking report
  '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                   If COMPANY_CODE = "HAI" Then
                        If RANKING_A = True Then
                            vRankType = RANK_FAST_MOVING
                            GoTo con
                        ElseIf RANKING_B = True Then
                            vRankType = RANK_MEDIUM_MOVING
                            GoTo con
                        ElseIf (N2Str2Zero(RSSHIPPING!Prev_Month) + N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                            vRankType = RANK_SLOW_MOVING
                            GoTo con
                        ElseIf (N2Str2Zero(RSSHIPPING!Prev_Month) + N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12) + N2Str2Zero(RSSHIPPING!months_13) + N2Str2Zero(RSSHIPPING!months_14) + N2Str2Zero(RSSHIPPING!months_15) + N2Str2Zero(RSSHIPPING!months_16) + N2Str2Zero(RSSHIPPING!months_17) + N2Str2Zero(RSSHIPPING!months_18) + N2Str2Zero(RSSHIPPING!months_19) + N2Str2Zero(RSSHIPPING!months_20) + N2Str2Zero(RSSHIPPING!months_21) + N2Str2Zero(RSSHIPPING!months_22) + N2Str2Zero(RSSHIPPING!months_23) + N2Str2Zero(RSSHIPPING!months_24)) > 0 Then
                            vRankType = RANK_NON_MOVING
                            GoTo con
                        Else
                            vRankType = RANK_NEW_PARTS
                            GoTo con
                        End If
                    Else
                        If Null2Date(RSPARTMAS!DATE_ENTERED) = "" Then
                            Number_Of_Years_No_Sale = 0
                        Else
                            Number_Of_Years_No_Sale = Int((CDate(LOGDATE) - Null2Date(RSPARTMAS!DATE_ENTERED)) / TTLDYSIN1YR)
                        End If
                        If Number_Of_Years_No_Sale > 0 Then
                            'update by: IEBV 11022010_1200pm
                            'description:   tcn 13150
                            'A_SALES = N2Str2Zero(RSSHIPPING!Prev_Month) + N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_4)
                            A_SALES = N2Str2Zero(RSSHIPPING!Prev_Month) + N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3)
                            If A_SALES > 0 Then
                                vRankType = RANK_FAST_MOVING
                                GoTo con
                            Else
                                'update by: IEBV 11022010_1200pm
                                'description:   tcn 13150
                                'B_SALES = N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)
                                B_SALES = N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6)
                                If B_SALES > 0 Then
                                    vRankType = RANK_MEDIUM_MOVING
                                    GoTo con
                                Else
                                    'update by: IEBV 11022010_1200pm
                                    'description:   tcn 13150
                                    'C_SALES = N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12) + N2Str2Zero(RSSHIPPING!months_13)
                                    C_SALES = N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12) + N2Str2Zero(RSSHIPPING!months_13)
                                    If C_SALES > 0 Then
                                        vRankType = RANK_SLOW_MOVING
                                        GoTo con
                                    Else
                                        vRankType = RANK_NON_MOVING
                                        GoTo con
                                    End If
                                End If
                            End If
                        Else
                            vRankType = RANK_NEW_PARTS
                            GoTo con
                        End If
 
                    End If
                        
                
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

con:
            
                SMonths_12 = N2Str2Zero(RSSHIPPING!months_12): SMonths_11 = N2Str2Zero(RSSHIPPING!Months_11): SMonths_10 = N2Str2Zero(RSSHIPPING!Months_10)
                SMonths_9 = N2Str2Zero(RSSHIPPING!Months_9): SMonths_8 = N2Str2Zero(RSSHIPPING!Months_8): SMonths_7 = N2Str2Zero(RSSHIPPING!Months_7)
                SMonths_6 = N2Str2Zero(RSSHIPPING!Months_6): SMonths_5 = N2Str2Zero(RSSHIPPING!Months_5): SMonths_4 = N2Str2Zero(RSSHIPPING!Months_4)
                SMonths_3 = N2Str2Zero(RSSHIPPING!Months_3): SMonths_2 = N2Str2Zero(RSSHIPPING!Months_2): SPrev_Month = N2Str2Zero(RSSHIPPING!Prev_Month)
                S_year1 = N2Str2Zero(RSSHIPPING!months_12) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Prev_Month)
                vTotSales = Format(S_year1, MAXIMUM_DIGIT)
            Else
                If Null2Date(RSPARTMAS!DATE_ENTERED) = "" Then
                    Number_Of_Years_No_Sale = 0
                Else
                    Number_Of_Years_No_Sale = Int((CDate(LOGDATE) - Null2Date(RSPARTMAS!DATE_ENTERED)) / TTLDYSIN1YR)
                End If
                If Number_Of_Years_No_Sale > 0 Then
                    vRankType = RANK_NON_MOVING
                Else
                    vRankType = RANK_NEW_PARTS
                End If
            End If

            RANK_COST = N2Str2Zero(RSPARTMAS!MAC)
            If RANK_COST < PESO_VALUE_FOR_ONE Then
                vSubClass = "1"
            ElseIf RANK_COST >= PESO_VALUE_FOR_ONE And RANK_COST < PESO_VALUE_FOR_TWO Then
                vSubClass = "2"
            ElseIf RANK_COST >= PESO_VALUE_FOR_ONE And RANK_COST < PESO_VALUE_FOR_THREE Then
                vSubClass = "3"
            Else
                vSubClass = "4"
            End If

            SIX_MONTHS_SALE = SMonths_6 + SMonths_5 + SMonths_4 + SMonths_3 + SMonths_2 + SPrev_Month
            MIDDLE_OF_SIX_MONTHS = SMonths_4 + SMonths_3
            CURRENT_DEMAND = SPrev_Month

            Set RSSHIPPING_COST = New ADODB.Recordset
            RSSHIPPING_COST.Open "Select * from PMIS_Shipping_Cost where [TYPE] = " & N2Str2Null(RSPARTMAS!Type) & " AND PARTNO = " & N2Str2Null(RSPARTMAS!STOCKNO), gconDMIS
            If Not RSSHIPPING_COST.EOF And Not RSSHIPPING_COST.BOF Then
                SMonths_12_Cost = N2Str2Zero(RSSHIPPING_COST!months_12): SMonths_11_Cost = N2Str2Zero(RSSHIPPING_COST!Months_11): SMonths_10_Cost = N2Str2Zero(RSSHIPPING_COST!Months_10)
                SMonths_9_Cost = N2Str2Zero(RSSHIPPING_COST!Months_9): SMonths_8_Cost = N2Str2Zero(RSSHIPPING_COST!Months_8): SMonths_7_Cost = N2Str2Zero(RSSHIPPING_COST!Months_7)
                SMonths_6_Cost = N2Str2Zero(RSSHIPPING_COST!Months_6): SMonths_5_Cost = N2Str2Zero(RSSHIPPING_COST!Months_5): SMonths_4_Cost = N2Str2Zero(RSSHIPPING_COST!Months_4)
                SMonths_3_Cost = N2Str2Zero(RSSHIPPING_COST!Months_3): SMonths_2_Cost = N2Str2Zero(RSSHIPPING_COST!Months_2): SPrev_Month_Cost = N2Str2Zero(RSSHIPPING_COST!Prev_Month)
                vTotSales_Cost = SPrev_Month_Cost + SMonths_2_Cost + SMonths_3_Cost + SMonths_4_Cost + SMonths_5_Cost + SMonths_6_Cost + SMonths_7_Cost + SMonths_8_Cost + SMonths_9_Cost + SMonths_10_Cost + SMonths_11_Cost + SMonths_12_Cost
                vMAD12_Cost = vTotSales_Cost / 12
            Else
                SMonths_12_Cost = 0: SMonths_11_Cost = 0: SMonths_10_Cost = 0
                SMonths_9_Cost = 0: SMonths_8_Cost = 0: SMonths_7_Cost = 0
                SMonths_6_Cost = 0: SMonths_5_Cost = 0: SMonths_4_Cost = 0
                SMonths_3_Cost = 0: SMonths_2_Cost = 0: SPrev_Month_Cost = 0
                vTotSales_Cost = SPrev_Month_Cost + SMonths_2_Cost + SMonths_3_Cost + SMonths_4_Cost + SMonths_5_Cost + SMonths_6_Cost + SMonths_7_Cost + SMonths_8_Cost + SMonths_9_Cost + SMonths_10_Cost + SMonths_11_Cost + SMonths_12_Cost
                vMAD12_Cost = vTotSales_Cost / 12
            End If

            Set rsForeCast_Qty = New ADODB.Recordset
            Set rsForeCast_Qty = gconDMIS.Execute("Select * from PMIS_Forecast_Qty Where Partno = " & N2Str2Null(RSPARTMAS!STOCKNO))
            If Not rsForeCast_Qty.EOF And Not rsForeCast_Qty.BOF Then
                SMonths_12_Qty = N2Str2Zero(rsForeCast_Qty!months_12): SMonths_11_Qty = N2Str2Zero(rsForeCast_Qty!Months_11): SMonths_10_Qty = N2Str2Zero(rsForeCast_Qty!Months_10)
                SMonths_9_Qty = N2Str2Zero(rsForeCast_Qty!Months_9): SMonths_8_Qty = N2Str2Zero(rsForeCast_Qty!Months_8): SMonths_7_Qty = N2Str2Zero(rsForeCast_Qty!Months_7)
                SMonths_6_Qty = N2Str2Zero(rsForeCast_Qty!Months_6): SMonths_5_Qty = N2Str2Zero(rsForeCast_Qty!Months_5): SMonths_4_Qty = N2Str2Zero(rsForeCast_Qty!Months_4)
                SMonths_3_Qty = N2Str2Zero(rsForeCast_Qty!Months_3): SMonths_2_Qty = N2Str2Zero(rsForeCast_Qty!Months_2): SPrev_Month_Qty = N2Str2Zero(rsForeCast_Qty!Prev_Month)
                ELEVEN_MONTHS_FORECAST = SPrev_Month_Qty + SMonths_2_Qty + SMonths_3_Qty + SMonths_4_Qty + SMonths_5_Qty + SMonths_6_Qty + SMonths_7_Qty + SMonths_8_Qty + SMonths_9_Qty + SMonths_10_Qty + SMonths_11_Qty
            Else
                SMonths_12_Qty = 0: SMonths_11_Qty = 0: SMonths_10_Qty = 0
                SMonths_9_Qty = 0: SMonths_8_Qty = 0: SMonths_7_Qty = 0
                SMonths_6_Qty = 0: SMonths_5_Qty = 0: SMonths_4_Qty = 0
                SMonths_3_Qty = 0: SMonths_2_Qty = 0: SPrev_Month_Qty = 0
                ELEVEN_MONTHS_FORECAST = SPrev_Month_Qty + SMonths_2_Qty + SMonths_3_Qty + SMonths_4_Qty + SMonths_5_Qty + SMonths_6_Qty + SMonths_7_Qty + SMonths_8_Qty + SMonths_9_Qty + SMonths_10_Qty + SMonths_11_Qty
            End If

            '============================================================================================
            'MAD REVISED BY FOLLOWING HARI STANDARD FOR SOQ AND vMAD12 = Format(vTotSales / 12, MAXIMUM_DIGIT)
            MOVING_AVERAGE_SIX_MONTHS = SIX_MONTHS_SALE / N_VALUE
            MOVING_MEDIAN = MIDDLE_OF_SIX_MONTHS / 2
            'CREATE A LINE AND FIND THE INTERCEPT AND DRAW THE SLOPE OF THE LINE
            SUMMATION_OF_X = 0: For PIYA = 1 To N_VALUE: SUMMATION_OF_X = SUMMATION_OF_X + PIYA: Next
            MEAN_OF_X = SUMMATION_OF_X / N_VALUE
            SUMMATION_OF_Y = SIX_MONTHS_SALE
            MEAN_OF_Y = SUMMATION_OF_Y / N_VALUE
            X_REGRESSION = SIX_MONTHS_SALE / N_VALUE
            If CURRENT_DEMAND > 0 Then
                Y_REGRESSION = (SIX_MONTHS_SALE * N_VALUE) / CURRENT_DEMAND
            Else
                Y_REGRESSION = 0
            End If
            SLOPE_OF_THE_LINE = Format(((SUMMATION_OF_X * SUMMATION_OF_Y) - (N_VALUE * (MEAN_OF_X * MEAN_OF_Y))) / ((SUMMATION_OF_X ^ 2) - (N_VALUE * (MEAN_OF_X ^ 2))), "###0.00")
            INTERCEPT_VALUE_AT_ZERO = Format(MEAN_OF_Y - (SLOPE_OF_THE_LINE * MEAN_OF_X), "###0.00")
            LINEAR_REGRESSION = Format(INTERCEPT_VALUE_AT_ZERO + (SLOPE_OF_THE_LINE * MEAN_OF_X), "###0.00")

            'MIGHT REVISED LINEAR REGRESSION - FML 05082007
            LINEAR_REGRESSION = INTERCEPT_VALUE_AT_ZERO + (SLOPE_OF_THE_LINE * SUMMATION_OF_X)

            'REVISED LINEAR REGRESSION FORMULA'AXP
            LVAL = LINEARREGRESSION(SMonths_6, SMonths_5, SMonths_4, SMonths_3, SMonths_2, SPrev_Month)
            SLOPE_OF_THE_LINE = LVAL(0)
            INTERCEPT_VALUE_AT_ZERO = LVAL(1)
            LINEAR_REGRESSION = LVAL(2)

            NEARES_DEMAND_FORECAST = Format(GET_LESS_ERROR_PERCENTAGE_DEVIATION(CURRENT_DEMAND, MOVING_AVERAGE_SIX_MONTHS, MOVING_MEDIAN, LINEAR_REGRESSION), "###0.00")
            MEAN_ABSOULTE_DEVIATION = Format(Abs((S_year1 / (N_VALUE * 2)) - ((ELEVEN_MONTHS_FORECAST + NEARES_DEMAND_FORECAST) / (N_VALUE * 2))), "###0.00")

            'USES FORMULA A FOR LEVEL OF SERVICE. WILL REVISE LATER FOR N.S. (ANY NON STOCK PARTNUMBER)
            '=====================================================================================================================================================

            VarPartNo = N2Str2Null(RSPARTMAS!STOCKNO): varRECORDDATE = "": varTRANYPE = "": varTRANNO = ""
            varC_REQUESTED = 0: varC_SERVED = 0: varC_UNSERVED = 0: varC_BACKORDER = 0: varC_FILLRATE = 0
            varS_REQUESTED = 0: varS_SERVED = 0: varS_UNSERVED = 0: varS_BACKORDER = 0: varS_FILLRATE = 0
            varD_ORDERED = 0: varD_SERVED = 0: varD_UNSERVED = 0: varD_BACKORDER = 0: varD_FILLRATE = 0:
            varD_EMERGENCY_PO = 0:

            Set rsDemand_Monitoring = New ADODB.Recordset
            Set rsDemand_Monitoring = gconDMIS.Execute("Select * from PMIS_vw_Demand_Monitoring where PARTNO = " & VarPartNo)
            If Not rsDemand_Monitoring.EOF And Not rsDemand_Monitoring.BOF Then
                varC_REQUESTED = varC_REQUESTED + N2Str2Zero(rsDemand_Monitoring!C_REQUESTED)
                varC_SERVED = varC_SERVED + N2Str2Zero(rsDemand_Monitoring!C_SERVED)
                varC_UNSERVED = varC_UNSERVED + N2Str2Zero(rsDemand_Monitoring!C_UNSERVED)
                varC_FILLRATE = varC_FILLRATE + N2Str2Zero(rsDemand_Monitoring!C_FILLRATE)
                varS_REQUESTED = varC_REQUESTED + N2Str2Zero(rsDemand_Monitoring!S_REQUESTED)
                varS_SERVED = varC_SERVED + N2Str2Zero(rsDemand_Monitoring!S_SERVED)
                varS_UNSERVED = varC_UNSERVED + N2Str2Zero(rsDemand_Monitoring!S_UNSERVED)
                varS_FILLRATE = varC_FILLRATE + N2Str2Zero(rsDemand_Monitoring!S_FILLRATE)
                varD_ORDERED = varD_ORDERED + N2Str2Zero(rsDemand_Monitoring!D_ORDERED)
                varD_SERVED = varD_SERVED + N2Str2Zero(rsDemand_Monitoring!D_SERVED)
                varD_UNSERVED = varD_UNSERVED + N2Str2Zero(rsDemand_Monitoring!D_UNSERVED)
                varD_BACKORDER = varD_BACKORDER + N2Str2Zero(rsDemand_Monitoring!D_BACKORDER)
                varD_FILLRATE = varD_FILLRATE + N2Str2Zero(rsDemand_Monitoring!D_FILLRATE)
                varD_EMERGENCY_PO = varD_EMERGENCY_PO + N2Str2Zero(rsDemand_Monitoring!D_EMERGENCY_PO)
            Else
                varC_REQUESTED = 0
                varC_SERVED = 0
                varC_UNSERVED = 0
                varC_FILLRATE = 0
                varS_REQUESTED = 0
                varS_SERVED = 0
                varS_UNSERVED = 0
                varS_FILLRATE = 0
                varD_ORDERED = 0
                varD_SERVED = 0
                varD_UNSERVED = 0
                varD_BACKORDER = 0
                varD_FILLRATE = 0
                varD_EMERGENCY_PO = 0
            End If
            TOTAL_DEMAND = varC_REQUESTED + varS_REQUESTED
            EMERGENCY_PURCHASES = varD_EMERGENCY_PO
            LOST_SALES = varC_UNSERVED + varS_UNSERVED
            If TOTAL_DEMAND > 0 Then
                LEVEL_OF_SERVICE = (TOTAL_DEMAND - (EMERGENCY_PURCHASES + LOST_SALES)) / TOTAL_DEMAND
            Else
                LEVEL_OF_SERVICE = 0
            End If
            LOS = LEVEL_OF_SERVICE
            'INVERSE FUNCTION IS 1/X
            SAFETY_FACTOR = Format(INVERSE_FUNCTION(LEVEL_OF_SERVICE), "###0.00")
            LEAD_TIME = Format(HARI_LEAD_TIME, "###0.00")
            ORDER_FREQUENCY = Format(HARI_ORDER_FREQUENCY, "###0.00")
            SAFETY_STOCK = Format(SAFETY_FACTOR * LEAD_TIME * MEAN_ABSOULTE_DEVIATION * (Sqr(LEAD_TIME + ORDER_FREQUENCY)), "###0.00")
            ONHAND = N2Str2Zero(RSPARTMAS!ONHAND): ON_ORDER = N2Str2Zero(RSPARTMAS!ONORDER): BACK_ORDER = N2Str2Zero(RSPARTMAS!BACKORDER)
            SUGGESTED_ORDER_QTY = Format((NEARES_DEMAND_FORECAST * (LEAD_TIME + ORDER_FREQUENCY)) + SAFETY_STOCK + BACK_ORDER - (ONHAND + ON_ORDER), "###0.00")
            gconDMIS.Execute "UPDATE PMIS_STOCKMAS set " & _
                "invclass = " & N2Str2Null(vRankType) & "," & _
                "subinvclas = " & N2Str2Null(vSubClass) & "," & _
                "LOST_SALES = " & N2Str2Zero(LOST_SALES) & "," & _
                "LEVEL_OF_SERVICE = " & N2Str2Zero(LEVEL_OF_SERVICE) & "," & _
                "SSTOCK = " & N2Str2Zero(SAFETY_STOCK) & "," & _
                "SOQ = " & N2Str2Zero(SUGGESTED_ORDER_QTY) & "," & _
                "mad = " & N2Str2Zero(MEAN_ABSOULTE_DEVIATION) & _
                " where TYPE = " & N2Str2Null(RSPARTMAS!Type) & " AND STOCKNO = " & N2Str2Null(RSPARTMAS!STOCKNO) & " AND ACTIVE = 'Y'"
                
            If Null2String(RSPARTMAS!Type) = "P" Then
                Set rsForeCast_Qty = New ADODB.Recordset
                Set rsForeCast_Qty = gconDMIS.Execute("Select * from PMIS_ForeCast_Qty where partno = " & N2Str2Null(RSPARTMAS!STOCKNO))
                If Not rsForeCast_Qty.EOF And Not rsForeCast_Qty.BOF Then
                    gconDMIS.Execute "update PMIS_FORECAST_QTY set " & _
                        "PREV_MONTH = " & N2Str2Zero(NEARES_DEMAND_FORECAST) & _
                        " where PARTNO = " & N2Str2Null(RSPARTMAS!STOCKNO)
                Else
                    gconDMIS.Execute "insert into PMIS_FORECAST_QTY (partno,curr_month,freq_curr)" & _
                        " values (" & N2Str2Null(RSPARTMAS!STOCKNO) & ", " & NumericVal(NEARES_DEMAND_FORECAST) & ", 1)"
                End If

                Set rsLevel_Of_Service = New ADODB.Recordset
                Set rsLevel_Of_Service = gconDMIS.Execute("Select * from PMIS_Level_Of_Service where partno = " & N2Str2Null(RSPARTMAS!STOCKNO))
                If Not rsLevel_Of_Service.EOF And Not rsLevel_Of_Service.BOF Then
                    gconDMIS.Execute "update PMIS_Level_Of_Service set " & _
                        "PREV_MONTH = " & N2Str2Zero(LEVEL_OF_SERVICE) & _
                        " where PARTNO = " & N2Str2Null(RSPARTMAS!STOCKNO)
                Else
                    gconDMIS.Execute "insert into PMIS_Level_Of_Service (partno,curr_month,freq_curr)" & _
                        " values (" & N2Str2Null(RSPARTMAS!STOCKNO) & ", " & NumericVal(LEVEL_OF_SERVICE) & ", 1)"
                End If
            End If

            gconDMIS.Execute "insert into PMIS_RankFle " & _
                "(TYPE,partno,partdesc,invclass,subinvclas,onhand,intercept0,slopeline,lr6,mm6,mad6,sales6,mad12,sales12,last_recd,mac,month_gen,prev_month,months_2,months_3,months_4,months_5,months_6,months_7,months_8,months_9,months_10,months_11,months_12,prevclass,prevsclas,date_gen)" & _
                " values ('P'," & N2Str2Null(RSPARTMAS!STOCKNO) & ", " & P_STOCKDESC & _
                "," & N2Str2Null(vRankType) & ", " & N2Str2Null(vSubClass) & ", " & P_Onhand & "," & INTERCEPT_VALUE_AT_ZERO & "," & SLOPE_OF_THE_LINE & "," & LINEAR_REGRESSION & "," & MOVING_MEDIAN & "," & SIX_MONTHS_SALE & "," & MOVING_AVERAGE_SIX_MONTHS & _
                "," & vMAD12 & ", " & NumericVal(vTotSales) & ", " & P_Last_recd & ", " & P_MAC & ", " & Month(LOGDATE) & ", " & SPrev_Month & _
                "," & SMonths_2 & ", " & SMonths_3 & ", " & SMonths_4 & _
                "," & SMonths_5 & ", " & SMonths_6 & ", " & SMonths_7 & _
                "," & SMonths_8 & ", " & SMonths_9 & ", " & SMonths_10 & _
                "," & SMonths_11 & ", " & SMonths_12 & ", " & vPrevClass & ", " & vPrevSClass & ", " & N2Date2Null(LOGDATE) & ")"
            
            If Null2String(RSPARTMAS!Type) = "P" Then
                gconDMIS.Execute "insert into PMIS_RankSales " & _
                    "(TYPE,partno,partdesc,invclass,subinvclas,onhand,mad12,sales12,last_recd,mac,month_gen,prev_month,months_2,months_3,months_4,months_5,months_6,months_7,months_8,months_9,months_10,months_11,months_12,prevclass,prevsclas,date_gen)" & _
                    " values ('P'," & N2Str2Null(RSPARTMAS!STOCKNO) & ", " & P_STOCKDESC & _
                    "," & N2Str2Null(vRankType) & ", " & N2Str2Null(vSubClass) & ", " & P_Onhand & _
                    "," & vMAD12_Cost & ", " & NumericVal(vTotSales_Cost) & ", " & P_Last_recd & ", " & P_MAC & ", " & Month(LOGDATE) & ", " & SPrev_Month_Cost & _
                    "," & SMonths_2_Cost & ", " & SMonths_3_Cost & ", " & SMonths_4_Cost & _
                    "," & SMonths_5_Cost & ", " & SMonths_6_Cost & ", " & SMonths_7_Cost & _
                    "," & SMonths_8_Cost & ", " & SMonths_9_Cost & ", " & SMonths_10_Cost & _
                    "," & SMonths_11_Cost & ", " & SMonths_12_Cost & ", " & vPrevClass & ", " & vPrevSClass & ", " & N2Date2Null(LOGDATE) & ")"

                gconDMIS.Execute "insert into PMIS_Demand_Monitoring " & _
                    "(partno,C_REQUESTED,C_SERVED,C_UNSERVED,C_BACKORDER,C_FILLRATE,S_REQUESTED,S_SERVED,S_UNSERVED,S_BACKORDER,S_FILLRATE,D_ORDERED,D_SERVED,D_UNSERVED,D_BACKORDER,D_FILLRATE,D_EMERGENCY_PO,D_ONORDER,date_gen)" & _
                    " values (" & VarPartNo & "," & varC_REQUESTED & "," & varC_SERVED & "," & varC_UNSERVED & "," & varC_BACKORDER & "," & varC_FILLRATE & "," & varS_REQUESTED & "," & varS_SERVED & "," & varS_UNSERVED & "," & varS_BACKORDER & "," & varS_FILLRATE & "," & varD_ORDERED & "," & varD_SERVED & "," & varD_UNSERVED & "," & varD_BACKORDER & "," & varD_FILLRATE & "," & varD_EMERGENCY_PO & "," & varD_ONORDER & "," & N2Date2Null(LOGDATE) & ")"
                gconDMIS.Execute "insert into PMIS_Demand_Forecast " & _
                    "(partno,description,Date_Gen,OH,OO,BO,Mad6,MM6,LR,MAD,SS,SOQ)" & _
                    " values (" & VarPartNo & "," & P_STOCKDESC & "," & N2Date2Null(LOGDATE) & "," & ONHAND & "," & ON_ORDER & "," & BACK_ORDER & "," & MOVING_AVERAGE_SIX_MONTHS & "," & MOVING_MEDIAN & "," & LINEAR_REGRESSION & "," & MEAN_ABSOULTE_DEVIATION & "," & SAFETY_STOCK & "," & SUGGESTED_ORDER_QTY & ")"
            End If
            i = i + 1: progCPB.Value = (i / RSPARTMAS.RecordCount) * 100: labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            RSPARTMAS.MoveNext
        Loop
        labProcessing.Caption = "": DoEvents: MsgSpeech "Generating Rank File Completed!"
    Else
        MsgSpeechBox "Error opening Part Master File"
    End If
    Screen.MousePointer = 0: MsgSpeech "Updating Demand Monitoring File": Me.Caption = "Updating Demand Monitoring File"
    labCPB.Caption = "Updating Demand Monitoring File... Please Wait...": DoEvents
    progCPB.Value = 0: DoEvents: Screen.MousePointer = 11
    
    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
        " PMIS_STOCKMAS.LM_S_ONREQUEST = ISNULL(S_ONREQUEST,0)," & _
        " PMIS_STOCKMAS.LM_S_REQSERVED = ISNULL(S_REQSERVED,0)," & _
        " PMIS_STOCKMAS.LM_S_REQUNSERVED = ISNULL(S_REQUNSERVED,0)," & _
        " PMIS_STOCKMAS.LM_S_REQFILLRATE = ISNULL(S_REQFILLRATE,0)," & _
        " PMIS_STOCKMAS.LM_C_ONREQUEST = ISNULL(ONREQUEST,0)," & _
        " PMIS_STOCKMAS.LM_C_REQSERVED = ISNULL(REQSERVED,0)," & _
        " PMIS_STOCKMAS.LM_C_REQUNSERVED = ISNULL(REQUNSERVED,0)," & _
        " PMIS_STOCKMAS.LM_C_REQFILLRATE = ISNULL(REQFILLRATE,0)," & _
        " PMIS_STOCKMAS.LM_D_ORDERED = ISNULL(ORDERED,0)," & _
        " PMIS_STOCKMAS.lastm_oo = ISNULL(ONORDER,0)," & _
        " PMIS_STOCKMAS.LM_D_SERVED = ISNULL(SERVED,0)," & _
        " PMIS_STOCKMAS.LM_D_UNSERVED = ISNULL(UNSERVED,0)," & _
        " PMIS_STOCKMAS.LM_D_BACKORDER = ISNULL(BACKORDER,0)," & _
        " PMIS_STOCKMAS.LM_D_FILLRATE = ISNULL(FILLRATE,0)," & _
        " PMIS_STOCKMAS.LM_D_EMERGENCY_PO = ISNULL(EMERGENCY_PO,0) WHERE ACTIVE = 'Y'"
    
    progCPB.Value = 50: DoEvents
    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
        " S_ONREQUEST = 0," & _
        " S_REQSERVED = 0," & _
        " S_REQUNSERVED = 0," & _
        " S_REQFILLRATE = 0," & _
        " ONREQUEST = 0," & _
        " REQSERVED = 0," & _
        " REQUNSERVED = 0," & _
        " REQFILLRATE = 0," & _
        " ORDERED = 0," & _
        " ONORDER = 0," & _
        " SERVED = 0," & _
        " UNSERVED = 0," & _
        " BACKORDER = 0," & _
        " FILLRATE = 0," & _
        " EMERGENCY_PO = 0 WHERE ACTIVE = 'Y'"
    progCPB.Value = 100: DoEvents
    Screen.MousePointer = 0: Me.Caption = "Updating Demand Monitoring Complete!": labCPB.Caption = "Updating Demand Monitoring Complete!"
    GenRankFile = True
    Exit Function
ErrorCode:
    
    GenRankFile = False
End Function

Function CreateStockStatus() As Boolean
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    progCPB.Value = 0
    Me.Caption = "Updating Part Master File"
    labCPB.Caption = "Updating Stocks Master File for Stock Status... Please Wait..."
    DoEvents
    progCPB.Value = 100
    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
        " sstock = ISNULL(mad,0) * " & PARTS_SSTOCK_NO_MONTHS & "," & _
        " resservice = ISNULL(mad,0)" & _
        " where invclass = 'A' AND ACTIVE = 'Y'"
    
    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
        " sstock = ISNULL(mad,0)," & _
        " resservice = 0" & _
        " where invclass <> 'A' AND ACTIVE = 'Y'"
    
    DoEvents: Screen.MousePointer = 11: progCPB.Value = 0: Me.Caption = "Creating Stock Status"
    labCPB.Caption = "Create Stock Status Master File... Please Wait...": DoEvents: progCPB.Value = 100
    
    gconDMIS.Execute "insert into PMIS_StkStat " & _
        "(TYPE, STOCKTYPE, NON_HARI, STOCKNO,STOCKDESC,onhand,mac,mad,sstock,resservice,onorder,ADJ_ADD,ADJ_MINUS,BACKORD,SOQ,SRP,TD,EM_PO,LS,LOS)" & _
        " select TYPE, STOCKTYPE, NON_HARI, STOCKNO,STOCKDESC,ISNULL(OnHand,0),ISNULL(Mac,0),ISNULL(Mad,0),ISNULL(SStock,0),ISNULL(ResService,0),ISNULL(OnOrder,0),ISNULL(TADJQTY_IN,0),ISNULL(TADJQTY_OUT,0),ISNULL(BACKORDER,0),ISNULL(SOQ,0),ISNULL(SRP,0),(ISNULL(ONREQUEST,0) + ISNULL(S_ONREQUEST,0)),ISNULL(EMERGENCY_PO,0),ISNULL(LOST_SALES,0),ISNULL(LEVEL_OF_SERVICE,0) from PMIS_STOCKMAS WHERE ACTIVE = 'Y' order by STOCKNO asc"
    
    gconDMIS.Execute "update PMIS_StkStat set date_gen = " & N2Date2Null(LOGDATE) & " where date_gen IS NULL"
    
    MsgSpeech "Create Stock Status Complete!"
    txtCurrentActivity.Text = txtCurrentActivity.Text & vbCrLf & "Month End Closing Successfully Completed!": DoEvents
    Screen.MousePointer = 0
    DoEvents
    CreateStockStatus = True
    Exit Function
ErrorCode:
    CreateStockStatus = False
End Function

Sub MoveTdaytran(aydi As Long)
    gconDMIS.Execute ("Insert into PMIS_DayTran Select * from PMIS_Tdaytran where id = " & aydi)
    gconDMIS.Execute ("Delete from PMIS_Tdaytran where id = " & aydi)
End Sub

Sub MoveOrdHd(aydi As Long)
    gconDMIS.Execute ("Insert into PMIS_Ord_Hist Select * from PMIS_Ord_Hd where id = " & aydi)
    gconDMIS.Execute ("delete from PMIS_Ord_Hd where id = " & aydi)
End Sub

Sub MoveRRhd(aydi As Long)
    gconDMIS.Execute ("Insert into PMIS_Rec_Hist Select * from PMIS_RR_Hd where id = " & aydi)
    gconDMIS.Execute ("delete from PMIS_RR_Hd where id = " & aydi)
End Sub

Sub MovePOhd(aydi As Long)
    gconDMIS.Execute ("Insert into PMIS_Po_Hist Select * from PMIS_Po_Hd where id = " & aydi)
    gconDMIS.Execute ("delete from PMIS_Po_Hd where id = " & aydi)
End Sub

Sub InitGrid()
    With grdTransactions
        .Cell(0, 1).Text = "Processed"
        .Cell(0, 2).Text = "TYPE"
        .Cell(0, 3).Text = "Tran. Type"
        .Cell(0, 4).Text = "Tran. No."
        .Cell(0, 5).Text = "Status"

        .Column(0).Width = 0
        .Column(1).Width = 60
        .Column(2).Width = 70
        .Column(3).Width = 60
        .Column(4).Width = 70
        .Column(5).Width = 60

        .Column(1).CellType = cellCheckBox
        .Column(1).Locked = True
        .Column(2).CellType = cellTextBox
        .Column(2).Locked = True
        .Column(3).CellType = cellTextBox
        .Column(3).Locked = True
        .Column(4).CellType = cellTextBox
        .Column(4).Locked = True
        .Column(5).CellType = cellTextBox
        .Column(5).Locked = True

        grdTransactions.DefaultFont = "TAHOMA"
    End With
End Sub

Sub CheckInventoryBalances()
    Dim vOH                                            As Double
    Dim vTP, vTR, vTI                                  As Double
    

    Dim RSPARTMAS                                      As ADODB.Recordset

    vOH = 0: vTP = 0: vTR = 0: vTI = 0:

    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select sum(LASTM_OH) as TOTAL_LASTM_OH,SUM(onhand) AS TOTAL_ONHAND, SUM(tpoqty) AS TOTAL_TPOQTY,SUM(trecqty) AS TOTAL_TRECQTY,SUM(tissqty) AS TOTAL_TISSQTY from PMIS_PARTMAS WHERE [TYPE] = 'P'")
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
    
'updated by : IEBV 0114211_0400PM
'------------------------------------------------------------------------
        ictr_parts = ictr_parts + 1
        If ictr_parts <> 1 Then
            vOH = vOH + N2Str2IntZero(RSPARTMAS!TOTAL_LASTM_OH)
        Else
            vOH = vOH + N2Str2IntZero(RSPARTMAS!total_onhand)
        End If
'------------------------------------------------------------------------
        vTP = vTP + N2Str2IntZero(RSPARTMAS!TOTAL_tpoqty)
        vTR = vTR + N2Str2IntZero(RSPARTMAS!TOTAL_trecqty)
        vTI = vTI + N2Str2IntZero(RSPARTMAS!TOTAL_tissqty)
        
        txtOH.Text = Format(vOH, DIGIT_FORMAT)
        txtTP.Text = Format(vTP, DIGIT_FORMAT)
        txtTR.Text = Format(vTR, DIGIT_FORMAT)
        txtTI.Text = Format(vTI, DIGIT_FORMAT)
    End If
    Set RSPARTMAS = Nothing

    vOH = 0: vTP = 0: vTR = 0: vTI = 0

    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select sum(LASTM_OH) as TOTAL_LASTM_OH,SUM(onhand) AS TOTAL_ONHAND, SUM(tpoqty) AS TOTAL_TPOQTY,SUM(trecqty) AS TOTAL_TRECQTY,SUM(tissqty) AS TOTAL_TISSQTY from PMIS_STOCKMAS WHERE [TYPE] = 'M'")
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
   
 'updated by : IEBV 0114211_0400PM
 '------------------------------------------------------------------------
        ictr_mat = ictr_mat + 1
        If ictr_mat <> 1 Then
            vOH = vOH + N2Str2IntZero(RSPARTMAS!TOTAL_LASTM_OH)
        Else
            vOH = vOH + N2Str2IntZero(RSPARTMAS!total_onhand)
        End If
 '------------------------------------------------------------------------
        'vOH = vOH + N2Str2IntZero(RSPARTMAS!TOTAL_ONHAND)
        vTP = vTP + N2Str2IntZero(RSPARTMAS!TOTAL_tpoqty)
        vTR = vTR + N2Str2IntZero(RSPARTMAS!TOTAL_trecqty)
        vTI = vTI + N2Str2IntZero(RSPARTMAS!TOTAL_tissqty)

        txtOHM.Text = Format(vOH, DIGIT_FORMAT)
        txtTPM.Text = Format(vTP, DIGIT_FORMAT)
        txtTRM.Text = Format(vTR, DIGIT_FORMAT)
        txtTIM.Text = Format(vTI, DIGIT_FORMAT)
    End If
    Set RSPARTMAS = Nothing

    vOH = 0: vTP = 0: vTR = 0: vTI = 0

    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select sum(LASTM_OH) as TOTAL_LASTM_OH,SUM(onhand) AS TOTAL_ONHAND, SUM(tpoqty) AS TOTAL_TPOQTY,SUM(trecqty) AS TOTAL_TRECQTY,SUM(tissqty) AS TOTAL_TISSQTY from PMIS_STOCKMAS WHERE [TYPE] = 'A'")
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
    
 'updated by : IEBV 0114211_0400PM
 '------------------------------------------------------------------------
        ictr_acc = ictr_acc + 1
        If ictr_acc <> 1 Then
            vOH = vOH + N2Str2IntZero(RSPARTMAS!TOTAL_LASTM_OH)
        Else
            vOH = vOH + N2Str2IntZero(RSPARTMAS!total_onhand)
        End If
 '------------------------------------------------------------------------
        'vOH = vOH + N2Str2IntZero(RSPARTMAS!TOTAL_ONHAND)
        vTP = vTP + N2Str2IntZero(RSPARTMAS!TOTAL_tpoqty)
        vTR = vTR + N2Str2IntZero(RSPARTMAS!TOTAL_trecqty)
        vTI = vTI + N2Str2IntZero(RSPARTMAS!TOTAL_tissqty)

        txtOHA.Text = Format(vOH, DIGIT_FORMAT)
        txtTPA.Text = Format(vTP, DIGIT_FORMAT)
        txtTRA.Text = Format(vTR, DIGIT_FORMAT)
        txtTIA.Text = Format(vTI, DIGIT_FORMAT)
    End If
    Set RSPARTMAS = Nothing
End Sub

Private Sub cmdT_Close_Click()
    PIC_UNPOSTED.Visible = False
End Sub

Private Sub cmdT_Print_Click()
    On Error GoTo ErrorCode:
    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim intRow                                         As Integer    ' counter
    Dim intCol                                         As Integer    ' counter
    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If
    On Error Resume Next
    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet
    wsXL.Name = "UNPOSTED TRANSACTION DETAILES"
    For intCol = 0 To LST_UNPOSTED.Columns.Count
        wsXL.Cells(1, intCol + 1).Value = "" & CStr(LST_UNPOSTED.Columns(intCol).Caption) & "  "
    Next
    For intRow = 0 To LST_UNPOSTED.Rows.Count
        For intCol = 0 To LST_UNPOSTED.Columns.Count
            wsXL.Cells(intRow + 2, intCol + 1).Value = "" & CStr(LST_UNPOSTED.Rows(intRow).Record(intCol).Value) & "  "
        Next
    Next
    For intCol = 1 To LST_UNPOSTED.Columns.Count
        wsXL.Columns(intCol).AutoFit
    Next
    wsXL.Range("A1", Right(wsXL.Columns(LST_UNPOSTED.Columns.Count).AddressLocal, 1) & LST_UNPOSTED.Rows.Count + 1).AutoFormat 2
    objXL.Visible = True
    Exit Sub
ErrorCode:
    MsgBox err.Description
    err.Clear
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    CheckInventoryBalances
    Me.Caption = "Month-End Processing"
    InitGrid
    AddColumnHeader "Date,Tran#,TranType,Type", LST_UNPOSTED
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Process", "MONTH-END PROCESSING") = False Then Exit Sub
    Dim RSMONTHENDVALIDATEION                          As ADODB.Recordset
    Dim UNPOSTED_RR                                    As Long
    Dim UNPOSTED_ORD                                   As Long
    Dim UNPOSTED_PO                                    As Long
    Dim SQLX                                           As String
    Dim RS_UNPOSTED_VALIDITY                           As ADODB.Recordset
    Dim RSHD                                           As ADODB.Recordset
    Dim CMD                                            As ADODB.Command
    Dim ictr                                           As Integer
    Dim COUNTER                                        As Integer
    Dim RSMAC                                          As ADODB.Recordset
    Dim RSMAC_CTR                                      As Integer
    Dim CMDMAC                                         As ADODB.Command
    Dim CMDMAC_FIX                                     As ADODB.Command
    
    Set RSMONTHENDVALIDATEION = gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_ORD_HD WHERE TRANTYPE IN ('RIV','CHG','CSH','DR','ADB','PRS') AND (ISNULL(STATUS,'')='' OR  ISNULL(STATUS,'')='U' OR ISNULL(STATUS,'')='N')")
    UNPOSTED_ORD = RSMONTHENDVALIDATEION.Fields(0).Value
    Set RSMONTHENDVALIDATEION = gconDMIS.Execute("SELECT COUNT(*) STATUS FROM PMIS_RR_HD WHERE  (ISNULL(STATUS,'')='' OR  ISNULL(STATUS,'')='U' OR ISNULL(STATUS,'')='N')")
    UNPOSTED_RR = RSMONTHENDVALIDATEION.Fields(0).Value
    Set RSMONTHENDVALIDATEION = gconDMIS.Execute("select COUNT(*) from pmis_po_hd where (isnull(status, '') = '' or status = 'N')")
    UNPOSTED_PO = RSMONTHENDVALIDATEION.Fields(0).Value
    If (UNPOSTED_ORD + UNPOSTED_RR + UNPOSTED_PO) > 0 Then
        MsgBox "" & vbCrLf & "Unposted Issuances:" & UNPOSTED_ORD & vbCrLf & _
               "Unposted Receiving: " & UNPOSTED_RR & vbCrLf & _
               "Unposted Purchase Order: " & UNPOSTED_PO & vbCrLf & _
               "Either Post or Cancel Those Transactions Prior To Proceed To Month End Processing", vbInformation, "Unposted Transaction Details!!"

        SQLX = "SELECT TRANDATE,TRANNO,TRANTYPE,TYPE FROM PMIS_ORD_HD WHERE ISNULL(STATUS2,'')<>'R' AND TRANTYPE IN ('RIV','CHG','CSH','DR','ADB','PRS') AND (ISNULL(STATUS,'')='' OR  ISNULL(STATUS,'')='U' OR ISNULL(STATUS,'')='N') "
        SQLX = SQLX & " Union All " & vbCrLf
        SQLX = SQLX & "SELECT TRANDATE,TRANNO,TRANTYPE + '-ADB FILL' ,TYPE FROM PMIS_ORD_HD WHERE ISNULL(STATUS2,'')='R' AND TRANTYPE ='RIV'  AND (ISNULL(STATUS,'')='' OR  ISNULL(STATUS,'')='U' OR ISNULL(STATUS,'')='N') " & vbCrLf
        SQLX = SQLX & " Union All " & vbCrLf
        SQLX = SQLX & " SELECT RRDATE,RRNO,'RR',TYPE FROM PMIS_RR_HD WHERE  (ISNULL(STATUS,'')='' OR  ISNULL(STATUS,'')='U' OR ISNULL(STATUS,'')='N') "
        SQLX = SQLX & " Union All " & vbCrLf
        SQLX = SQLX & " SELECT PODATE,PONO,'PO',TYPE FROM PMIS_PO_HD WHERE  (ISNULL(STATUS,'')='' OR  ISNULL(STATUS,'')='N') "
        
        Set RS_UNPOSTED_VALIDITY = gconDMIS.Execute(SQLX)
        Label.Caption = RS_UNPOSTED_VALIDITY.RecordCount
        flex_FillReportView RS_UNPOSTED_VALIDITY, LST_UNPOSTED
        PIC_UNPOSTED.Visible = True
        PIC_UNPOSTED.ZOrder 0
        InitializeRC
        Exit Sub
    End If

'updated by: IEBV 0943_120211
'description: to check if the ledger is balance and fix the wrong mac
'---------------------------------------------------------------------------------------------
    COUNTER = 3
    ReDim XTYPE(COUNTER) As String
    XTYPE(1) = "A"  'Accessories
    XTYPE(2) = "M"  'Materials
    XTYPE(3) = "P"  'Parts
    Dim strmessage                      As String
    Dim VPARTS                          As String
    Screen.MousePointer = 11
    For ictr = 1 To COUNTER
        If XTYPE(ictr) = "P" Then
            VPARTS = "..Parts.."
        ElseIf XTYPE(ictr) = "A" Then
            VPARTS = "..Accessories.."
        Else
            VPARTS = "..Materials.."
        End If
        txtCurrentActivity.Text = "Checking Ledger Balance........( " & VPARTS & " )": DoEvents
        Set CMD = New ADODB.Command
        With CMD
            .NamedParameters = True
            .CommandType = adCmdStoredProc
            .CommandText = "SP_PMIS_VIEW_BALANCED_STOCK"
            .ActiveConnection = gconDMIS
            .CommandTimeout = 1000
            .Parameters.Append .CreateParameter("@STOCKTYPE", adVarChar, adParamInput, 50, XTYPE(ictr))
        End With

        Set RSHD = CMD.Execute: DoEvents
        labCPB.Caption = "Please Wait"
        On Error Resume Next
        progCPB.Value = (RSHD.AbsolutePosition / RSHD.RecordCount) * 100: DoEvents

        If Not (RSHD.EOF And RSHD.BOF) Then
            strmessage = "The monthend cannot continue. " & vbCrLf
            strmessage = strmessage + "The Ledger is not Balance. " & vbCrLf
            strmessage = strmessage + "Please run the view unbalance and balance the ledger. " & vbCrLf
            MsgBox strmessage, vbInformation + vbOKOnly
            txtCurrentActivity.Text = "Checking Ledger Balance........"
            labCPB.Caption = "100%"
            Screen.MousePointer = 0
            Exit Sub
        End If
    Next ictr
    
    txtCurrentActivity.Text = "Checking Ledger Balance........" & vbCrLf
    Set CMDMAC = New ADODB.Command
        With CMDMAC
                .NamedParameters = True
                .CommandType = adCmdStoredProc
                .CommandText = "sp_mac_error_finder"
                .ActiveConnection = gconDMIS
                .CommandTimeout = 1000
        End With
    DoEvents
    txtCurrentActivity.Text = txtCurrentActivity.Text + "Checking Mac........" & vbCrLf
    CMDMAC.Execute: DoEvents
    DoEvents
'2 : BEGINNING MAC AND COST ARE NOT EQUAL and tag it to 3
    RSMAC_CTR = 0
    RSMAC_CTR = (gconDMIS.Execute("Select count(*) from pmis_stockmas where tag_mac = '2'").Fields(0).Value)
    If RSMAC_CTR > 0 Then
        Set CMDMAC_FIX = New ADODB.Command
        Screen.MousePointer = 11
        With CMDMAC_FIX
            .NamedParameters = True
            .CommandType = adCmdStoredProc
            .CommandText = "sp_fix_mac_2"
            .ActiveConnection = gconDMIS
            .CommandTimeout = 1000
        End With
        DoEvents
        CMDMAC_FIX.Execute: DoEvents
        DoEvents
    End If
'2 : WRONG COMPUTATION OF MAC OR COST and tag it to 3
    RSMAC_CTR = 0
    RSMAC_CTR = (gconDMIS.Execute("Select count(*) from pmis_stockmas where tag_mac = '1'").Fields(0).Value)
    If RSMAC_CTR > 0 Then
        Set CMDMAC_FIX = New ADODB.Command
        Screen.MousePointer = 11
        With CMDMAC_FIX
            .NamedParameters = True
            .CommandType = adCmdStoredProc
            .CommandText = "sp_fix_mac_1"
            .ActiveConnection = gconDMIS
            .CommandTimeout = 1000
        End With
        DoEvents
        CMDMAC_FIX.Execute: DoEvents
        DoEvents
    End If
'fix mac issues
    RSMAC_CTR = (gconDMIS.Execute("Select count(*) from pmis_stockmas where tag_mac = '3'").Fields(0).Value)
    If RSMAC_CTR > 0 Then
        Set CMDMAC_FIX = New ADODB.Command
        Screen.MousePointer = 11
        With CMDMAC_FIX
            .NamedParameters = True
            .CommandType = adCmdStoredProc
            .CommandText = "sp_mac_fixer"
            .ActiveConnection = gconDMIS
            .CommandTimeout = 1000
        End With
        DoEvents
        txtCurrentActivity.Text = txtCurrentActivity.Text + "Fixing Mac........" & vbCrLf
        CMDMAC_FIX.Execute: DoEvents
        DoEvents
    End If
'---------------------------------------------------------------------------------------------
    Dim str_MSG                                        As String
    str_MSG = "Error Appear In During @UTX83912839123" & vbCrLf
    str_MSG = str_MSG & "Parts Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    If MsgQuestionBox("Close All Transactions, Are You Sure?", "Month End Processing") = True Then
        cmdPost.Enabled = False
        cmdExit.Enabled = False
        gconDMIS.BeginTrans

        chkUpdateMaster.Value = 1
        If UpdateMaster = False Then
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Update Master file")
            MsgBox str_MSG, vbCritical, "Month End Error"
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If

        chkBatchPosting.Value = 1
        If BatchPosting = False Then
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Update Master file")
            MsgBox str_MSG, vbCritical, "Month End Error"
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If

        chkMonthEnd.Value = 1
        If MonthEndUpdate = False Then
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Month End Update")
            MsgBox str_MSG, vbCritical, "Month End Error"
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If

        chkGenerateRankFile.Value = 1
        If GenRankFile = False Then
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Month End Update")
            MsgBox str_MSG, vbCritical, "Rank File Error "
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If

        chkCreateStockStatus.Value = 1
        If CreateStockStatus = False Then
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Month End Update")
            MsgBox str_MSG, vbCritical, "Error on Creating Stock Status."
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If

        gconDMIS.CommitTrans

        MsgSpeech "Month End Closing Successfully Completed!": MsgBox "Month End Closing Successfully Completed!", vbInformation, "Completed..."
        LogAudit "G", "Month End Processing"
        cmdExit.Enabled = True

    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Public Function LINEARREGRESSION(ParamArray Values() As Variant) As Double()
    Dim X                                              As Integer
    Dim Y()                                            As Double
    Dim INTLOOP                                        As Integer
    Dim N                                              As Integer

    Dim Q1                                             As Double
    Dim Q2                                             As Double
    Dim Q3                                             As Double

    Dim XSUM                                           As Double
    Dim YSUM                                           As Double
    Dim XSQUAREDSUM                                    As Double
    Dim YSQUAREDSUM                                    As Double
    Dim XYSUM                                          As Double
    Dim LVAL(3)                                        As Double
    X = UBound(Values) + 1
    ReDim Y(1 To X) As Double
    For INTLOOP = 1 To X
        Y(INTLOOP) = Values(INTLOOP - 1)              'Copy values to X
    Next INTLOOP

    For INTLOOP = 1 To X
        XSUM = XSUM + (INTLOOP)
        YSUM = YSUM + Y(INTLOOP)
        XSQUAREDSUM = XSQUAREDSUM + (INTLOOP * INTLOOP)
        YSQUAREDSUM = YSQUAREDSUM + (Y(INTLOOP) * Y(INTLOOP))
        XYSUM = XYSUM + (Y(INTLOOP) * INTLOOP)
    Next INTLOOP

    N = X                                             'Number of periods in calculation
    Q1 = (XYSUM - ((XSUM * YSUM) / N))
    Q2 = (XSQUAREDSUM - ((XSUM * XSUM) / N))
    Q3 = (YSQUAREDSUM - ((YSUM * YSUM) / N))
    LVAL(0) = FormatNumber((Q1 / Q2))                 'Slope
    LVAL(1) = FormatNumber((YSUM - LVAL(0) * XSUM) / N)    'Intercept
    LVAL(2) = FormatNumber(((N + 1) * LVAL(0)) + LVAL(1))    'Forecast
    LINEARREGRESSION = LVAL
End Function

Private Sub LST_UNPOSTED_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
'Dim i As Integer
'Dim cnt As Integer
'cnt = label.Caption
'i = 1
'While i <= 4
'    Row.Record(i).BackColor = &HC0FFC0
'
'    i = i + 2
'Wend
Metrics.BackColor = &HC0FFC0
End Sub

Private Sub LST_UNPOSTED_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record(2).Value = "PO" And Row.Record(3).Value = "P" Then
    
        If Module_Access(LOGID, "PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISTrans_Purchase
            frmPMISTrans_Purchase.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "PO" And Row.Record(3).Value = "A" Then
    
        If Module_Access(LOGID, "ACCESSORIES PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            FormExistsShow frmPMISAC_Purchase
            frmPMISAC_Purchase.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "PO" And Row.Record(3).Value = "M" Then
    
        If Module_Access(LOGID, "MATERIALS PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISMAT_Purchase
            frmPMISMAT_Purchase.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "RR" And Row.Record(3).Value = "P" Then
    
        If Module_Access(LOGID, "PURCHASE RECEIVING STORING", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISTrans_Receiving2
            frmPMISTrans_Receiving2.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "RR" And Row.Record(3).Value = "A" Then
    
    
        If Module_Access(LOGID, "ACCESSORIES RECEIVING", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISTrans_Receiving2_AC
            frmPMISTrans_Receiving2_AC.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "RR" And Row.Record(3).Value = "M" Then
    
       If Module_Access(LOGID, "MATERIALS RECEIVING", "TRANSACTION") = False Then Exit Sub
            FormExistsShow frmPMISTrans_Receiving2_MAT
            frmPMISTrans_Receiving2_MAT.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "PRS" And Row.Record(3).Value = "P" Then
    
       If Module_Access(LOGID, "PARTS REQUISTION SLIP", "TRANSACTION") = False Then Exit Sub
            WAREHOUSETYPE = "PRS"
            frmPMISTrans_PrisForms.txtTranType.Text = "PRS"
            FormExistsShow frmPMISTrans_PrisForms
            frmPMISTrans_PrisForms.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "ARS" And Row.Record(3).Value = "A" Then
    
            On Error Resume Next
        If Module_Access(LOGID, "ACCESSORIES REQUISITION SLIP", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISAC_ARISForms
            WAREHOUSETYPE = "ARS"
            frmPMISAC_ARISForms.txtTranType.Text = "ARS"
            FormExistsShow frmPMISAC_ARISForms
            frmPMISAC_ARISForms.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "MRS" And Row.Record(3).Value = "M" Then
    
            On Error Resume Next
        If Module_Access(LOGID, "MATERIALS REQUISITION SLIP", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISMAT_MRISForms
            WAREHOUSETYPE = "MRS"
            frmPMISMAT_MRISForms.txtTranType.Text = "MRS"
            FormExistsShow frmPMISMAT_MRISForms
            frmPMISMAT_MRISForms.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "CSH" And Row.Record(3).Value = "P" Then
    
        If Module_Access(LOGID, "PARTS ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_ADB_Issuances
            Unload frmPMISTrans_CustomerOrder
            COUNTERTYPE = "CSH"
            frmPMISTrans_CustomerOrder.txtTranType.Text = "CSH"
            FormExistsShow frmPMISTrans_CustomerOrder
            frmPMISTrans_CustomerOrder.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "CHG" And Row.Record(3).Value = "P" Then
    
        If Module_Access(LOGID, "PARTS ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder
            Unload frmPMISTrans_ADB_Issuances
            COUNTERTYPE = "CHG"
            frmPMISTrans_CustomerOrder.txtTranType.Text = "CHG"
            FormExistsShow frmPMISTrans_CustomerOrder
            frmPMISTrans_CustomerOrder.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "RIV" And Row.Record(3).Value = "P" Then
    
        If Module_Access(LOGID, "PARTS ISSUANCE SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder
            Unload frmPMISTrans_ADB_Issuances
            COUNTERTYPE = "RIV"
            frmPMISTrans_CustomerOrder.txtTranType.Text = "RIV"
            FormExistsShow frmPMISTrans_CustomerOrder
            frmPMISTrans_CustomerOrder.textSearch.Text = Row.Record(1).Value
    
    ElseIf Row.Record(2).Value = "CSH" And Row.Record(3).Value = "A" Then
            On Error Resume Next
        If Module_Access(LOGID, "ACCESSORIES ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISTrans_CustomerOrder_AC
            COUNTERTYPE = "CSH"
            frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "CSH"
            FormExistsShow frmPMISTrans_CustomerOrder_AC
            frmPMISTrans_CustomerOrder_AC.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "CHG" And Row.Record(3).Value = "A" Then
            On Error Resume Next
        If Module_Access(LOGID, "ACCESSORIES ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
            Unload frmPMISTrans_CustomerOrder_AC
            COUNTERTYPE = "CHG"
            frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "CHG"
            FormExistsShow frmPMISTrans_CustomerOrder_AC
            frmPMISTrans_CustomerOrder_AC.textSearch.Text = Row.Record(1).Value
     
    ElseIf Row.Record(2).Value = "RIV" And Row.Record(3).Value = "A" Then

        If Module_Access(LOGID, "ACCESSORIES SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder_AC
            COUNTERTYPE = "RIV"
            frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "RIV"
            FormExistsShow frmPMISTrans_CustomerOrder_AC
            frmPMISTrans_CustomerOrder_AC.textSearch.Text = Row.Record(1).Value
            
    ElseIf Row.Record(2).Value = "CSH" And Row.Record(3).Value = "M" Then

            On Error Resume Next
        If Module_Access(LOGID, "MATERIALS ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub

            Unload frmPMISTrans_CustomerOrder_MAT
            COUNTERTYPE = "CSH"
            frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "CSH"
            FormExistsShow frmPMISTrans_CustomerOrder_MAT
            frmPMISTrans_CustomerOrder_MAT.textSearch.Text = Row.Record(1).Value
                
    ElseIf Row.Record(2).Value = "RIV" And Row.Record(3).Value = "M" Then
     
        If Module_Access(LOGID, "MATERIALS SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
            On Error Resume Next
            Unload frmPMISTrans_CustomerOrder_MAT
            COUNTERTYPE = "RIV"
            frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "RIV"
            FormExistsShow frmPMISTrans_CustomerOrder_MAT
            frmPMISTrans_CustomerOrder_MAT.textSearch.Text = Row.Record(1).Value
     
    End If
       
    
End Sub
Sub InitializeRC()
    With LST_UNPOSTED
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
    End With
End Sub

Function RANKING_A() As Boolean
                    If (N2Str2Zero(RSSHIPPING!Prev_Month) + N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3)) > 0 Then
                        If (N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4)) > 0 Then
                            If (N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5)) > 0 Then
                                If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6)) > 0 Then
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            'vRankType = RANK_FAST_MOVING
                                                            RANKING_A = True
                                                        Else
                                                            'vRankType = RANK_FAST_MOVING
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            'vRankType = RANK_FAST_MOVING
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                RANKING_A = False
                                            End If
                                        End If
                                    Else
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                           RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            RANKING_A = False
                                        End If
                                    End If
                                Else
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                RANKING_A = False
                                            End If
                                        End If
                                    Else
                                        RANKING_A = False
                                    End If
                                End If
                            Else
                                If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6)) > 0 Then
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                     If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                     Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                     End If
                                                Else
                                                     RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                RANKING_A = False
                                            End If
                                        End If
                                    Else
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            RANKING_A = False
                                        End If
                                    End If
                                Else
                                   RANKING_A = False
                                End If
                            End If
                        
                        Else
                            If (N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5)) > 0 Then
                                If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6)) > 0 Then
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                   RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                             RANKING_A = True
                                                        Else
                                                             RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                             RANKING_A = True
                                                        Else
                                                             RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                             RANKING_A = True
                                                        Else
                                                             RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                               RANKING_A = False
                                            End If
                                        End If
                                    Else
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            RANKING_A = False
                                        End If
                                    End If
                                Else
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                             RANKING_A = True
                                                        Else
                                                             RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                             RANKING_A = True
                                                        Else
                                                             RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                             RANKING_A = True
                                                        Else
                                                             RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                            RANKING_A = True
                                                        Else
                                                            RANKING_A = True
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                RANKING_A = False
                                            End If
                                        End If
                                    Else
                                        RANKING_A = False
                                    End If
                                End If
                            Else
                                RANKING_A = False
                            End If
                        End If
                    Else
                        If (N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4)) > 0 Then
                                If (N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5)) > 0 Then
                                    If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                             RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                 RANKING_A = True
                                                            Else
                                                                 RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                 RANKING_A = True
                                                            Else
                                                                 RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                 RANKING_A = True
                                                            Else
                                                                 RANKING_A = True
                                                            End If
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                RANKING_A = False
                                            End If
                                        End If
                                    Else
                                        If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                             RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            RANKING_A = False
                                        End If
                                    End If
                                Else
                                    If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                 RANKING_A = True
                                                            Else
                                                                 RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                             If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                             Else
                                                                RANKING_A = False
                                                             End If
                                                        End If
                                                    Else
                                                         RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    RANKING_A = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            RANKING_A = False
                                                        End If
                                                    End If
                                                Else
                                                    If (N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                                        If (N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = True
                                                            End If
                                                        Else
                                                            If (N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                                RANKING_A = True
                                                            Else
                                                                RANKING_A = False
                                                            End If
                                                        End If
                                                    Else
                                                        RANKING_A = False
                                                    End If
                                                End If
                                            Else
                                                 RANKING_A = False
                                            End If
                                        End If
                                    Else
                                        RANKING_A = False
                                    End If
                                End If
                        Else
                            RANKING_A = False
                        End If
                    
                    End If

End Function

Function RANKING_B() As Boolean
                    If (N2Str2Zero(RSSHIPPING!Prev_Month) + N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6)) > 0 Then
                         If (N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                            If (N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                'vRankType = RANK_MEDIUM_MOVING
                                                RANKING_B = True
                                            Else
                                                RANKING_B = True
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = False
                                            End If
                                        End If
                                    Else
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = True
                                            End If
                                        Else
                                            RANKING_B = False
                                        End If
                                    End If
                                Else
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = True
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = False
                                            End If
                                        End If
                                    Else
                                        RANKING_B = False
                                    End If
                                End If
                            Else
                                If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = True
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = False
                                            End If
                                        End If
                                    Else
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = True
                                            End If
                                        Else
                                             RANKING_B = False
                                        End If
                                    End If
                                Else
                                    RANKING_B = False
                                End If
                            End If
                         Else
                            If (N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                                If (N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                    If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = True
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = True
                                                End If
                                            Else
                                                RANKING_B = False
                                            End If
                                        End If
                                    Else
                                        If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = False
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = False
                                                End If
                                            End If
                                        Else
                                            RANKING_B = False
                                        End If
                                    End If
                                Else
                                    If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = True
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = True
                                                End If
                                            Else
                                                RANKING_B = False
                                            End If
                                        End If
                                    Else
                                         RANKING_B = False
                                    End If
                                End If
                            Else
                                RANKING_B = False
                            End If
                         End If
                    Else
                        If (N2Str2Zero(RSSHIPPING!Months_2) + N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7)) > 0 Then
                            If (N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                 RANKING_B = True
                                            Else
                                                 RANKING_B = True
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                 RANKING_B = True
                                            Else
                                                 RANKING_B = False
                                            End If
                                        End If
                                    Else
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                 RANKING_B = True
                                            Else
                                                 RANKING_B = True
                                            End If
                                        Else
                                            RANKING_B = False
                                        End If
                                    End If
                                Else
                                    If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = True
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                RANKING_B = True
                                            Else
                                                RANKING_B = False
                                            End If
                                        End If
                                    Else
                                        RANKING_B = False
                                    End If
                                End If
                            Else
                                If (N2Str2Zero(RSSHIPPING!Months_3) + N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8)) > 0 Then
                                     If (N2Str2Zero(RSSHIPPING!Months_4) + N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9)) > 0 Then
                                        If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = True
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = False
                                                End If
                                            End If
                                        Else
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = True
                                                End If
                                            Else
                                               RANKING_B = False
                                            End If
                                        End If
                                     Else
                                        If (N2Str2Zero(RSSHIPPING!Months_5) + N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10)) > 0 Then
                                            If (N2Str2Zero(RSSHIPPING!Months_6) + N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11)) > 0 Then
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = True
                                                End If
                                            Else
                                                If (N2Str2Zero(RSSHIPPING!Months_7) + N2Str2Zero(RSSHIPPING!Months_8) + N2Str2Zero(RSSHIPPING!Months_9) + N2Str2Zero(RSSHIPPING!Months_10) + N2Str2Zero(RSSHIPPING!Months_11) + N2Str2Zero(RSSHIPPING!months_12)) > 0 Then
                                                    RANKING_B = True
                                                Else
                                                    RANKING_B = False
                                                End If
                                            End If
                                        Else
                                           RANKING_B = False
                                        End If
                                     End If
                                Else
                                    RANKING_B = False
                                End If
                            End If
                        Else
                            RANKING_B = False
                        End If
                    End If
End Function





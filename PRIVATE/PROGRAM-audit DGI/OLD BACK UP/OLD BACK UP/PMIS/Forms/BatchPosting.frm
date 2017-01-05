VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizprogbar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "xpbutton.ocx"
Begin VB.Form frmPMISBatchPosting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Posting"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "BatchPosting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5745
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post"
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
      Left            =   4200
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "BatchPosting.frx":01CA
      MousePointer    =   99  'Custom
      Picture         =   "BatchPosting.frx":031C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Press F11 for Posting By Range"
      Top             =   675
      Width           =   705
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
      Left            =   4950
      MouseIcon       =   "BatchPosting.frx":0641
      MousePointer    =   99  'Custom
      Picture         =   "BatchPosting.frx":0793
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   675
      Width           =   705
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   30
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   750
         Width           =   3615
         Begin VB.Label labProcessing 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   2
            Top             =   -30
            Width           =   3525
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   3
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   4
            Top             =   0
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "BatchPosting.frx":0AF9
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "BatchPosting.frx":0B15
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "BatchPosting.frx":0B31
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
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmPMISBatchPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTDAYTRAN, rsPartMas, rsShipping  As ADODB.Recordset
Attribute rsPartMas.VB_VarUserMemId = 1073938432
Attribute rsShipping.VB_VarUserMemId = 1073938432
Dim rsRR_HD, rsOrd_Hd, rsORD_HIST      As ADODB.Recordset
Attribute rsRR_HD.VB_VarUserMemId = 1073938435
Attribute rsOrd_Hd.VB_VarUserMemId = 1073938435
Attribute rsORD_HIST.VB_VarUserMemId = 1073938435
Dim rsREC_HIST, rsPO_HD, rsPO_HIST     As ADODB.Recordset
Attribute rsREC_HIST.VB_VarUserMemId = 1073938438
Attribute rsPO_HD.VB_VarUserMemId = 1073938438
Attribute rsPO_HIST.VB_VarUserMemId = 1073938438
Dim rsPO_Stat, rsDAYTRAN, rsNOHeader   As ADODB.Recordset
Attribute rsPO_Stat.VB_VarUserMemId = 1073938441
Attribute rsDAYTRAN.VB_VarUserMemId = 1073938441
Attribute rsNOHeader.VB_VarUserMemId = 1073938441
Dim rsNODetail, rsNO_Mstr, rsSupplier  As ADODB.Recordset
Attribute rsNODetail.VB_VarUserMemId = 1073938444
Attribute rsNO_Mstr.VB_VarUserMemId = 1073938444
Attribute rsSupplier.VB_VarUserMemId = 1073938444

Dim vSupplier, vVatAmt, AddSql, upsql  As String
Attribute vSupplier.VB_VarUserMemId = 1073938447
Attribute vVatAmt.VB_VarUserMemId = 1073938447
Attribute AddSql.VB_VarUserMemId = 1073938447
Attribute upsql.VB_VarUserMemId = 1073938447
Dim vTDTranno, vTDPartOrd, vTDTranType As String
Attribute vTDTranno.VB_VarUserMemId = 1073938451
Attribute vTDPartOrd.VB_VarUserMemId = 1073938451
Attribute vTDTranType.VB_VarUserMemId = 1073938451
Dim vTDInOut, vTDStatus                As String
Attribute vTDInOut.VB_VarUserMemId = 1073938454
Attribute vTDStatus.VB_VarUserMemId = 1073938454
Dim vTotTranCost, vMAC                 As Double
Attribute vTotTranCost.VB_VarUserMemId = 1073938456
Attribute vMAC.VB_VarUserMemId = 1073938456
Dim vTDRecNo, vPMRecNo                 As Long
Attribute vTDRecNo.VB_VarUserMemId = 1073938458
Attribute vPMRecNo.VB_VarUserMemId = 1073938458
Dim vPMOnhand, vPMTrecqty, vPMTissqty  As Integer
Attribute vPMOnhand.VB_VarUserMemId = 1073938460
Attribute vPMTrecqty.VB_VarUserMemId = 1073938460
Attribute vPMTissqty.VB_VarUserMemId = 1073938460
Dim vPMLast_Recd, vTDTranDate          As String
Attribute vPMLast_Recd.VB_VarUserMemId = 1073938463
Attribute vTDTranDate.VB_VarUserMemId = 1073938463
Dim vPMReceipts, vPMIssuances, vTDTranQTY, vTDRRNetCost As Integer
Attribute vPMReceipts.VB_VarUserMemId = 1073938465
Attribute vPMIssuances.VB_VarUserMemId = 1073938465
Attribute vTDTranQTY.VB_VarUserMemId = 1073938465
Attribute vTDRRNetCost.VB_VarUserMemId = 1073938465
Dim vTDNetPrice, vTDNetCost, vTDTranucost, vTDRRInvAmt As Double
Attribute vTDNetPrice.VB_VarUserMemId = 1073938469
Attribute vTDNetCost.VB_VarUserMemId = 1073938469
Attribute vTDTranucost.VB_VarUserMemId = 1073938469
Attribute vTDRRInvAmt.VB_VarUserMemId = 1073938469
Dim vORDTotPrice, vTDTranuprice, vTDTranInvAmt As Double
Attribute vORDTotPrice.VB_VarUserMemId = 1073938473
Attribute vTDTranuprice.VB_VarUserMemId = 1073938473
Attribute vTDTranInvAmt.VB_VarUserMemId = 1073938473
Dim vShCurrMonth                       As Integer
Attribute vShCurrMonth.VB_VarUserMemId = 1073938476
Dim vShRecNo                           As Long
Attribute vShRecNo.VB_VarUserMemId = 1073938477
Dim vNetPrice, vNetCost                As Double
Attribute vNetPrice.VB_VarUserMemId = 1073938478
Attribute vNetCost.VB_VarUserMemId = 1073938478
Dim vOrdHDRecNo, vRRHDRecNo, vPOHDRecNo As Long
Attribute vOrdHDRecNo.VB_VarUserMemId = 1073938480
Attribute vRRHDRecNo.VB_VarUserMemId = 1073938480
Attribute vPOHDRecNo.VB_VarUserMemId = 1073938480

Private Sub cmdExit_Click()
    Unload Me
End Sub


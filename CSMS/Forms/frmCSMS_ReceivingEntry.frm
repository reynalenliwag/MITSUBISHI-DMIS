VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMS_ReceivingEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Receiving Entry "
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_ReceivingEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11610
   Visible         =   0   'False
   Begin VB.TextBox txtAPJNO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   7530
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2160
      TabIndex        =   49
      Top             =   3120
      Width           =   9375
      Begin TabDlg.SSTab SSTab1 
         Height          =   3015
         Left            =   0
         TabIndex        =   50
         Top             =   120
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5318
         _Version        =   393216
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Sublet Labor"
         TabPicture(0)   =   "frmCSMS_ReceivingEntry.frx":1082
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstJobSublet"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Materials"
         TabPicture(1)   =   "frmCSMS_ReceivingEntry.frx":109E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstMaterials"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Parts"
         TabPicture(2)   =   "frmCSMS_ReceivingEntry.frx":10BA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstparts"
         Tab(2).ControlCount=   1
         Begin MSComctlLib.ListView lstJobSublet 
            Height          =   2595
            Left            =   60
            TabIndex        =   51
            Top             =   360
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4577
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LineNo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Job Description"
               Object.Width           =   8643
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Sublet Cost"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "LIVIL"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Contractor_Amt"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "WCODE"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lstMaterials 
            Height          =   2595
            Left            =   -74940
            TabIndex        =   52
            Top             =   360
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4577
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LineNo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Job Description"
               Object.Width           =   8643
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Sublet Cost"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "LIVIL"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Contractor_amt"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "WCODE"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lstparts 
            Height          =   2595
            Left            =   -74940
            TabIndex        =   64
            Top             =   360
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4577
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LineNo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Job Description"
               Object.Width           =   8643
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Sublet Cost"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "LIVIL"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Contractor_amt"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "WCODE"
               Object.Width           =   0
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2160
      TabIndex        =   37
      Top             =   -30
      Width           =   9375
      Begin Crystal.CrystalReport rptSubletPo_RR 
         Left            =   120
         Top             =   2100
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   3510
         Top             =   150
      End
      Begin VB.TextBox txtContractorAdd 
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
         ForeColor       =   &H00404040&
         Height          =   735
         Left            =   1410
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1800
         Width           =   3915
      End
      Begin VB.ComboBox cboPoNumber 
         Appearance      =   0  'Flat
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
         Left            =   1410
         TabIndex        =   1
         Text            =   "cboPoNumber"
         Top             =   540
         Width           =   1995
      End
      Begin VB.TextBox txtInvNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   6930
         MaxLength       =   35
         TabIndex        =   8
         Top             =   2580
         Width           =   2295
      End
      Begin VB.TextBox txtDelReceipt 
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
         Left            =   1410
         MaxLength       =   35
         TabIndex        =   7
         Top             =   2580
         Width           =   1995
      End
      Begin VB.TextBox txtNetAmount 
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   6930
         MaxLength       =   35
         TabIndex        =   46
         Top             =   2190
         Width           =   2295
      End
      Begin VB.TextBox txtVatAmount 
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   6930
         MaxLength       =   35
         TabIndex        =   44
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cboContractor 
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
         Left            =   1410
         TabIndex        =   5
         Text            =   "cboContractor"
         Top             =   1380
         Width           =   3915
      End
      Begin VB.TextBox txtReceive 
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
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   2
         Top             =   960
         Width           =   1995
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   5190
         TabIndex        =   4
         Top             =   540
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DateIsNull      =   -1  'True
         Format          =   95617025
         CurrentDate     =   39559
      End
      Begin VB.TextBox txtTotalAmount 
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   6930
         MaxLength       =   35
         TabIndex        =   42
         Top             =   1410
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   5190
         TabIndex        =   3
         Top             =   150
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   95617025
         CurrentDate     =   39559
      End
      Begin VB.TextBox txtRcNumber 
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
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   0
         Top             =   120
         Width           =   1995
      End
      Begin VB.Label labDET 
         Caption         =   "labDET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4080
         TabIndex        =   63
         Top             =   2610
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contractor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   300
         TabIndex        =   66
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   540
         TabIndex        =   65
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblPOSTED 
         Alignment       =   2  'Center
         Caption         =   "POSTED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   7110
         TabIndex        =   62
         Top             =   180
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label labID 
         Caption         =   "labID"
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
         Left            =   3930
         TabIndex        =   61
         Top             =   1050
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "RR DATE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4260
         TabIndex        =   60
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "RR NO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   690
         TabIndex        =   59
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Ref Inv #"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5970
         TabIndex        =   48
         Top             =   2670
         Width           =   870
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ref DR #"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   450
         TabIndex        =   47
         Top             =   2670
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5730
         TabIndex        =   45
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "VaT Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5640
         TabIndex        =   43
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5610
         TabIndex        =   41
         Top             =   1470
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Receive From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   40
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PO NO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   660
         TabIndex        =   39
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PO Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4260
         TabIndex        =   38
         Top             =   630
         Width           =   780
      End
   End
   Begin VB.PictureBox picJobs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   285
      Left            =   2160
      ScaleHeight     =   255
      ScaleWidth      =   9345
      TabIndex        =   31
      Top             =   6300
      Width           =   9375
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   7110
         TabIndex        =   36
         Top             =   30
         Width           =   2445
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4950
         TabIndex        =   35
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Jobs"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3360
         TabIndex        =   34
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Jobs"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1740
         TabIndex        =   33
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Jobs"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   32
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7515
      Left            =   0
      TabIndex        =   9
      Top             =   -30
      Width           =   2115
      Begin VB.TextBox txtSearch 
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
         TabIndex        =   12
         Top             =   960
         Width           =   1995
      End
      Begin VB.OptionButton optContractor 
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   630
         Width           =   1875
      End
      Begin VB.OptionButton optRCNo 
         Caption         =   "RR number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lvwTran 
         Height          =   6105
         Left            =   60
         TabIndex        =   13
         Top             =   1350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   10769
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
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":10D6
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tranno"
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label22 
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
         TabIndex        =   14
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.PictureBox picParts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   285
      Left            =   2160
      ScaleHeight     =   255
      ScaleWidth      =   9345
      TabIndex        =   53
      Top             =   6300
      Width           =   9375
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   58
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1740
         TabIndex        =   57
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3360
         TabIndex        =   56
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   5070
         TabIndex        =   55
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   7110
         TabIndex        =   54
         Top             =   30
         Width           =   2445
      End
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   870
      Left            =   2160
      ScaleHeight     =   870
      ScaleWidth      =   9405
      TabIndex        =   18
      Top             =   6630
      Width           =   9405
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
         Left            =   8580
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":1238
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":138A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
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
         Left            =   7800
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":16F0
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":1842
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelPO 
         Caption         =   "Cancel Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7020
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":1BA8
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6240
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":2034
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":2186
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Unpost this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5460
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":24CB
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":261D
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   795
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
         Left            =   4680
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":2942
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":2A94
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   795
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
         Left            =   3900
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":2DF0
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":2F42
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
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
         Height          =   795
         Left            =   3120
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":3255
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":33A7
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
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
         Height          =   795
         Left            =   2340
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":36F7
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":3849
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   795
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
         Left            =   1560
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":3BA7
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":3CF9
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   795
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
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":3FF3
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":4145
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   795
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
         Left            =   0
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":449D
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":45EF
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   825
      Left            =   9930
      ScaleHeight     =   825
      ScaleWidth      =   1650
      TabIndex        =   15
      Top             =   6630
      Width           =   1650
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
         Left            =   840
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":494E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":4AA0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   795
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
         MouseIcon       =   "frmCSMS_ReceivingEntry.frx":4DDE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ReceivingEntry.frx":4F30
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "APJ NO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   30
      TabIndex        =   69
      Top             =   7530
      Width           =   1005
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2430
      TabIndex        =   68
      Top             =   7530
      Visible         =   0   'False
      Width           =   9135
   End
End
Attribute VB_Name = "frmCSMS_ReceivingEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                          As String
Dim rsINFO                                             As ADODB.Recordset
Dim vSublet_TOTAL_AMT                                  As Double
Dim vSublet_TOTAl_VAT                                  As Double
Dim vSublet_NET_AMT                                    As Double
Dim lstLivil                                           As String
Dim lstLine_No                                         As String
Dim str_MSG                                            As String

'Function CheckIfAlreadyAdded(vLINE_NO As String, vPONO As String) As Boolean
'    Dim rsTmp As New ADODB.Recordset
'    Set rsTmp = gconDMIS.Execute("SELECT LINE_NO,STATUS FROM CSMS_PO_RC_DT WHERE LINE_NO = '" & vLINE_NO & "' and po_no = '" & vPONO & "'")
'    If Not (rsTmp.BOF And rsTmp.EOF) Then
'        If Null2String(rsTmp!Status) = "P" Then
'            CheckIfAlreadyAdded = True
'        ElseIf Null2String(rsTmp!Status) = "C" Then
'            CheckIfAlreadyAdded = False
'        ElseIf Null2String(rsTmp!Status) = "" Then
'            CheckIfAlreadyAdded = True
'        ElseIf Null2String(rsTmp!Status) = "R" Then
'        End If
'    Else
'        CheckIfAlreadyAdded = False
'    End If
'    Set rsTmp = Nothing
'End Function

Function SetContractorAdd(XXX As String) As String
    On Error Resume Next
    Dim rsContractorAdd                                As New ADODB.Recordset
    'Set rsContractorAdd = gconDMIS.Execute("Select * from CSMS_Contractor Where CompanyName = '" & XXX & "'")
    Set rsContractorAdd = gconDMIS.Execute("Select * from ALL_VENDOR_TABLE Where NAMEOFVENDOR = '" & LTrim(RTrim(XXX)) & "' and CODE IS NOT NULL")
    If Not rsContractorAdd.EOF And Not rsContractorAdd.BOF Then
        SetContractorAdd = Null2String(rsContractorAdd!Address)
        txtReceive.Text = Null2String(rsContractorAdd!Code)
    End If
    Set rsContractorAdd = Nothing
End Function

Function passID(XXX As Variant) As Variant
    Call rsRefresh
    rsINFO.Find ("id=" & labID)
    Call StoreMemVars
End Function

Sub initMemvars()
    Dim rsRCNUMBER_Counter                             As New ADODB.Recordset
'    Set rsRCNUMBER_Counter = Nothing
    
    If COMPANY_CODE = "DJM" Then
        rsRCNUMBER_Counter.Open "select isnull(MAx((REPLICATE('0', 6 - LEN( RC_NO)) +  RC_NO)),0)  as RC_NO from CSMS_PO_RC_HD", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsRCNUMBER_Counter.EOF And Not rsRCNUMBER_Counter.BOF Then
            txtRcNumber.Text = Format(NumericVal(Mid$(rsRCNUMBER_Counter!RC_NO, 1, 6)) + 1, "000000")
        Else
            txtRcNumber.Text = "000001"
        End If
        txtRcNumber.Locked = True
    Else
        rsRCNUMBER_Counter.Open "select isnull(MAx((REPLICATE('0', 6 - LEN( RC_NO)) +  RC_NO)),0)  as RC_NO from CSMS_PO_RC_HD", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsRCNUMBER_Counter.EOF And Not rsRCNUMBER_Counter.BOF Then
            txtRcNumber.Text = Format(NumericVal(Mid$(rsRCNUMBER_Counter!RC_NO, 1, 6)) + 1, "000000")
        Else
            txtRcNumber.Text = "000001"
        End If
        txtRcNumber.Locked = False
    End If

    cboPoNumber.ListIndex = -1
    txtReceive.Text = ""
    'cboContractor.ListIndex = -1
    txtContractorAdd.Text = ""
    txtDelReceipt.Text = ""
    txtInvNo.Text = ""
    DTPicker1.Value = Date
    txtTotalAmount.Text = ""
    txtVatAmount.Text = ""
    txtNetAmount.Text = ""
    lblPOSTED.Caption = ""
    DTPicker2.Value = Date
End Sub

Sub initCboContractor()
    Dim rsContractor                                   As New ADODB.Recordset
    'UPDATED BY: JUN
    'DATE UPDATED: 09042008
    'DESCRIPTION: RETRIEVE THE NAME OF VENDOR REGISTERED IN AMIS MODULE
    'Set rsContractor = gconDMIS.Execute("Select * from CSMS_Contractor Order by code asc")
    Set rsContractor = gconDMIS.Execute("Select * from ALL_VENDOR_TABLE where CODE IS NOT NULL Order by code asc")
    If Not rsContractor.EOF And Not rsContractor.BOF Then
        rsContractor.MoveFirst: cboContractor.Clear
        Do While Not rsContractor.EOF
            cboContractor.AddItem Null2String(rsContractor!nameofvendor)
            rsContractor.MoveNext
        Loop
    End If
    Set rsContractor = Nothing
End Sub

'Sub initCboPoNumber()
'    Dim rsPoNumber As ADODB.Recordset
'    Set rsPoNumber = New ADODB.Recordset
'        Set rsPoNumber = gconDMIS.Execute("Select Po_No from CSMS_PO_HD where Status = 'P'  Order by PO_NO asc")
'               If Not rsPoNumber.EOF And Not rsPoNumber.BOF Then
'                    rsPoNumber.MoveFirst: cboPoNumber.Clear
'                    Do While Not rsPoNumber.EOF
'                       Dim rsstatus As ADODB.Recordset
'                       Set rsstatus = gconDMIS.Execute("Select STATUS from CSMS_PO_RC_HD where PO_NO ='" & rsPoNumber!PO_NO & "'")
'                       If Not rsstatus.EOF And Not rsstatus.BOF Then
'                            If Null2String(rsstatus!Status) = "P" Or Null2String(rsstatus!Status) = "" Or Null2String(rsstatus!Status) = "R" Then
'                                 rsPoNumber.MoveNext
'                            End If
'                        Else
'                            cboPoNumber.AddItem Null2String(rsPoNumber!PO_NO)
'                            rsPoNumber.MoveNext
'                        End If
'                    Loop
'                End If
'    Set rsPoNumber = Nothing
'    Set rsstatus = Nothing
'End Sub

Sub initCboPoNumber()
    Dim rsPoNumber                                     As New ADODB.Recordset
    
'    Set rsPoNumber = gconDMIS.Execute("Select Po_No from CSMS_PO_HD where Status = 'P'  Order by PO_NO asc")
'    DoEvents
'    If Not rsPoNumber.EOF And Not rsPoNumber.BOF Then
'        rsPoNumber.MoveFirst: cboPoNumber.Clear
'        DoEvents
'        Do While Not rsPoNumber.EOF
'            Dim rsstatus                               As ADODB.Recordset
'            Set rsstatus = gconDMIS.Execute("Select top 100 STATUS from CSMS_PO_RC_HD where PO_NO ='" & rsPoNumber!PO_NO & "' AND (STATUS = 'P' or STATUS = 'R' or STATUS is NULL)")
'            If Not rsstatus.EOF And Not rsstatus.BOF Then
'                'If Null2String(rsstatus!Status) = "P" Or Null2String(rsstatus!Status) = "R" Then
'                '    'rsPoNumber.MoveNext
'                'End If
'            Else
'                cboPoNumber.AddItem Null2String(rsPoNumber!PO_NO)
'            End If
'            'cboPoNumber.AddItem Null2String(rsPoNumber!PO_NO)
'            rsPoNumber.MoveNext
'        Loop
'    End If
'    Set rsPoNumber = Nothing
'    Set rsstatus = Nothing

'updated by: IEBV
'description: For faster loading of PO
    Set rsPoNumber = gconDMIS.Execute("Select Po_No from CSMS_PO_HD where Status = 'P' and PO_NO not in ( " & _
                                      " Select PO_NO from CSMS_PO_RC_HD where isnull(STATUS,'') IN ('P','R','') ) " & _
                                      " Order by PO_NO asc")
    If Not rsPoNumber.EOF And Not rsPoNumber.BOF Then
        rsPoNumber.MoveFirst: cboPoNumber.Clear
        DoEvents
        Do While Not rsPoNumber.EOF
            cboPoNumber.AddItem Null2String(rsPoNumber!PO_NO)
            rsPoNumber.MoveNext
        Loop
    End If
    Set rsPoNumber = Nothing
End Sub

Sub rsRefresh()
    Set rsINFO = New ADODB.Recordset
    rsINFO.Open "select * from CSMS_PO_RC_HD order by RC_NO DESC", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub FillInfo()
    Dim rsRetreivePO_HD                                As New ADODB.Recordset

    Set rsRetreivePO_HD = gconDMIS.Execute("Select CONTRACTOR_CODE,Contractor_Name,Contractor_Address,Po_Date,SUBLET_TOTAL_AMT,SUBLET_TOTAL_VAT,SUBLET_TOTAL_NET_AMT from CSMS_PO_HD where STATUS = 'P' and Po_No ='" & cboPoNumber.Text & "'")
    DoEvents
    If Not rsRetreivePO_HD.EOF And Not rsRetreivePO_HD.BOF Then
        txtReceive.Text = Null2String(rsRetreivePO_HD!Contractor_Code)
        cboContractor.Text = Null2String(rsRetreivePO_HD!Contractor_Name)
        txtContractorAdd.Text = Null2String(rsRetreivePO_HD!Contractor_Address)
        DTPicker1.Value = rsRetreivePO_HD!Po_Date
        txtTotalAmount.Text = Format((N2Str2Zero(rsRetreivePO_HD!Sublet_TOTAL_AMT)), MAXIMUM_DIGIT)
        txtVatAmount.Text = Format((N2Str2Zero(rsRetreivePO_HD!Sublet_TOTAl_VAT)), MAXIMUM_DIGIT)
        txtNetAmount.Text = Format((N2Str2Zero(rsRetreivePO_HD!SUBLET_TOTAL_NET_AMT)), MAXIMUM_DIGIT)

        Dim Item                                       As ListItem
        Dim rsRC_PO_dt                                 As New ADODB.Recordset

        Me.lstJobSublet.Sorted = True: Me.lstJobSublet.ListItems.Clear: Me.lstJobSublet.Enabled = False
        Set rsRC_PO_dt = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,DETAIL from CSMS_Po_Dt where Po_no = '" & cboPoNumber.Text & "' and livil = '1' order by LINE_NO asc")

        If Not rsRC_PO_dt.EOF And Not rsRC_PO_dt.BOF Then
            Do While Not rsRC_PO_dt.EOF
                Set Item = lstJobSublet.ListItems.Add(, , Null2String(rsRC_PO_dt!LINE_NO))
                Item.SubItems(1) = Null2String(rsRC_PO_dt!DETCDE)
                Item.SubItems(2) = Null2String(rsRC_PO_dt!Detail)
                Item.SubItems(3) = Format(NumericVal(rsRC_PO_dt!DET_AMT), MAXIMUM_DIGIT)
                Item.SubItems(4) = Null2String(rsRC_PO_dt!ID)

                rsRC_PO_dt.MoveNext
            Loop
            Me.lstJobSublet.Enabled = True: Me.lstJobSublet.Sorted = False: Me.lstJobSublet.Refresh
        End If
    End If
End Sub

Sub StoreMemVars()
    If Not rsINFO.EOF And Not rsINFO.BOF Then
        labID.Caption = rsINFO!ID
        txtRcNumber.Text = Null2String(rsINFO!RC_NO)
        cboPoNumber.Text = Null2String(rsINFO!PO_NO)
        txtReceive.Text = Null2String(rsINFO!Contractor_Code)
        DTPicker1.Value = rsINFO!Po_Date
        DTPicker2.Value = rsINFO!RC_DATE
        cboContractor.Text = Null2String(rsINFO!Contractor_Name)
        txtContractorAdd.Text = Null2String(rsINFO!Contractor_Address)
        txtInvNo.Text = Null2String(rsINFO!invoice_no)
        txtDelReceipt.Text = Null2String(rsINFO!delivery_or)
        txtTotalAmount.Text = Format(NumericVal(rsINFO!Sublet_TOTAL_AMT), MAXIMUM_DIGIT)
        txtVatAmount.Text = Format(NumericVal(rsINFO!Sublet_TOTAl_VAT), MAXIMUM_DIGIT)
        txtNetAmount.Text = Format(NumericVal(rsINFO!SUBLET_TOTAL_NET_AMT), MAXIMUM_DIGIT)

        Call FillListview(txtRcNumber)
        Call getAPJ_NO(Trim(txtReceive), Trim(txtRcNumber))

        If Null2String(rsINFO!Status) = "P" Then
            cmdPrint.Enabled = True
            cmdPost.Enabled = False
            cmdUnPost.Enabled = True
            cmdEdit.Enabled = False
            cmdCancelPO.Enabled = False
            lblPOSTED = "**POSTED**"
        ElseIf Null2String(rsINFO!Status) = "C" Then
            cmdEdit.Enabled = False
            cmdPrint.Enabled = False
            cmdPost.Enabled = False
            cmdCancelPO.Enabled = False
            cmdUnPost.Enabled = False
            lblPOSTED = "**CANCELLED**"
        Else
            cmdEdit.Enabled = True
            cmdPrint.Enabled = False
            cmdPost.Enabled = True
            cmdUnPost.Enabled = False
            cmdCancelPO.Enabled = True
            lblPOSTED = ""
        End If
    Else
        Call ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub FilllstPO_HD()
    Dim i                                              As Integer
    Listview_Loadval lvwTran.ListItems, gconDMIS.Execute("Select TOP 30 RC_NO,ID  from CSMS_PO_RC_HD  order by RC_NO desc")
End Sub

Sub passINFO()
    Dim rsCustomer                                     As ADODB.Recordset

    Set rsCustomer = gconDMIS.Execute("Select RO_no,Cust_Name from CSMS_PO_HD where Po_No = '" & cboPoNumber & "'")
    frmCSMS_SubletJob.txtCustomer.Text = Null2String(rsCustomer!Cust_name)
    frmCSMS_SubletJob.txtROno.Text = Null2String(rsCustomer!RO_NO)
End Sub

Sub editJobs()
    Dim rsEditJobs                                     As New ADODB.Recordset
    Dim vTypeOfJob                                     As String
    Dim rsCustname                                     As New ADODB.Recordset

    On Error Resume Next

    Set rsCustname = gconDMIS.Execute("Select Cust_name From CSMS_PO_HD where PO_NO = '" & cboPoNumber & "'")
    Set rsEditJobs = gconDMIS.Execute("Select Rep_or,ROTYPE,JOBTYPE,DETCDE,DETDSC,DET_AMT,CONTRACTAMOUNT,COMPAMOUNT,WCODE,DETAIL,SUBLET_TYPE from CSMS_PO_RC_DT where ID ='" & labDET & "'")

    If Trim(rsEditJobs!DETCDE) = "SRLABOR" Then
        vTypeOfJob = "SUBLET LABOR"
    ElseIf Trim(rsEditJobs!DETCDE) = "SRPARTS" Then
        vTypeOfJob = "SUBLET PARTS"
    Else
        vTypeOfJob = "SUBLET MATERIALS"
    End If


    If Not rsEditJobs.EOF And Not rsEditJobs.BOF Then
        With frmCSMS_SubletJob
            .txtCustomer = Null2String(rsCustname!Cust_name)
            .txtROno = Null2String(rsEditJobs!REP_OR)
            .txtOPCODE = Null2String(rsEditJobs!DETCDE)
            .txtJobDesc = Null2String(rsEditJobs!DETDSC)
            .txtSubletAmount = Format(NumericVal(rsEditJobs!DET_AMT), MAXIMUM_DIGIT)
            .txtContracAmount = Format(NumericVal(rsEditJobs!CONTRACTAMOUNT), MAXIMUM_DIGIT)
            .txtCompAmount = Format(NumericVal(rsEditJobs!COMPAMOUNT), MAXIMUM_DIGIT)
            .cboJobChargeTo = Null2String(rsEditJobs!wCode)
            .txtNote = Null2String(rsEditJobs!Detail)
            .labDET = labDET.Caption
            .cboSubletCategory = vTypeOfJob
            .cboBPorGJ = Null2String(rsEditJobs!JOBTYPE)
            
            If COMPANY_CODE = "CMC" Or COMPANY_CODE = "DSSC" Then
                   If Null2String(rsEditJobs!SUBLET_TYPE) = "T" Then
                           .cbosublettype.Text = "Tinsmith"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "P" Then
                           .cbosublettype.Text = "Painting"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "A" Then
                           .cbosublettype.Text = "Aircon"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "U" Then
                           .cbosublettype.Text = "Undercoating"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "D" Then
                           .cbosublettype.Text = "Detailing"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "S" Then
                           .cbosublettype.Text = "Sublet"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "O" Then
                           .cbosublettype.Text = "Towing"
                       End If
                       
                 End If
                 
                'UPDATE BY : MJP 07222008
                If UCase(.cboSubletCategory.Text) = "SUBLET LABOR" Then
                    If Null2String(rsEditJobs!JOBTYPE) = "BP" Then
                        .cboBP_TYPE.Visible = True
                        .Label4.Visible = True
                        If Null2String(rsEditJobs!transtatus) = "M" Then
                            .cboBP_TYPE.Text = "Major"
                        Else
                            .cboBP_TYPE.Text = "Minor"
                        End If
                    Else
                        .cboBP_TYPE.Visible = False
                        .Label4.Visible = False
                    End If
                End If
            
        End With
    End If
    Set rsEditJobs = Nothing
End Sub

Function EnabledFrame(COND As Boolean)
    Picture1.Enabled = COND
    Frame3.Enabled = COND
    Frame2.Enabled = COND
End Function

Sub deleteJobs()
    Dim rsDelJob                                       As New ADODB.Recordset
    Dim ans                                            As String

    'check for if Po is already posted in creation of PO
    '    Dim MINNIE As ADODB.Recordset
    '    Set MINNIE = gconDMIS.Execute("Select PO_NO from CSMS_PO_HD  where PO_NO ='" & cboPoNumber & "' and Status ='P'")
    '    If Not MINNIE.EOF And Not MINNIE.BOF Then
    '        MsgBox ("Cannot Delete this Job." & vbCrLf & "It's currently working."), vbInformation, "INFORMATION"
    '        Exit Sub
    '    End If

    Dim RS_RR_RONUM                                    As New ADODB.Recordset
    Dim RR_RONUM                                       As String
    Dim RR_LINENO                                      As String
    Dim RR_LIVIL                                       As String
    Dim RR_PONUM                                       As String

    Set RS_RR_RONUM = gconDMIS.Execute("SELECT PO_NO,REP_OR,LIVIL,LINE_NO FROM CSMS_PO_RC_DT WHERE RC_NO ='" & txtRcNumber & "' and ID = '" & labDET & "'")

    If Not RS_RR_RONUM.EOF And Not RS_RR_RONUM.BOF Then
        RR_PONUM = Null2String(RS_RR_RONUM!PO_NO)
        RR_RONUM = Null2String(RS_RR_RONUM!REP_OR)
        RR_LIVIL = Null2String(RS_RR_RONUM!LIVIL)
        RR_LINENO = Null2String(RS_RR_RONUM!LINE_NO)
    End If

    Dim rsPosted                                       As New ADODB.Recordset
    Set rsPosted = gconDMIS.Execute("Select * from CSMS_PO_Dt  where PO_NO = '" & RR_PONUM & "' and LINE_NO = '" & RR_LINENO & "' and LIVIL = '" & RR_LIVIL & "' and STATUS = 'P'")

    If Not rsPosted.EOF And Not rsPosted.BOF Then
        MsgBox "Cannot Delete this Job it's already Posted" & vbCrLf & "In Purchase Order Data Entry", vbOKOnly + vbInformation, "INFORMATION"
        Exit Sub
    End If

    If MsgBox("Are you sure do you want to DELETE this Job?", vbQuestion + vbYesNo) = vbYes Then
        SQL_STATEMENT = "Delete from CSMS_ro_Det where rep_or = '" & RR_RONUM & "' and ROTYPE = 'SR' and SUBPOCODE = '" & cboPoNumber & "' and LIVIL = '" & RR_LIVIL & "' and LINE_NO= '" & RR_LINENO & "' "
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "Delete from CSMS_PO_DT where rep_or = '" & RR_RONUM & "' and LIVIL = '" & RR_LIVIL & "' and LINE_NO = '" & RR_LINENO & "'"
        gconDMIS.Execute SQL_STATEMENT

        gconDMIS.Execute ("Delete from CSMS_PO_RC_DT where ID = '" & labDET & "' and livil ='" & lstLivil & "'")

        Dim rsComputeTotalCost                         As New ADODB.Recordset
        Set rsComputeTotalCost = gconDMIS.Execute("Select DETAMT,TAXVAL,DET_AMT from CSMS_PO_RC_DT where RC_NO ='" & txtRcNumber & "'")

        vSublet_TOTAL_AMT = 0
        vSublet_TOTAl_VAT = 0
        vSublet_NET_AMT = 0

        If Not rsComputeTotalCost.EOF And Not rsComputeTotalCost.BOF Then
            Do While Not rsComputeTotalCost.EOF
                vSublet_TOTAL_AMT = N2Str2Zero(rsComputeTotalCost!DETAMT) + N2Str2Zero(vSublet_TOTAL_AMT)
                vSublet_TOTAl_VAT = N2Str2Zero(rsComputeTotalCost!TAXVAL) + N2Str2Zero(vSublet_TOTAl_VAT)
                vSublet_NET_AMT = N2Str2Zero(rsComputeTotalCost!DET_AMT) + N2Str2Zero(vSublet_NET_AMT)
                rsComputeTotalCost.MoveNext
            Loop
        End If

        gconDMIS.Execute "Update CSMS_PO_RC_HD set " & _
            "SUBLET_TOTAL_AMT = " & vSublet_TOTAL_AMT & "," & _
            "SUBLET_TOTAL_VAT = " & vSublet_TOTAl_VAT & "," & _
            "SUBLET_TOTAL_NET_AMT = " & vSublet_NET_AMT & " " & _
            "where RC_NO =" & txtRcNumber

        gconDMIS.Execute "Update CSMS_PO_HD set " & _
            "SUBLET_TOTAL_AMT = " & vSublet_TOTAL_AMT & "," & _
            "SUBLET_TOTAL_VAT = " & vSublet_TOTAl_VAT & "," & _
            "SUBLET_TOTAL_NET_AMT = " & vSublet_NET_AMT & " " & _
            "where PO_NO =" & RR_PONUM
        
        Call ShowDeletedMsg
    Else
        labDET.Caption = ""
        Exit Sub
    End If

    Call rsRefresh
    rsINFO.Find ("ID =" & labID)
    Call StoreMemVars
    Set rsComputeTotalCost = Nothing
    Set RS_RR_RONUM = Nothing
    Set rsPosted = Nothing
End Sub

Sub FillListview(XXX As String)
    Dim Item                                           As ListItem
    Dim rsPO_dt                                        As New ADODB.Recordset

    'LABOR
    Me.lstJobSublet.Sorted = True: Me.lstJobSublet.ListItems.Clear: Me.lstJobSublet.Enabled = False
    Set rsPO_dt = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_PO_RC_DT where RC_NO = " & XXX & " and livil = '1' and Status <> 'C' order by LINE_NO asc")

    If Not rsPO_dt.EOF And Not rsPO_dt.BOF Then
        Do While Not rsPO_dt.EOF
            Set Item = lstJobSublet.ListItems.Add(, , Null2String(rsPO_dt!LINE_NO))
            Item.SubItems(1) = Null2String(rsPO_dt!DETCDE)
            Item.SubItems(2) = Null2String(rsPO_dt!Detail)
            Item.SubItems(3) = Format(NumericVal(rsPO_dt!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(4) = Null2String(rsPO_dt!ID)
            Item.SubItems(5) = Null2String(rsPO_dt!LIVIL)
            Item.SubItems(6) = Null2String(rsPO_dt!CONTRACTAMOUNT)
            Item.SubItems(7) = Null2String(rsPO_dt!wCode)
            rsPO_dt.MoveNext
        Loop
        Me.lstJobSublet.Enabled = True: Me.lstJobSublet.Sorted = False: Me.lstJobSublet.Refresh
    End If

    Set rsPO_dt = Nothing

    'for materials
    Dim rsMaterials                                    As New ADODB.Recordset

    Me.lstMaterials.Sorted = True: Me.lstMaterials.ListItems.Clear: Me.lstMaterials.Enabled = False
    Set rsMaterials = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_PO_RC_DT where RC_NO = " & XXX & " and livil = '3' order by LINE_NO asc")

    If Not rsMaterials.EOF And Not rsMaterials.BOF Then
        Do While Not rsMaterials.EOF
            Set Item = lstMaterials.ListItems.Add(, , Null2String(rsMaterials!LINE_NO))
            Item.SubItems(1) = Null2String(rsMaterials!DETCDE)
            Item.SubItems(2) = Null2String(rsMaterials!Detail)
            Item.SubItems(3) = Format(NumericVal(rsMaterials!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(4) = Null2String(rsMaterials!ID)
            Item.SubItems(5) = Null2String(rsMaterials!LIVIL)
            Item.SubItems(6) = Null2String(rsMaterials!CONTRACTAMOUNT)
            Item.SubItems(7) = Null2String(rsMaterials!wCode)
            rsMaterials.MoveNext
        Loop
        Me.lstMaterials.Enabled = True: Me.lstMaterials.Sorted = False: Me.lstMaterials.Refresh
    End If

    Set rsMaterials = Nothing

    Dim rsParts                                        As New ADODB.Recordset

    'for parts
    Me.lstparts.Sorted = True: Me.lstparts.ListItems.Clear: Me.lstparts.Enabled = False
    Set rsParts = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_PO_RC_DT where RC_NO = " & XXX & " and livil = '2' order by LINE_NO asc")

    If Not rsParts.EOF And Not rsParts.BOF Then
        Do While Not rsParts.EOF
            Set Item = lstparts.ListItems.Add(, , Null2String(rsParts!LINE_NO))
            Item.SubItems(1) = Null2String(rsParts!DETCDE)
            Item.SubItems(2) = Null2String(rsParts!Detail)
            Item.SubItems(3) = Format(NumericVal(rsParts!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(4) = Null2String(rsParts!ID)
            Item.SubItems(5) = Null2String(rsParts!LIVIL)
            Item.SubItems(6) = Null2String(rsParts!CONTRACTAMOUNT)
            Item.SubItems(7) = Null2String(rsParts!wCode)
            rsParts.MoveNext
        Loop
        Me.lstparts.Enabled = True: Me.lstparts.Sorted = False: Me.lstparts.Refresh
    End If

    Set rsParts = Nothing
End Sub

Sub UnpostDelete()
    Dim rsUnpostDelete                                 As New ADODB.Recordset
    
    If MsgBox("Do You Want to Un Post this Transaction", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    SQL_STATEMENT = "update CSMS_PO_RC_HD set STATUS = 'R' WHERE ID = " & labID
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("U", "SUBLET RECEIVING", SQL_STATEMENT, labID, "", "RR NO: " & txtRcNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "update CSMS_PO_RC_dt set STATUS = 'R' WHERE RC_NO=" & txtRcNumber & ""
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UU", "SUBLET RECEIVING", SQL_STATEMENT, labID, "", "RR NO: " & txtRcNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call rsRefresh
    rsINFO.Find ("id=" & labID)
    Call StoreMemVars
End Sub

Private Sub cboContractor_Change()
    txtContractorAdd.Text = SetContractorAdd(cboContractor)
End Sub

Private Sub cboContractor_Click()
    txtContractorAdd.Text = SetContractorAdd(Trim(cboContractor))
End Sub

Private Sub cboContractor_LostFocus()
    cboContractor.Text = SetContractorName(txtReceive)
End Sub

Function SetContractorName(XXX As String) As String
    Dim rsContractorName                               As New ADODB.Recordset
    'Set rsContractorAdd = gconDMIS.Execute("Select * from CSMS_Contractor Where CompanyName = '" & Repleys(XXX) & "'")
    Set rsContractorName = gconDMIS.Execute("Select * from ALL_VENDOR_TABLE Where Code = '" & Repleys(XXX) & "' AND CODE IS NOT NULL")
    If Not rsContractorName.EOF And Not rsContractorName.BOF Then
        SetContractorName = Null2String(rsContractorName!nameofvendor)
    End If
    Set rsContractorName = Nothing
End Function

Private Sub cboPoNumber_Change()
    Call FillInfo
End Sub

Private Sub cboPoNumber_Click()
    Call FillInfo
End Sub

Private Sub cboPoNumber_DropDown()
    Call initCboPoNumber
End Sub

Private Sub cboPoNumber_LostFocus()
    Dim rsCheckPo_IsInvoice                            As New ADODB.Recordset
    Dim rsAlreadyInvoice                               As New ADODB.Recordset

    If COMPANY_CODE = "HSB" And COMPANY_CODE = "HGC" Then
        Set rsCheckPo_IsInvoice = gconDMIS.Execute("Select RO_NO from CSMS_PO_HD where PO_NO = '" & cboPoNumber & "'")
        If Not rsCheckPo_IsInvoice.EOF And Not rsCheckPo_IsInvoice.BOF Then
            Set rsAlreadyInvoice = gconDMIS.Execute("Select * from CSMS_Repor where INVOICE IS NOT NULL AND REP_OR = '" & Null2String(rsCheckPo_IsInvoice!RO_NO) & "'")
            If Not rsAlreadyInvoice.EOF And Not rsAlreadyInvoice.BOF Then
                Call FillInfo
            Else
                MsgBox "This PO NO: '" & cboPoNumber.Text & "' & is not yet invoice..." & vbCrLf & "It must be invoice first before you receive.", vbInformation, "INFORMATION"
                cmdCancel.Value = True
                Exit Sub
            End If
        End If
    Else
        Call FillInfo
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "ACESS_ADD", "SUBLET RECEIVING") = False Then Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    lstJobSublet.ListItems.Clear
    lstparts.ListItems.Clear
    lstMaterials.ListItems.Clear
    Call initMemvars
    Call initCboPoNumber
    Call initCboContractor
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Frame2.Enabled = True
    Frame3.Enabled = True
    Picture1.Visible = True
    'UPDATE BY: JUN-------------
    'DATE UPDATED: 10152008 208
    cboPoNumber.Enabled = True
    'UPDATE BY: JUN-------------
    Picture2.Visible = False
    txtSearch.Enabled = True
    Call StoreMemVars
    txtRcNumber.Locked = True
End Sub

Private Sub cmdCancelPO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "SUBLET RECEIVING") = False Then Exit Sub

    If MsgBox("Do You Want to Cancel this Transaction", vbInformation + vbYesNo) = vbNo Then Exit Sub

    SQL_STATEMENT = "update CSMS_PO_RC_HD set STATUS = 'C'  WHERE ID = " & labID
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("C", "SUBLET RECEIVING", SQL_STATEMENT, labID, "", "RR NO: " & txtRcNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Update CSMS_PO_RC_dt set STATUS = 'C' WHERE RC_NO = " & txtRcNumber & ""
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("CC", "SUBLET RECEIVING", SQL_STATEMENT, labID, "", "RR NO: " & txtRcNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call rsRefresh
    rsINFO.Find ("ID = " & labID)
    Call StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "ACESS_EDIT", "SUBLET RECEIVING") = False Then Exit Sub

    If CheckIfImported((RTrim(LTrim(txtReceive))), (RTrim(LTrim(txtInvNo)))) = True Then
        MsgBox "You can't edit this transaction" & vbCrLf & "It's Already Posted in accounting", vbInformation, "INFORMATION"
        Exit Sub
    Else
        'proceed to edit
    End If
    
    'txttest = labID
    If COMPANY_CODE = "DJM" Then
        txtRcNumber.Locked = True
    Else
        txtRcNumber.Locked = False
    End If
    
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    'UPDATED BY: JUN-------------------------------------------------------------------------------------------------------------------------------------------------------------
    'DATE UPDATED: 1010152008 204
    'DESCRIPTION: RESTRICT THE USER TO SELECT ANOTHER PO_NO IN EDIT MODE BECAUSE IT WILL AFFECT THE OTHER PO_NO TRANSACTION AND WILL AFFECT THE DATE BEING IMPORTED IN ACCOUNTING
    cboPoNumber.Enabled = False
    'UPDATED BY: JUN-------------------------------------------------------------------------------------------------------------------------------------------------------------
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSearch.BackColor = vbYellow
    txtSearch.Text = ""
    txtSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    rsINFO.MoveFirst
    Call ShowFirstRecordMsg
    Call StoreMemVars
End Sub

Private Sub cmdLast_Click()
    rsINFO.MoveLast
    Call ShowLastRecordMsg
    Call StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsINFO.MoveNext
    If rsINFO.EOF Then
        rsINFO.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "ACESS_POST", "SUBLET RECEIVING") = False Then Exit Sub

    If lstJobSublet.ListItems.Count = 0 And lstMaterials.ListItems.Count = 0 And lstparts.ListItems.Count = 0 Then
        MsgBox ("You cannot post this Transaction... There is no Job Selected."), vbOKOnly, "Information"
        Exit Sub
    End If

    If MsgBox("Do You Want to Post this Transaction", vbInformation + vbYesNo) = vbNo Then Exit Sub
    SQL_STATEMENT = "update CSMS_PO_RC_HD set STATUS = 'P' WHERE ID = " & labID
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("P", "SUBLET RECEIVING", SQL_STATEMENT, labID, "", "RR NO: " & txtRcNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Update CSMS_PO_RC_Dt set STATUS = 'P' WHERE RC_NO = " & txtRcNumber & ""
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("PP", "SUBLET RECEIVING", SQL_STATEMENT, labID, "", "RR NO: " & txtRcNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call rsRefresh
    rsINFO.Find ("ID = " & labID)
    Call StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsINFO.MovePrevious
    If rsINFO.BOF Then
        rsINFO.MoveFirst
        Call ShowFirstRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "ACESS_PRINT", "SUBLET PURCHASE") = False Then Exit Sub

    Screen.MousePointer = 11

    rptSubletPo_RR.Formulas(0) = "Company Name = '" & COMPANY_NAME & "'"
    rptSubletPo_RR.Formulas(1) = "Company Address = '" & COMPANY_ADDRESS & "'"
    rptSubletPo_RR.Formulas(2) = "G_M = '" & GENERAL_MANAGER & "'"

    rptSubletPo_RR.ReportTitle = "Purchase Order Receiving "
    rptSubletPo_RR.WindowTitle = "Purchase Order Receiving"
    PrintSQLReport rptSubletPo_RR, CSMS_REPORT_PATH & "SubletPO_RR.rpt", "{CSMS_PO_RC_DT.RC_NO} = '" & txtRcNumber & "' and {CSMS_PO_RC_DT.STATUS} <> 'C' ", CSMS_REPORT_CONNECTION, 1

    Screen.MousePointer = 0
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SUBLET PURCHASE", "", labID, "", "PO NO: " & txtRcNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
End Sub
Private Sub cmdSave_Click()
   
    
    str_MSG = "Error Appear In During @UTX83912839123" & vbCrLf
    str_MSG = str_MSG & "Service Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
    str_MSG = str_MSG & "Telephone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)"
    
    gconDMIS.BeginTrans
    If SAVERECEVING = False Then
        str_MSG = Replace(str_MSG, "@UTX83912839123", "Recieving PO")
        MsgBox str_MSG, vbCritical, "Saving Error"
        cmdCancel.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    gconDMIS.CommitTrans
    If AddorEdit = "ADD" Then
        Call ShowSuccessFullyAdded
    Else
        Call ShowSuccessFullyUpdated
    End If
    
    Screen.MousePointer = 0

End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "ACESS_UNPOST", "SUBLET RECEIVING") = False Then Exit Sub
    If CheckIfImported((RTrim(LTrim(txtReceive))), (RTrim(LTrim(txtInvNo)))) = True Then
        MsgBox "You can't unpost this transaction" & vbCrLf & "It's Already Posted in accounting", vbInformation, "INFORMATION"
        Exit Sub
    End If
    If txtAPJNO.Text <> "" Then
        MsgBox "You can't unpost this transaction" & vbCrLf & "It's already imported in accounting", vbInformation, "INFORMATION"
        Exit Sub
    End If
    Call UnpostDelete
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    
    Call rsRefresh
    Call initMemvars
    Call initCboContractor
    Call StoreMemVars
    Call FilllstPO_HD
    Call optRCNo_Click
    Timer1.Enabled = True
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

Private Sub lstJobSublet_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lstJobSublet.ListItems.Count = 0 Then Exit Sub
    labDET.Caption = lstJobSublet.SelectedItem.SubItems(4)
End Sub

Private Sub lstJobSublet_DblClick()
    If lstJobSublet.ListItems.Count = 0 Then
        MsgBox "This transaction has no Job to edit.", vbInformation, "INFORMATION"
        Exit Sub
    End If
    If CheckIfImported((RTrim(LTrim(txtReceive))), (RTrim(LTrim(txtInvNo)))) = True Then
        MsgBox "You can't edit this transaction" & vbCrLf & "It's Already Posted in accounting", vbInformation, "INFORMATION"
        Exit Sub
    Else
        'proceed to edit
    End If

    If Picture1.Visible = True Then
        If Null2String(rsINFO!Status) = "C" Then
            MsgBox "Purchase order already Cancelled. Cannot Edit this Job.", vbInformation, "INFORMATION"
            Exit Sub

        ElseIf Null2String(rsINFO!Status) = "P" Then
            MsgBox "Purchase order already posted. Cannot Edit this Job.", vbInformation, "INFORMATION"
            Exit Sub
        Else
            frmCSMS_SubletJob.LINE_NO.Caption = lstLine_No
            Call editJobs
            frmCSMS_SubletJob.lblAddorEdit.Caption = "REDIT"
            Call EnabledFrame(False)
            frmCSMS_SubletJob.txtSubletAmount.Enabled = False
            frmCSMS_SubletJob.cbosublettype.Enabled = False
            frmCSMS_SubletJob.Show
        End If
    End If
End Sub

Private Sub lstJobSublet_Click()
    If lstJobSublet.ListItems.Count = 0 Then Exit Sub
    labDET.Caption = (lstJobSublet.SelectedItem.SubItems(4))
    lstLivil = lstJobSublet.SelectedItem.SubItems(5)
    lstLine_No = lstJobSublet.SelectedItem.Text
End Sub

Private Sub lstMaterials_DblClick()
    If lstJobSublet.ListItems.Count = 0 And lstMaterials.ListItems.Count = 0 And lstparts.ListItems.Count = 0 Then
        MsgBox "This transaction has no Job to edit.", vbInformation, "INFORMATION"
        Exit Sub
    End If

    If CheckIfImported((RTrim(LTrim(txtReceive))), (RTrim(LTrim(txtInvNo)))) = True Then
        MsgBox "You can't edit this transaction" & vbCrLf & "It's Already Posted in accounting", vbInformation, "INFORMATION"
        Exit Sub
    End If

    If Picture1.Visible = True Then
        If Null2String(rsINFO!Status) = "C" Then
            MsgBox "Purchase order already Cancelled. Cannot Edit this Job.", vbInformation, "INFORMATION"
            Exit Sub
        ElseIf Null2String(rsINFO!Status) = "P" Then
            MsgBox "Purchase order already posted. Cannot Edit this Job.", vbInformation, "INFORMATION"
            Exit Sub
        Else
            frmCSMS_SubletJob.LINE_NO.Caption = lstLine_No
            Call editJobs
            frmCSMS_SubletJob.lblAddorEdit.Caption = "REDIT"
            Call EnabledFrame(False)
            frmCSMS_SubletJob.txtSubletAmount.Enabled = False
            frmCSMS_SubletJob.Show
        End If
    End If
End Sub

Private Sub lstMaterials_Click()
    If lstMaterials.ListItems.Count = 0 Then Exit Sub
    labDET.Caption = lstMaterials.SelectedItem.SubItems(4)
    lstLivil = lstMaterials.SelectedItem.SubItems(5)
    lstLine_No = lstMaterials.SelectedItem.Text
End Sub

Private Sub lstparts_Click()
    If lstparts.ListItems.Count = 0 Then Exit Sub
    labDET.Caption = lstparts.SelectedItem.SubItems(4)
    lstLivil = lstparts.SelectedItem.SubItems(5)
    lstLine_No = lstparts.SelectedItem.Text
End Sub

Private Sub lstParts_DblClick()
    If lstJobSublet.ListItems.Count = 0 And lstMaterials.ListItems.Count = 0 And lstparts.ListItems.Count = 0 Then
        MsgBox "This transaction has no Job to edit.", vbInformation, "INFORMATION"
        Exit Sub
    End If

    If CheckIfImported((RTrim(LTrim(txtReceive))), (RTrim(LTrim(txtInvNo)))) = True Then
        MsgBox "You can't edit this transaction" & vbCrLf & "It's Already Posted in accounting", vbInformation, "INFORMATION"
        Exit Sub
    Else
        'proceed to edit
    End If

    If Picture1.Visible = True Then
        If Null2String(rsINFO!Status) = "C" Then
            MsgBox "Purchase order already Cancelled. Cannot Edit this Job.", vbInformation, "INFORMATION"
            Exit Sub

        ElseIf Null2String(rsINFO!Status) = "P" Then
            MsgBox "Purchase order already posted. Cannot Edit this Job.", vbInformation, "INFORMATION"
            Exit Sub
        Else
            frmCSMS_SubletJob.LINE_NO.Caption = lstLine_No
            Call editJobs
            frmCSMS_SubletJob.lblAddorEdit.Caption = "REDIT"
            Call EnabledFrame(False)
            frmCSMS_SubletJob.txtSubletAmount.Enabled = False
            frmCSMS_SubletJob.Show
        End If
    End If
End Sub

Private Sub lvwTran_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsINFO.MoveFirst
    rsINFO.Find ("ID=" & Item.ListSubItems(1).Text)
    Call StoreMemVars
End Sub

Private Sub optRCNo_Click()
    Dim rsoptRCNUM                                     As New ADODB.Recordset
    lvwTran.Enabled = False
    lvwTran.Sorted = False: lvwTran.ListItems.Clear
    
    Set rsoptRCNUM = gconDMIS.Execute("Select top 30 RC_NO,ID from CSMS_PO_RC_HD order by RC_NO desc")
    If Not (rsoptRCNUM.EOF And rsoptRCNUM.BOF) Then
        Listview_Loadval Me.lvwTran.ListItems, rsoptRCNUM
        lvwTran.Refresh
    End If
    lvwTran.Enabled = True
    Set rsoptRCNUM = Nothing
    
    On Error Resume Next
    txtSearch.Text = ""
    txtSearch.SetFocus
End Sub

Private Sub optContractor_Click()
    Dim rsoptContractor                                As New ADODB.Recordset
    lvwTran.Enabled = False
    lvwTran.Sorted = False: lvwTran.ListItems.Clear
    Set rsoptContractor = gconDMIS.Execute("Select TOP 30 contractor_name,ID from CSMS_PO_RC_HD order by contractor_name asc")
    If Not (rsoptContractor.EOF And rsoptContractor.BOF) Then
        Listview_Loadval Me.lvwTran.ListItems, rsoptContractor
        lvwTran.Refresh
    End If
    lvwTran.Enabled = True
    Set rsoptContractor = Nothing
    
    On Error Resume Next
    txtSearch.Text = ""
    txtSearch.SetFocus
End Sub

Private Sub Timer1_Timer()
    If lblPOSTED.Caption <> "" Then
        If lblPOSTED.Visible = True Then
            lblPOSTED.Visible = False
        Else
            lblPOSTED.Visible = True
        End If
    End If

    If Label17.Caption <> "" Then
        If Label17.Visible = True Then
            Label17.Visible = False
        Else
            Label17.Visible = True
        End If
    End If
End Sub


Private Sub txtRcNumber_LostFocus()
    txtRcNumber = Format(txtRcNumber, "000000")
End Sub

Private Sub txtSearch_Change()
    Dim rsSearch                                        As New ADODB.Recordset
    Dim rcNUMBER                                        As String
    Dim rcNUMBER2                                       As String
    Dim rcNUMBER3                                       As String
    Dim vContractor                                     As String
    Dim k                                               As Integer

    If optRCNo.Value = True Then
        rcNUMBER = UCase(txtSearch.Text)
        If txtSearch = "" Then
            lvwTran.Enabled = False
            lvwTran.Sorted = False: lvwTran.ListItems.Clear
            Set rsSearch = gconDMIS.Execute("Select top 30 RC_NO ,ID from CSMS_PO_RC_HD order by RC_NO desc ")
            If Not (rsSearch.EOF And rsSearch.BOF) Then
                Listview_Loadval Me.lvwTran.ListItems, rsSearch
                lvwTran.Refresh
            End If
            lvwTran.Enabled = True
        ElseIf rcNUMBER <> "" Then
            If IsNumeric(rcNUMBER) = True Then
                'rcNUMBER = Format(Right(rcNUMBER, 6), "000000")
                rcNUMBER = rcNUMBER
            Else
                For k = 1 To Len(rcNUMBER)
                    rcNUMBER2 = Mid(rcNUMBER, k, 1)
                    If IsNumeric(rcNUMBER2) = True Then rcNUMBER3 = rcNUMBER3 + rcNUMBER2
                Next
                'rcNUMBER = Format(rcNUMBER3, "000000")
                rcNUMBER = rcNUMBER3
            End If
        End If
        If IsNumeric(rcNUMBER) = True Then
            lvwTran.Enabled = False
            lvwTran.Sorted = False: lvwTran.ListItems.Clear
            Set rsSearch = gconDMIS.Execute("Select top 30 RC_NO ,ID from CSMS_PO_RC_HD where RC_NO like '" & rcNUMBER & "%'")
            If Not (rsSearch.EOF And rsSearch.BOF) Then
                Listview_Loadval Me.lvwTran.ListItems, rsSearch
                lvwTran.Refresh
            End If
            lvwTran.Enabled = True
        End If
    Else
        vContractor = UCase(txtSearch.Text)
        If txtSearch.Text = "" Then
            lvwTran.Enabled = False
            lvwTran.Sorted = False: lvwTran.ListItems.Clear
            Set rsSearch = gconDMIS.Execute("Select top 30 Contractor_name ,ID from CSMS_PO_RC_HD order by contractor_name asc")
            If Not (rsSearch.EOF And rsSearch.BOF) Then
                Listview_Loadval Me.lvwTran.ListItems, rsSearch
                lvwTran.Refresh
            End If
            lvwTran.Enabled = True
        Else
            lvwTran.Enabled = False
            lvwTran.Sorted = False: lvwTran.ListItems.Clear
            Set rsSearch = gconDMIS.Execute("Select top 30 Contractor_name ,ID from CSMS_PO_RC_HD where Contractor_name like '" & vContractor & "%'")
            If Not (rsSearch.EOF And rsSearch.BOF) Then
                Listview_Loadval Me.lvwTran.ListItems, rsSearch
                lvwTran.Refresh
            End If
            lvwTran.Enabled = True
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SUBLET RECEIVING)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "SUBLET RECEIVING", "")

        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsINFO!Status) = "C" Then
                    MsgBox "Purchase order already Cancelled. Cannot Add Further Job.", vbInformation, "INFORMATION"
                    Exit Sub

                ElseIf Null2String(rsINFO!Status) = "P" Then
                    MsgBox "Purchase order already posted. Cannot Add Further Job.", vbInformation, "INFORMATION"
                    Exit Sub
                Else
                    'UPDATED BY: JUN
                    'DATE UPDATED: 09052008
                    'DESCRIPTION: USER NOT ALLOWED TO ADD JOB IF REPAIRORDER IS ALREADY INVOICED
                    Dim rsInvoice                      As New ADODB.Recordset
                    Set rsInvoice = gconDMIS.Execute("Select DTE_COMP from CSMS_repor where rep_or ='" & Null2String(rsINFO!RO_NO) & "' and dte_comp is not null")
                    If Not rsInvoice.EOF And Not rsInvoice.BOF Then
                        MsgBox "You Cannot Add further Job... Already Invoice", vbOKOnly, "INFORMATION"
                        Exit Sub
                    Else
                        frmCSMS_SubletJob.lblAddorEdit.Caption = "RADD"
                        passINFO
                        frmCSMS_SubletJob.Show
                        Exit Sub
                    End If
                End If
            End If

        Case vbKeyF4
            MsgBox "Double Click the Item you want to Edit", vbInformation, "Information"
            '            If lstJobSublet.ListItems.Count = 0 And lstMaterials.ListItems.Count = 0 And lstparts.ListItems.Count = 0 Then
            '                MsgBox "This transaction has no Job to edit.", vbInformation, "INFORMATION"
            '                Exit Sub
            '            End If
            '            If Picture1.Visible = True Then
            '                If Null2String(rsInfo!Status) = "C" Then
            '                    MsgBox "Purchase order already Cancelled. Cannot Edit this Job.", vbInformation, "INFORMATION"
            '                    Exit Sub
            '
            '                ElseIf Null2String(rsInfo!Status) = "P" Then
            '                    MsgBox "Purchase order already posted. Cannot Edit this Job.", vbInformation, "INFORMATION"
            '                    Exit Sub
            '                Else
            '                    frmCSMS_SubletJob.LINE_NO.Caption = lstLine_No
            '                    Call editJobs
            '                    frmCSMS_SubletJob.lblAddorEdit.Caption = "REDIT"
            '                    frmCSMS_SubletJob.txtSubletAmount.Enabled = False
            '
            '                End If
            '            End If

        Case vbKeyF5
            If Picture1.Visible = True Then
                If Null2String(rsINFO!Status) = "C" Then
                    MsgBox "Purchase order already Cancelled. Cannot Delete this Job.", vbInformation, "INFORMATION"
                    Exit Sub

                ElseIf Null2String(rsINFO!Status) = "P" Then
                    MsgBox "Purchase order already posted. Cannot Delete this Job.", vbInformation, "INFORMATION"
                    Exit Sub
                Else
                    Call deleteJobs
                    Exit Sub
                End If
            End If
        Case vbKeyF8
            If cmdPost.Enabled = True And Picture1.Visible = True Then
                cmdPost.Value = True
            End If

        Case vbKeyF12
            cmdUnPost.Value = True
            '            If cmdUnPost.Enabled = True And Picture1.Visible = True Then
            '            cmdUnPost.Value = True
            '            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Function CheckIfImported(xVendorCode As String, xInvoiceNo As String) As Boolean
    Dim rsImported                                     As New ADODB.Recordset
    Set rsImported = gconDMIS.Execute("Select * from AMIS_journal_hd where VENDORCODE = '" & xVendorCode & "' AND INVOICETYPE = 'SUBLET' AND INVOICENO = '" & xInvoiceNo & "' and STATUS <> 'C'")
    If Not rsImported.EOF And Not rsImported.BOF Then
        CheckIfImported = True
    Else
        CheckIfImported = False
    End If
    Set rsImported = Nothing
End Function

Sub getAPJ_NO(xVendorCode As String, xInvoiceNo As String)
    Dim rsImported                                     As New ADODB.Recordset
    Set rsImported = gconDMIS.Execute("Select * from AMIS_journal_hd where VENDORCODE = '" & xVendorCode & "' AND INVOICETYPE = 'SUBLET' AND INVOICENO = '" & xInvoiceNo & "' and STATUS <> 'C'")
    If Not rsImported.EOF And Not rsImported.BOF Then
        txtAPJNO.Text = Null2String(rsImported!VOUCHERNO)
        Label17.Caption = "**Already Imported In Accounting**"
    Else
        txtAPJNO.Text = ""
        Label17.Caption = "**Not Yet Imported**"
    End If
    Set rsImported = Nothing
End Sub

Function SAVERECEVING() As Boolean
    Dim rsRcNumDup                                     As New ADODB.Recordset
    Dim vcboPoNumber                                   As String
    Dim vtxtReceive                                    As String
    Dim vcboContractor                                 As String
    Dim vtxtContractorAdd                              As String
    Dim vtxtDelReceipt                                 As String
    Dim vtxtInvNo                                      As String
    Dim vDTPicker1                                     As String
    Dim vtxtTotalAmount                                As Double
    Dim vtxtVatAmount                                  As Double
    Dim vtxtNetAmount                                  As Double
    Dim vtxtRcNumber                                   As String
    Dim vDTPicker2                                     As String
    Dim Vusercode                                      As String
    Dim vSavedate                                      As Date
    Dim VStatus                                        As String
    Dim ictr                                           As Integer
    Dim Ictr2                                          As Integer
    Dim Ictr3                                          As Integer
    Dim Ictr4                                          As Integer
    Dim invexist                                       As Integer
    Dim drexist                                        As Integer
    
    On Error GoTo ErrorCode

    If txtInvNo.Text = "" And txtDelReceipt = "" Then
        'MsgBox "Delivery or Invoice number must have a value Receipt must have a value", vbInformation, "INFORMATION"
        str_MSG = "Delivery or Invoice number must have a value Receipt must have a value"
        GoTo ErrorCode
    End If

    If txtRcNumber = "" Then
        'MsgBox "RR number must have a value", vbInformation, "INformation"
        str_MSG = "RR number must have a value"
        GoTo ErrorCode
    End If

    If cboPoNumber.Text = "" Then
         'MsgBox "PO number must have a value", vbInformation, "INformation"
         str_MSG = "PO number must have a value"
        GoTo ErrorCode
  End If
    
    vcboPoNumber = N2Str2Null(cboPoNumber.Text)
    vtxtReceive = N2Str2Null(txtReceive.Text)
    vcboContractor = N2Str2Null(cboContractor)
    vtxtContractorAdd = N2Str2Null(txtContractorAdd.Text)
    vtxtDelReceipt = N2Str2Null(txtDelReceipt.Text)
    vtxtInvNo = N2Str2Null(txtInvNo.Text)
    vDTPicker1 = N2Date2Null(DTPicker1.Value)
    vtxtTotalAmount = NumericVal(txtTotalAmount.Text)
    vtxtVatAmount = NumericVal(txtVatAmount.Text)
    vtxtNetAmount = NumericVal(txtNetAmount.Text)
    'vtxtDelReceipt = N2Str2Null(txtInvNo.Text)
    vtxtInvNo = N2Str2Null(txtInvNo.Text)
    vtxtRcNumber = Format(N2Str2Null(txtRcNumber.Text), "000000")
    vDTPicker2 = N2Date2Null(DTPicker2.Value)
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    vSavedate = DateValue(Now) & " " & TimeValue(Now)
    VStatus = "'R'"
    
    invexist = NumericVal((gconDMIS.Execute("Select count(*) from CSMS_PO_RC_HD where contractor_code = " & vtxtReceive & " and invoice_no = '" & RTrim(LTrim(txtInvNo.Text)) & "'").Fields(0).Value))
    drexist = NumericVal((gconDMIS.Execute("Select count(*) from CSMS_PO_RC_HD where contractor_code = " & vtxtReceive & " and delivery_or = '" & RTrim(LTrim(txtInvNo.Text)) & "'").Fields(0).Value))
    ictr = NumericVal((gconDMIS.Execute("Select count(*) from CSMS_PO_RC_HD where Po_no = " & vcboPoNumber & " and STATUS = 'R'").Fields(0).Value))
    Ictr2 = NumericVal((gconDMIS.Execute("Select count(*) from CSMS_PO_RC_HD where Po_no = " & vcboPoNumber & " and STATUS = 'P'").Fields(0).Value))
    Ictr3 = NumericVal((gconDMIS.Execute("Select count(*) from csms_po_hd where po_no = '" & cboPoNumber.Text & "'").Fields(0).Value))
    Ictr4 = NumericVal((gconDMIS.Execute("Select count(*) from csms_po_hd where po_no = '" & cboPoNumber.Text & "' and status in ('R','C')").Fields(0).Value))
    If AddorEdit = "ADD" Then
'updated by: IEBV 05032011_1030AM
'description: PO number cannot be recieve if already recieve
'---------------------------------------------------------------------------------------------------
        If ictr > 0 Then
            'MsgSpeechBox "PO number already recieve but not yet posted!"
            str_MSG = "PO number already recieve but not yet posted!"
            On Error Resume Next
            cboPoNumber.SetFocus
            GoTo ErrorCode
        End If
        
        If Ictr2 > 0 Then
            'MsgSpeechBox "PO number already recieve!"
            str_MSG = "PO number already recieve!"
            On Error Resume Next
            cboPoNumber.SetFocus
            GoTo ErrorCode
        End If
        
        If Ictr3 = 0 Then
            'MsgSpeechBox "PO number did not exist!"
            str_MSG = "PO number did not exist!"
            On Error Resume Next
            cboPoNumber.SetFocus
            GoTo ErrorCode
        End If
        
        If Ictr4 > 0 Then
            'MsgSpeechBox "PO number not yet posted!"
            str_MSG = "PO number not yet posted!"
            On Error Resume Next
            cboPoNumber.SetFocus
            GoTo ErrorCode
        End If
        
        If invexist > 0 Then
            'MsgSpeechBox "Ref. invoice number already exist!"
            str_MSG = "Ref. invoice number already exist!"
            On Error Resume Next
            txtInvNo.SetFocus
            GoTo ErrorCode
        End If
        
        If drexist > 0 Then
            str_MSG = "Ref. DR number already exist!"
            On Error Resume Next
            txtInvNo.SetFocus
            GoTo ErrorCode
        End If
'---------------------------------------------------------------------------------------------------
        
        rsRcNumDup.Open "Select RC_NO from CSMS_PO_RC_HD where RC_NO = '" & Trim(txtRcNumber) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsRcNumDup.EOF And Not rsRcNumDup.BOF Then
            'MsgSpeechBox "RC Number already exist!"
            str_MSG = "RR Number already exist!"
            On Error Resume Next
            txtRcNumber.SetFocus
            GoTo ErrorCode
        End If
        
        
    Else
        rsRcNumDup.Open "select RC_NO from CSMS_PO_RC_HD where RC_NO = '" & Trim(txtRcNumber) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If LTrim(RTrim(txtRcNumber)) <> Null2String(rsRcNumDup!RC_NO) Then
            If Not rsRcNumDup.EOF And Not rsRcNumDup.BOF Then
                'MsgSpeechBox "RC Number already exist!"
                str_MSG = "RR Number already exist!"
                On Error Resume Next
                txtRcNumber.SetFocus
                GoTo ErrorCode
            End If
         End If
        If RTrim(LTrim(txtInvNo.Text)) <> Null2String(rsINFO!invoice_no) Then
            If invexist > 0 Then
                'MsgSpeechBox "Ref. invoice number already exist!"
                str_MSG = "Ref. invoice number already exist!"
                On Error Resume Next
                txtInvNo.SetFocus
                GoTo ErrorCode
            End If
        End If
        
        If RTrim(LTrim(txtDelReceipt.Text)) <> Null2String(rsINFO!delivery_or) Then
            If drexist > 0 Then
                'MsgSpeechBox "Ref. invoice number already exist!"
                str_MSG = "Ref. DR number already exist!"
                On Error Resume Next
                txtInvNo.SetFocus
                GoTo ErrorCode
            End If
        End If
        
        If LTrim(RTrim(cboPoNumber)) <> Null2String(rsINFO!PO_NO) Then
            If ictr > 0 Then
                'MsgSpeechBox "PO number already recieve but not yet posted!"
                str_MSG = "PO number already recieve but not yet posted!"
                On Error Resume Next
                cboPoNumber.SetFocus
                GoTo ErrorCode
            End If
            
            If Ictr2 > 0 Then
                'MsgSpeechBox "PO number already recieve!"
                str_MSG = "PO number already recieve!"
                On Error Resume Next
                cboPoNumber.SetFocus
                GoTo ErrorCode
            End If
            If Ictr3 = 0 Then
                'MsgSpeechBox "PO number did not exist!"
                str_MSG = "PO number did not exist!"
                On Error Resume Next
                cboPoNumber.SetFocus
                GoTo ErrorCode
            End If
            
            If Ictr4 > 0 Then
                'MsgSpeechBox "PO number not yet posted!"
                str_MSG = "PO number not yet posted!"
                On Error Resume Next
                cboPoNumber.SetFocus
                GoTo ErrorCode
            End If
        End If
        
    End If

    Dim rsReceiveDetail                                As New ADODB.Recordset
    Dim dtPONO                                         As String
    Dim dtRep_or                                       As String
    Dim dtROTYPE                                       As String
    Dim dtJOBTYPE                                      As String
    Dim dtLIVIL                                        As String
    Dim dtLINE_NO                                      As String
    Dim dtDETAMT                                       As Double
    Dim dtDETCDE                                       As String
    Dim dtDETDSC                                       As String
    Dim dtTECHNICIAN                                   As String
    Dim dtWCODE                                        As String
    Dim dtTAXRATE                                      As Double
    Dim dtTAXVAL                                       As Double
    Dim dtSTATUS                                       As String
    Dim dtDETAIL                                       As String
    Dim dtDET_AMT                                      As Double
    Dim dtTECHCODE                                     As String
    Dim dtCONTRACTAMOUNT                               As Double
    Dim dtCOMPAMOUNT                                   As Double
    Dim dtDONE                                         As String
    Dim vSUBLET_TYPE                                   As String
    
    rsReceiveDetail.Open "Select * from CSMS_PO_DT where PO_NO = '" & cboPoNumber.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If AddorEdit = "ADD" Then
        If Not rsReceiveDetail.EOF And Not rsReceiveDetail.BOF Then
            Do While Not rsReceiveDetail.EOF
                'If CheckIfAlreadyAdded(Null2String(rsReceiveDetail!LINE_NO), Null2String(rsReceiveDetail!po_no)) = False Then
                dtPONO = N2Str2Null(rsReceiveDetail!PO_NO)
                dtRep_or = N2Str2Null(rsReceiveDetail!REP_OR)
                dtROTYPE = N2Str2Null(rsReceiveDetail!ROTYPE)
                dtJOBTYPE = N2Str2Null(rsReceiveDetail!JOBTYPE)
                dtLIVIL = N2Str2Null(rsReceiveDetail!LIVIL)
                dtLINE_NO = N2Str2Null(rsReceiveDetail!LINE_NO)
                dtDETAMT = NumericVal(rsReceiveDetail!DETAMT)
                dtDETCDE = N2Str2Null(rsReceiveDetail!DETCDE)
                dtDETDSC = N2Str2Null(rsReceiveDetail!DETDSC)
                dtTECHNICIAN = N2Str2Null(rsReceiveDetail!Technician)
                dtWCODE = N2Str2Null(rsReceiveDetail!wCode)
                dtTAXRATE = NumericVal(rsReceiveDetail!taxrate)
                dtTAXVAL = NumericVal(rsReceiveDetail!TAXVAL)
                dtSTATUS = "'R'"
                dtDETAIL = N2Str2Null(rsReceiveDetail!Detail)
                dtDET_AMT = Null2String(rsReceiveDetail!DET_AMT)
                dtTECHCODE = N2Str2Null(rsReceiveDetail!TechCode)
                dtCONTRACTAMOUNT = NumericVal(rsReceiveDetail!CONTRACTAMOUNT)
                dtCOMPAMOUNT = NumericVal(rsReceiveDetail!COMPAMOUNT)
                dtDONE = N2Str2Null(rsReceiveDetail!DONE)
                vSUBLET_TYPE = N2Str2Null(rsReceiveDetail!SUBLET_TYPE)
                
                gconDMIS.Execute "Insert into CSMS_PO_RC_DT " & _
                    "(RC_NO, PO_NO, REP_OR, ROTYPE, JOBTYPE, LIVIL, LINE_NO, DETAMT, DETCDE, DETDSC, TECHNICIAN, WCODE, TAXRATE, TAXVAL, STATUS, DETAIL, DET_AMT, USERCODE, SAVEDATE, TECHCODE, CONTRACTAMOUNT, COMPAMOUNT, DONE,SUBLET_TYPE)" & _
                    " values(" & vtxtRcNumber & _
                    ", " & dtPONO & _
                    ", " & dtRep_or & _
                    ", " & dtROTYPE & _
                    ", " & dtJOBTYPE & _
                    ", " & dtLIVIL & _
                    ", " & dtLINE_NO & _
                    ", " & dtDETAMT & _
                    ", " & dtDETCDE & _
                    ", " & dtDETDSC & _
                    ", " & dtTECHNICIAN & _
                    ", " & dtWCODE & _
                    ", " & dtTAXRATE & _
                    ", " & dtTAXVAL & _
                    ", " & dtSTATUS & _
                    ", " & dtDETAIL & _
                    ", " & dtDET_AMT & _
                    ", " & Vusercode & _
                    ", '" & vSavedate & _
                    "', " & dtTECHCODE & _
                    ", " & dtCONTRACTAMOUNT & _
                    ", " & dtCOMPAMOUNT & _
                    ", " & dtDONE & ", " & vSUBLET_TYPE & ")"
                rsReceiveDetail.MoveNext
            Loop
        End If
    Else
        'Update the header  only
        Dim SQLTXT As String
        
        SQLTXT = "Update CSMS_PO_RC_DT set RC_NO = '" & Trim(txtRcNumber) & "' where PO_NO = '" & cboPoNumber & "'"
        gconDMIS.Execute (SQLTXT)
        'sqltxt = sqltxt & " AND ID = '" & labID & "'"
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_PO_RC_HD" & _
            "(RC_NO, RC_DATE, Po_No, Ro_No, Po_Date, contractor_code, Contractor_Name, Contractor_Address, SUBLET_TOTAL_AMT, SUBLET_TOTAL_VAT, SUBLET_TOTAL_NET_AMT, INVOICE_NO, DELIVERY_OR, USER_CODE, LAST_UPDATE, STATUS)" & _
            "values(" & vtxtRcNumber & _
            ", " & vDTPicker2 & _
            ", " & vcboPoNumber & _
            ", " & dtRep_or & _
            ", " & vDTPicker1 & _
            ", " & vtxtReceive & _
            ", " & vcboContractor & _
            ", " & Replace(vtxtContractorAdd, ",", " ") & _
            ", " & vtxtTotalAmount & _
            ", " & vtxtVatAmount & _
            ", " & vtxtNetAmount & _
            ", " & vtxtInvNo & _
            ", " & vtxtDelReceipt & _
            ", " & Vusercode & _
            ", '" & vSavedate & _
            "', " & VStatus & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("A", "SUBLET RECEIVING", SQL_STATEMENT, FindTransactionID(N2Str2Null(vtxtRcNumber), "RC_NO", "CSMS_PO_RC_HD"), "", "RC NO: " & txtRcNumber, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

    Else
        SQL_STATEMENT = "update CSMS_Po_RC_HD set " & _
            "RC_NO = " & vtxtRcNumber & "," & _
            "RC_DATE = " & vDTPicker2 & "," & _
            "Po_No = " & vcboPoNumber & "," & _
            "Po_Date = " & vDTPicker1 & "," & _
            "contractor_code = " & vtxtReceive & "," & _
            "contractor_name = " & vcboContractor & "," & _
            "Contractor_address = " & vcboContractor & "," & _
            "SUBLET_TOTAL_AMT = " & vtxtTotalAmount & "," & _
            "SUBLET_TOTAL_VAT = " & vtxtVatAmount & "," & _
            "SUBLET_TOTAL_NET_AMT = " & vtxtNetAmount & "," & _
            "INVOICE_NO = " & vtxtInvNo & "," & _
            "DELIVERY_OR = " & vtxtDelReceipt & "," & _
            "LAST_UPDATE = '" & vSavedate & "'" & _
            "where ID =" & labID
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "SUBLET RECEIVING", SQL_STATEMENT, labID, "", "RC NO: " & txtRcNumber, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

    End If
    
    SAVERECEVING = True
    Call rsRefresh

    If AddorEdit = "EDIT" Then
        rsINFO.Find ("ID =" & labID)
    Else
        rsINFO.Find ("RC_NO =" & txtRcNumber)
    End If
    
    txtRcNumber.Locked = True
    Call FilllstPO_HD
    cmdCancel.Value = True
    Exit Function
ErrorCode:
     SAVERECEVING = False
     Exit Function
End Function



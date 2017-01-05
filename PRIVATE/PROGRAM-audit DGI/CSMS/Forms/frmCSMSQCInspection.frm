VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSQCInspection 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quality Control  Inspection"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   FillColor       =   &H00808080&
   Icon            =   "frmCSMSQCInspection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCSMSQCInspection.frx":08CA
   ScaleHeight     =   7515
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox thePromp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   4560
      ScaleHeight     =   1095
      ScaleWidth      =   3735
      TabIndex        =   40
      Top             =   2790
      Width           =   3765
      Begin VB.PictureBox Picture13 
         BackColor       =   &H00808000&
         Height          =   225
         Left            =   -30
         ScaleHeight     =   165
         ScaleWidth      =   4935
         TabIndex        =   42
         Top             =   1050
         Width           =   4995
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H00808000&
         Height          =   135
         Left            =   0
         ScaleHeight     =   75
         ScaleWidth      =   4935
         TabIndex        =   41
         Top             =   -60
         Width           =   4995
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "No Finish Job Available"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   390
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00808000&
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   12915
      TabIndex        =   38
      Top             =   7410
      Width           =   12975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1980
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   8
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSMSQCInspection.frx":0BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSMSQCInspection.frx":15C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame TheFrame 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   2040
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame TheClock 
         BackColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   1680
         TabIndex        =   31
         Top             =   780
         Visible         =   0   'False
         Width           =   6435
         Begin MSComctlLib.ListView ListClock 
            Height          =   2295
            Left            =   120
            TabIndex        =   34
            Top             =   420
            Width           =   6135
            _ExtentX        =   10821
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "In"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Out"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Reason"
               Object.Width           =   3175
            EndProperty
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H00808000&
            Height          =   135
            Left            =   -180
            ScaleHeight     =   75
            ScaleWidth      =   6495
            TabIndex        =   33
            Top             =   3120
            Width           =   6555
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00808000&
            Height          =   135
            Left            =   -180
            ScaleHeight     =   75
            ScaleWidth      =   6495
            TabIndex        =   32
            Top             =   0
            Width           =   6555
         End
         Begin VB.Label lblClose 
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   5640
            MouseIcon       =   "frmCSMSQCInspection.frx":1FB8
            MousePointer    =   99  'Custom
            TabIndex        =   36
            ToolTipText     =   "Close"
            Top             =   2820
            Width           =   615
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Clock In/Out History"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   35
            Top             =   180
            Width           =   2115
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808000&
         Height          =   345
         Index           =   0
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   9225
         TabIndex        =   4
         Top             =   0
         Width           =   9285
         Begin VB.PictureBox Picture2 
            BackColor       =   &H000080FF&
            Height          =   135
            Left            =   -600
            ScaleHeight     =   75
            ScaleWidth      =   9795
            TabIndex        =   7
            Top             =   240
            Width           =   9855
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   28
         Top             =   3510
         Width           =   2595
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Approve"
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
            Left            =   240
            MouseIcon       =   "frmCSMSQCInspection.frx":22C2
            MousePointer    =   99  'Custom
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reject"
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
            Left            =   1500
            MouseIcon       =   "frmCSMSQCInspection.frx":25CC
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00000000&
         Height          =   195
         Left            =   5580
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   26
         Top             =   360
         Width           =   195
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00008000&
         Height          =   195
         Left            =   7680
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   21
         Top             =   360
         Width           =   195
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Height          =   195
         Left            =   6540
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   19
         Top             =   360
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808000&
         Height          =   165
         Index           =   1
         Left            =   0
         ScaleHeight     =   105
         ScaleWidth      =   9225
         TabIndex        =   5
         Top             =   4380
         Width           =   9285
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   7800
         MouseIcon       =   "frmCSMSQCInspection.frx":28D6
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Close"
         Top             =   3600
         Width           =   1005
      End
      Begin MSComctlLib.ListView listAlJob 
         Height          =   2835
         Left            =   240
         TabIndex        =   2
         Top             =   630
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   5001
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   0
         BackColor       =   -2147483644
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSQCInspection.frx":2BE0
         NumItems        =   41
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "JobType"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description "
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "FlatRate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Technician"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DetCode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Dealertype"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Transtype"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "EstimateNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Appno"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "RoType"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "livil"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Line_no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "det_hrs"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "hrswrk"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "detunt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "detvol"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "detprc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "detcost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "detamt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "wcode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "taxrate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "disval"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "pocode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "detail"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "detamt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "disval"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "discount2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Ref RIV ADB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "user code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "savedate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "savetime"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "transtatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "status1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "techcode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "disrate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "rep0r2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.PictureBox ThePic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   2850
         ScaleHeight     =   1665
         ScaleWidth      =   4305
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdCanUpdate 
            Caption         =   "Cancel"
            Height          =   435
            Left            =   3390
            MouseIcon       =   "frmCSMSQCInspection.frx":2D42
            MousePointer    =   99  'Custom
            TabIndex        =   44
            Top             =   1140
            Width           =   765
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            Height          =   435
            Left            =   2610
            MouseIcon       =   "frmCSMSQCInspection.frx":304C
            MousePointer    =   99  'Custom
            TabIndex        =   37
            Top             =   1140
            Width           =   765
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00808000&
            Height          =   135
            Left            =   0
            ScaleHeight     =   75
            ScaleWidth      =   4335
            TabIndex        =   18
            Top             =   1620
            Width           =   4395
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00808000&
            Height          =   135
            Left            =   0
            ScaleHeight     =   75
            ScaleWidth      =   4275
            TabIndex        =   17
            Top             =   -60
            Width           =   4335
         End
         Begin VB.Label lblApprove 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   390
            TabIndex        =   25
            Top             =   120
            Width           =   3615
         End
         Begin VB.Label lblOption 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   960
            TabIndex        =   24
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "This Job?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1980
            TabIndex        =   23
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lbljob 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   16
            Top             =   1260
            Width           =   2715
         End
         Begin VB.Label lbljob 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   15
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lbljob 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   14
            Top             =   660
            Width           =   3075
         End
         Begin VB.Label lbljob 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   13
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Technician:"
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
            TabIndex        =   12
            Top             =   1260
            Width           =   1035
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Flat Rate:"
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
            Left            =   240
            TabIndex        =   11
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Job Type:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "For QC"
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
         Left            =   5820
         TabIndex        =   27
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved"
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
         Left            =   7980
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Rejected"
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
         Left            =   6780
         TabIndex        =   20
         Top             =   360
         Width           =   795
      End
      Begin VB.Label LblRO 
         BackStyle       =   0  'Transparent
         Caption         =   "TheRo"
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
         Height          =   285
         Left            =   300
         TabIndex        =   6
         Top             =   360
         Width           =   2025
      End
   End
   Begin MSComctlLib.ListView ListAlFinish 
      Height          =   7275
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   12832
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
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCSMSQCInspection.frx":3356
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Repair Order"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Plate No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Model"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Promise date"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Hour"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Datefinish"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Status"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label TheInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Information:Double Click The  Items For QC Inspection"
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
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   7140
      Width           =   4275
   End
End
Attribute VB_Name = "frmCSMSQCInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theRo                                              As String
Dim thetotal_lineNo                                    As Integer
Dim theCheckflag, flag                                 As Boolean
Attribute flag.VB_VarUserMemId = 1073938434
Dim TheDetCode, thejobtype, thedetdcs, thetechnician, theflatrate, TheApprove, thedetdsc, thetechcode As String
Attribute TheDetCode.VB_VarUserMemId = 1073938436
Attribute thejobtype.VB_VarUserMemId = 1073938436
Attribute thedetdcs.VB_VarUserMemId = 1073938436
Attribute thetechnician.VB_VarUserMemId = 1073938436
Attribute theflatrate.VB_VarUserMemId = 1073938436
Attribute TheApprove.VB_VarUserMemId = 1073938436
Attribute thedetdsc.VB_VarUserMemId = 1073938436
Attribute thetechcode.VB_VarUserMemId = 1073938436
Dim theID, theline_no                                  As String
Attribute theID.VB_VarUserMemId = 1073938444
Attribute theline_no.VB_VarUserMemId = 1073938444

Sub DisplayAllFinishRo()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer



    SQL = "SELECT Ro_no,Customer,Plate_no,Model,xhrswork,DateFinish,PromiseDate,Status FROM CSMS_vw_RepairOrder WHERE Status='Finish Job'"



    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListAlFinish.ListItems.Clear
    cnt = 0
    With RS
        If .EOF And .BOF Then
            thePromp.Visible = True
            ListAlFinish.Enabled = False
        End If
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = ListAlFinish.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!RO_NO)
            ITEM.SubItems(2) = Null2String(!Customer)
            ITEM.SubItems(3) = Null2String(!PLATE_NO)
            ITEM.SubItems(4) = Null2String(!MODEL)
            ITEM.SubItems(5) = Null2String(!PromiseDate)
            ITEM.SubItems(6) = Null2String(!xHrsWork)
            ITEM.SubItems(7) = Null2String(!datefinish)
            ITEM.SubItems(8) = Null2String(!Status)
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub DisplayJob()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt, I                                         As Integer
    Dim X, y                                           As Boolean

    SQL = "SELECT * FROM CSMS_Ro_Det Where Rep_or ='" & theRo & "' and livil='1'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listAlJob.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = listAlJob.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!JOBTYPE)
            ITEM.SubItems(2) = Null2String(!DETDSC)
            ITEM.SubItems(3) = Null2String(!FLATRATE)
            ITEM.SubItems(4) = Null2String(!Technician)
            ITEM.SubItems(5) = Null2String(!DETCDE)

            ITEM.SubItems(12) = Null2String(!LINE_NO)
            ITEM.SubItems(35) = Null2String(!Approve)
            ITEM.SubItems(37) = Null2String(!TechCode)
            ITEM.SubItems(40) = Null2String(!ID)

            If ITEM.SubItems(35) = "Approved" Then
                For I = 1 To listAlJob.ColumnHeaders.Count - 1
                    listAlJob.ListItems(cnt).ListSubItems(I).ForeColor = &H8000&
                Next
            End If
            If ITEM.SubItems(35) = "Rejected" Then
                For I = 1 To listAlJob.ColumnHeaders.Count - 1
                    listAlJob.ListItems(cnt).ListSubItems(I).ForeColor = vbRed
                Next
            End If
            .MoveNext
            If ITEM.SubItems(35) = "" Then
                ITEM.SubItems(35) = "ForQC"
            End If

        Loop
    End With
    Set RS = Nothing
End Sub

Sub InsertData(ID, thestatus)
    Dim SQL                                            As String
    Dim X                                              As String

    X = "Frm QC:"

    Call AddlineNo

    SQL = " INSERT INTO CSMS_RO_det "
    SQL = SQL & " SELECT "
    SQL = SQL & " DEALER_TYPE , TRANSTYPE, REP_OR, "
    SQL = SQL & " ESTIMATENO, APPTNO, ROTYPE,"
    SQL = SQL & " JOBTYPE, LIVIL, " & thetotal_lineNo & ","
    SQL = SQL & " DETCDE,DETDSC, TECHNICIAN,"
    SQL = SQL & " FLATRATE, DET_HRS, HRSWRK,"
    SQL = SQL & " DETUNT, DETVOL, '0',"
    SQL = SQL & " DETCOST, DETAMT, CODE,"
    SQL = SQL & " WCODE, TAXRATE, DISCRATE,"
    SQL = SQL & " TAXVAL, DISVAL, POCODE,"
    SQL = SQL & " REP_OR2, DETAIL, '0',"
    SQL = SQL & " DIS_VAL, DISCOUNT_2, status,"
    SQL = SQL & " REF_RIV_ADB, " & N2Str2Null(LOGCODE) & ", " & N2Str2Null(Date) & " ,"
    SQL = SQL & " SAVETIME, TranStatus, 'Rejected',"
    SQL = SQL & " STATUS1 ,'N', TechCode"
    SQL = SQL & " From CSMS_Ro_Det WHERE ID=" & ID

    gconDMIS.Execute (SQL)

End Sub

Sub CheckIfApprove()
    If theCheckflag = True Then
        If StrComp(TheApprove, "Approved") = 0 Then
            lblApprove.Visible = True
            lblApprove.Caption = "This Job Is Aready Approved"
            lblApprove.ForeColor = &H8000&
            lblOption.Visible = False
            Label7.Visible = False
            cmdUpdate.Enabled = False
        Else
            lblApprove.Visible = False
            lblOption.Visible = True
            Label7.Visible = True
            cmdUpdate.Enabled = True
        End If
        If StrComp(TheApprove, "Rejected") = 0 Then
            lblApprove.Visible = True
            lblApprove.Caption = "This Job Is Aready Rejected"
            lblApprove.ForeColor = &H8000&
            lblOption.Visible = False
            Label7.Visible = False
            cmdUpdate.Enabled = False
        End If
    End If
    If theCheckflag = False Then
        If StrComp(TheApprove, "Rejected") = 0 Then
            lblApprove.Visible = True
            lblApprove.Caption = "This Job Is Already Rejected"
            lblApprove.ForeColor = vbRed
            lblOption.Visible = False
            Label7.Visible = False
            cmdUpdate.Visible = False
        Else
            lblApprove.Visible = False
            lblOption.Visible = True
            Label7.Visible = True
            cmdUpdate.Enabled = True
        End If
        If StrComp(TheApprove, "Approved") = 0 Then
            lblApprove.Visible = True
            lblApprove.Caption = "This Job Is Aready Approved"
            lblApprove.ForeColor = &H8000&
            lblOption.Visible = False
            Label7.Visible = False
            cmdUpdate.Enabled = False
        End If

        If StrComp(TheApprove, "") = 0 Then
            lblApprove.Visible = False
            lblOption.Visible = True
            Label7.Visible = True
            cmdUpdate.Enabled = True
        End If
    End If
End Sub

Sub DisplayTheClock()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    SQL = "SELECT clockin,clockout,reasonforclockout From CSMS_jobClock WHERE ro_no='" & theRo & "' and line_no='" & theline_no & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListClock.ListItems.Clear

    With RS
        Do While Not .EOF
            Set ITEM = ListClock.ListItems.Add(, , !clockin)
            ITEM.SubItems(1) = Null2String(!clockout)
            ITEM.SubItems(2) = Null2String(!reasonforclockout)
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub checkApprove()
    If StrComp(TheApprove, "Approved") = 0 Or StrComp(TheApprove, "Rejected") = 0 Then
        Check1.Enabled = False
        Check2.Enabled = False
    Else
        Check1.Enabled = True
        Check2.Enabled = True
    End If
End Sub

Sub AddlineNo()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim total                                          As Integer
    SQL = "SELECT line_no From CSMS_Ro_det Where Rep_or='" & theRo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    With RS
        .MoveFirst
        While .EOF = False

            total = total + 1
            .MoveNext

        Wend

        thetotal_lineNo = total + 1

    End With
    Set RS = Nothing
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        If listAlJob.SelectedItem.SubItems(35) = "Approved" Then
            MsgBox "Job Is Already Approved", vbInformation, "Quality Control"
            If Check2.Value = 1 Then Check2.Value = 0
            Check1.Value = 0
            Exit Sub
        End If

        theCheckflag = True
        If Check1.Value = 1 Then
            cmdClose.Visible = False
            If theflatrate <> "" Then
                Call checkApprove
                cmdUpdate.Visible = True
                thepic.Visible = True
                lblOption.Caption = "Approve"
                lblOption.ForeColor = &H8000&
            Else
                MsgBox "Please Select A job!", vbExclamation, "Information"
                Check1.Value = 0
                Check2.Value = 0
            End If
        Else
            Check2.Enabled = True
            cmdUpdate.Visible = False
            thepic.Visible = False
        End If
    End If
End Sub

Private Sub Check2_Click()

    If Check2.Value = 1 Then
        If listAlJob.SelectedItem.SubItems(35) = "Approved" Then
            MsgBox "This Job Is Already Approved", vbInformation, "Quality Control"
            If Check1.Value = 1 Then Check1.Value = 0
            Check2.Value = 0
            Exit Sub
        End If
        theCheckflag = False

        If Check2.Value = 1 Then
            cmdClose.Visible = False
            If theflatrate <> "" Then
                Check1.Enabled = False
                cmdUpdate.Visible = True
                thepic.Visible = True

                checkApprove
                lblOption.Caption = "Reject"
                lblOption.ForeColor = vbRed
            Else
                MsgBox "Please Select A job!", vbExclamation, "Information"
                Check1.Value = 0
                Check2.Value = 0
            End If

        Else
            Check1.Enabled = True
            cmdUpdate.Visible = False
            thepic.Visible = False
        End If
    End If
End Sub

Private Sub cmdCanUpdate_Click()
    thepic.Visible = False
    cmdClose.Visible = True
    Frame1.Enabled = True
    Check1.Value = 0
    Check2.Value = 0
End Sub

Private Sub cmdClose_Click()
    TheFrame.Visible = False
    TheClock.Visible = False
    ListAlFinish.Enabled = True
    Check1.Value = 0
    Check2.Value = 0
    Check1.Enabled = True
    Check2.Enabled = True
    DisplayAllFinishRo
End Sub

Private Sub cmdUpdate_Click()
    If Function_Access(LOGID, "Acess_EDIT", "QUALITY INSPECTION") = False Then Exit Sub

    Dim thestatus                                      As String
    Dim theAnswer                                      As String
    Dim Description                                    As String

    If Check1.Value = 1 Then
        thestatus = "Approved"
    Else
        thestatus = "Rejected"
    End If

    theAnswer = MsgBox("Are You Sure You Want to Update This Ro", vbYesNo + vbQuestion, "InFormation")
    If theAnswer = vbYes Then
        If TheDetCode = "" Then
            MsgBox "Pls Select a Job To Be Update!", vbExclamation, "Warning!"
        Else
            If thestatus = "Approved" Then
                gconDMIS.Execute "update CSMS_Ro_det set approve ='" & thestatus & "' WHERE Id  ='" & theID & "' and DETCDE = '" & TheDetCode & "'"
                MsgBox "All Information Has been Update.Job Has Been Approved", vbInformation, "Confirm"

                Check1.Value = 0
                Check2.Value = 0
            Else
                InsertData listAlJob.SelectedItem.ListSubItems(40).Text, thestatus
                gconDMIS.Execute "Update CSMS_repairOrder set status ='Back Job',jStatus = 'W' Where Ro_no='" & theRo & " '"
                Description = "frm QC:" + thedetdsc
                gconDMIS.Execute "update CSMS_Ro_det set detdsc ='" & Description & "' WHERE Id  ='" & theID & "' and DETCDE = '" & TheDetCode & "'"
                gconDMIS.Execute "update CSMS_vw_Technician set AssignedRO = '" & theRo & "',JStatus = 'S' where Technician = '" & thetechcode & "'"
                MsgBox "All Information Has been Update.job Has Been Rejected", vbInformation, "Confirm"

                Check1.Value = 0
                Check2.Value = 0
            End If
        End If

        cmdClose.Visible = True
    Else
        Check1.Value = 0
        Check2.Value = 0
    End If
    LogAudit "E", "QUALITY INFORMATION:FOR RO ", "RO/STATUS:" & theRo & "/" & thestatus
    Call DisplayJob
    thepic.Visible = False
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    thePromp.Visible = False
    DisplayAllFinishRo
    TheFrame.Visible = False
    cmdUpdate.Visible = False

    thepic.Visible = False
    TheClock.Visible = False
End Sub

Private Sub Label10_Click()
    TheClock.Visible = False
    Check1.Enabled = True
    Check2.Enabled = True
End Sub

Private Sub lblclose_Click()
    cmdClose.Visible = True
    Frame1.Enabled = True

    Call listAlJob_Click
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblclose.ForeColor = vbRed
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblclose.ForeColor = vbBlack
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblclose.ForeColor = vbBlack
    TheClock.Visible = False
    Check1.Enabled = True
    Check2.Enabled = True
End Sub

Private Sub ListAlFinish_DblClick()
    If Function_Access(LOGID, "Acess_Add", "QUALITY INSPECTION") = False Then Exit Sub
    On Error Resume Next

    theRo = ListAlFinish.SelectedItem.SubItems(1)
    TheFrame.Visible = True
    lblRO.Caption = theRo

    Call DisplayJob
    ListAlFinish.Enabled = False
    TheInfo.Visible = False

    Call listAlJob_Click
End Sub

Private Sub listAlJob_Click()
    On Error Resume Next

    thejobtype = listAlJob.SelectedItem.SubItems(1)
    theflatrate = listAlJob.SelectedItem.SubItems(3)
    thetechnician = listAlJob.SelectedItem.SubItems(4)
    TheDetCode = listAlJob.SelectedItem.SubItems(5)
    theID = listAlJob.SelectedItem.SubItems(40)
    TheApprove = listAlJob.SelectedItem.SubItems(35)
    theline_no = listAlJob.SelectedItem.SubItems(12)
    thedetdsc = listAlJob.SelectedItem.SubItems(2)
    thetechcode = listAlJob.SelectedItem.SubItems(37)

    lbljob(0).Caption = thejobtype
    lbljob(1).Caption = listAlJob.SelectedItem.SubItems(2)
    lbljob(2).Caption = listAlJob.SelectedItem.SubItems(3)
    lbljob(3).Caption = listAlJob.SelectedItem.SubItems(4)


    If listAlJob.SelectedItem.SubItems(35) = "Approved" Or listAlJob.SelectedItem.SubItems(35) = "Rejected" Then
        Check2.Enabled = False: Check1.Enabled = False
    Else
        Check2.Enabled = True: Check1.Enabled = True
    End If


End Sub

Private Sub listAlJob_DblClick()
    If listAlJob.SelectedItem.SubItems(35) = "Approved" Then
        TheClock.Visible = True
        Frame1.Enabled = False

        cmdClose.Visible = False
        Call DisplayTheClock
    Else

    End If
End Sub

Private Sub txtsearch_Change()
    DisplayAllFinishRo
End Sub


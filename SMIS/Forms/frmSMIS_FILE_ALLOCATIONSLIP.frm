VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_FILE_ALLOCATIONSLIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALLOCATION SLIP"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11550
   Icon            =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   11550
   Begin VB.Frame FRAPO 
      Height          =   615
      Left            =   8160
      TabIndex        =   46
      Top             =   150
      Width           =   3255
      Begin VB.ComboBox cboPO 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   900
         TabIndex        =   47
         Text            =   "Combo1"
         Top             =   165
         Width           =   2265
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Po No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.PictureBox PICSEARCH 
      BorderStyle     =   0  'None
      Height          =   6345
      Left            =   0
      ScaleHeight     =   6345
      ScaleWidth      =   2715
      TabIndex        =   41
      Top             =   0
      Width           =   2715
      Begin VB.OptionButton Option2 
         Caption         =   "Account Name"
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
         Left            =   180
         TabIndex        =   45
         ToolTipText     =   "Search By Customer Account Name"
         Top             =   360
         Width           =   1665
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PO No"
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
         Left            =   180
         TabIndex        =   44
         Top             =   90
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   30
         TabIndex        =   43
         Top             =   690
         Width           =   2625
      End
      Begin MSComctlLib.ListView lstPO 
         Height          =   5205
         Left            =   0
         TabIndex        =   42
         Top             =   1110
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   9181
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PO NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CLIENT NAME"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin Crystal.CrystalReport rptAllocation 
      Left            =   11850
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   3225
      Left            =   2730
      TabIndex        =   5
      Top             =   2100
      Width           =   8805
      Begin VB.TextBox txtMonths 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   51
         Text            =   "0"
         Top             =   180
         Width           =   585
      End
      Begin VB.TextBox txtBAnk 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4530
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   750
         Width           =   4035
      End
      Begin VB.ComboBox cboConsultant 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1950
         TabIndex        =   34
         Text            =   "Combo1"
         Top             =   1290
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker txt_RecieveDate 
         Height          =   315
         Left            =   1950
         TabIndex        =   14
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
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
         CheckBox        =   -1  'True
         Format          =   54067201
         CurrentDate     =   39408
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   1950
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   2355
      End
      Begin VB.ComboBox cbomode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1950
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   240
         Width           =   2385
      End
      Begin VB.CheckBox creditAPP 
         Caption         =   "Credit Approved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5580
         TabIndex        =   10
         Top             =   180
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker txt_ProposedDate 
         Height          =   345
         Left            =   1950
         TabIndex        =   15
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
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
         Format          =   54067201
         CurrentDate     =   39408
      End
      Begin MSComCtl2.DTPicker txt_ApprovedDate 
         Height          =   345
         Left            =   1950
         TabIndex        =   17
         Top             =   2070
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
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
         Format          =   54067201
         CurrentDate     =   39408
      End
      Begin VB.Label Label16 
         Caption         =   "/Months"
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
         Left            =   7980
         TabIndex        =   52
         Top             =   240
         Width           =   795
      End
      Begin VB.Label labid 
         Caption         =   "0"
         Height          =   405
         Left            =   6270
         TabIndex        =   40
         Top             =   1530
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label15 
         Caption         =   "Bank Approver:"
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
         Left            =   4530
         TabIndex        =   39
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label Label11 
         Caption         =   "-If With Reservation,Please Provide copy of the official Receipt,VSA signed by Cashier"
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
         Left            =   180
         TabIndex        =   29
         Top             =   2670
         Width           =   7725
      End
      Begin VB.Label Label12 
         Caption         =   "-If with company/Bank PO,please provide copy of purchased order issued by the company/bank"
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
         Left            =   180
         TabIndex        =   28
         Top             =   2880
         Width           =   7425
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Top             =   2400
         Width           =   555
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Approve:"
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
         Left            =   150
         TabIndex        =   16
         Top             =   2070
         Width           =   1665
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Proposed date Released:"
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
         Left            =   180
         TabIndex        =   11
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Received by:"
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
         Left            =   180
         TabIndex        =   9
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Received:"
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
         Left            =   180
         TabIndex        =   8
         Top             =   930
         Width           =   1665
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
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
         Left            =   180
         TabIndex        =   7
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Mode Of Payment:"
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
         Left            =   210
         TabIndex        =   6
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vehicle Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   2730
      TabIndex        =   0
      Top             =   0
      Width           =   8805
      Begin VB.TextBox txtCuscde 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1980
         TabIndex        =   50
         Text            =   "Text2"
         Top             =   1110
         Width           =   1365
      End
      Begin VB.CommandButton cmdAddFromPO 
         Height          =   345
         Left            =   8160
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Add Customer From List"
         Top             =   1110
         Width           =   345
      End
      Begin VB.ComboBox cboSA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1980
         TabIndex        =   37
         Text            =   "Combo1"
         Top             =   1500
         Width           =   3405
      End
      Begin VB.ComboBox cbocolor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1980
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   720
         Width           =   3405
      End
      Begin VB.ComboBox cboModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1980
         TabIndex        =   35
         Text            =   "Combo1"
         Top             =   330
         Width           =   3405
      End
      Begin VB.TextBox txtcust 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3390
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   1110
         Width           =   4755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Consultant:"
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
         Left            =   270
         TabIndex        =   4
         Top             =   1500
         Width           =   1605
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Client Name:"
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
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Color:"
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
         Left            =   1290
         TabIndex        =   2
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Model/Unit:"
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
         Left            =   810
         TabIndex        =   1
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   10020
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   30
      Top             =   5400
      Width           =   1800
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
         Left            =   750
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":0A94
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":0BE6
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Left            =   60
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":0F24
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":1076
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   5790
      ScaleHeight     =   915
      ScaleWidth      =   8805
      TabIndex        =   18
      Top             =   5370
      Width           =   8805
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
         Left            =   5010
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":13C6
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":1518
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   4320
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":187E
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":19D0
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Print this Record"
         Top             =   60
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
         Left            =   3630
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":1D36
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":1E88
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
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
         Left            =   2940
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":21B3
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":2305
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   2250
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":2661
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":27B3
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   1560
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":2AC6
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":2C18
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   870
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":2F12
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":3064
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   180
         MouseIcon       =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":33BC
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_FILE_ALLOCATIONSLIP.frx":350E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmSMIS_FILE_ALLOCATIONSLIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AddorEdit                                                         As String
Private WithEvents frm                                                As frmSMIS_Mis_SearchMaster
Attribute frm.VB_VarHelpID = -1
Dim RS                                                                As New ADODB.Recordset

Sub LoadCbo()
    Combo_Loadval cboModel, gconDMIS.Execute("SELECT DESCRIPT FROM ALL_MODEL order by descript")
    Combo_Loadval cboColor, gconDMIS.Execute("SELECT color_desc  FROM all_color order by color_desc")
    Combo_Loadval cboConsultant, gconDMIS.Execute("SELECT NAME FROM SMIS_VW_SREP  order by NAME ")
    Combo_Loadval cboSA, gconDMIS.Execute("SELECT NAME FROM SMIS_VW_SREP  order by NAME ")
    cbomode.Clear
    cbomode.AddItem "Letter of credit"
    cbomode.AddItem "Open Account"
    cbomode.AddItem "Promissory Note"
    cbomode.AddItem "Finance Co"
    cbomode.AddItem "Cash"
    cbomode.AddItem "Company Purchased"
    cbomode.AddItem "Bank Purchased"
    cbomode.AddItem "Reservation"
End Sub

Sub SaveAllocation()
    Dim SQL                                                           As String

    Dim isApproved                                                    As String
    Dim ismonth                                                       As String


    On Error GoTo ErrorCode:

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT SMIS_ALLOCATION(so_no,model,color,clientname,sc,modeofpayment,amount,creditapprove,months,proposedaterel,Receivedby,datereceived,dateApprove,bankApprover) VALUES('" & cboPO.Text & _
                        "','" & cboModel & "','" & cboColor & "','" & txtcust.Text & "','" & cboSA.Text _
                      & "','" & cbomode & "','" & NumericVal(txtAmount) & "','" & isApproved & "'," & NumericVal(txtMonths) & _
                        ",'" & txt_ProposedDate.Value & "','" & cboConsultant & "','" & txt_RecieveDate & "','" & txt_ApprovedDate & "','" & txtBAnk.Text & "')"

        '**********NEW LOG AUDIT**************
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "A", "ALLOCATION SLIP", SQL_STATEMENT, FindTransactionID(N2Str2Null(cboPO), "SO_NO", "SMIS_ALLOCATION"), "", "PO NO: " & cboPO, "", ""
        '**********NEW LOG AUDIT**************

        LogAudit "A", "SALES ALLOCATION", "FOR CUSTOMER " & txtcust & " PO" & cboPO
    Else

        SQL_STATEMENT = ("update smis_po set cuscde='" & Repleys(txtCuscde) & "' where po_no=" & N2Str2Null(RS!SO_NO))

        '**********NEW LOG AUDIT**************
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "EE", "ALLOCATION SLIP", SQL_STATEMENT, N2Str2Null(labid), "", "PO NO: " & cboPO, "", ""
        '**********NEW LOG AUDIT**************

        '*********RESET THE SQL_STATEMENT VARIABLE************
        SQL_STATEMENT = ""
        '*********RESET THE SQL_STATEMENT VARIABLE************

        SQL_STATEMENT = "UPDATE SMIS_ALLOCATION set so_no='" & cboPO.Text & "', color='" & cboColor.Text & _
                        "',clientname='" & txtcust.Text & _
                        "',Months='" & NumericVal(txtMonths) & _
                        "',sc='" & cboSA.Text & "',modeofpayment='" & cbomode.Text & _
                        "',amount=" & NumericVal(txtAmount) & ",datereceived='" & txt_RecieveDate.Value & _
                        "',dateapprove='" & txt_ApprovedDate.Value & "',bankapprover='" & txtBAnk.Text & _
                        "' where ID= '" & labid & "'"
        gconDMIS.Execute (SQL_STATEMENT)

        '**********NEW LOG AUDIT**************
        NEW_LogAudit "E", "ALLOCATION SLIP", SQL_STATEMENT, N2Str2Null(labid), "", "PO NO: " & cboPO, "", ""
        '**********NEW LOG AUDIT**************

        LogAudit "E", "SALES ALLOCATION", "FOR CUSTOMER " & txtcust & " PO" & cboPO
    End If
    '*********RESET THE SQL_STATEMENT VARIABLE************
    SQL_STATEMENT = ""
    '*********RESET THE SQL_STATEMENT VARIABLE************

    SQL_STATEMENT = ("update smis_po set cuscde='" & Repleys(txtCuscde) & "' where po_no=" & N2Str2Null(cboPO))

    gconDMIS.Execute (SQL_STATEMENT)
    '**********NEW LOG AUDIT**************
    NEW_LogAudit "EE", "ALLOCATION SLIP", SQL_STATEMENT, N2Str2Null(labid), "", "PO NO: " & cboPO, "", ""
    '**********NEW LOG AUDIT**************

    MsgBox "Allocation slip has been save..", vbInformation, "INFORMATION"
    rsRefresh
    If AddorEdit = "EDIT" Then
        RS.Find ("ID=" & labid)
    End If
    cmdCancel.Value = True

    Exit Sub

ErrorCode:
    MsgBox "Error..pls Check the data entry", vbExclamation, "INFORMATION"
End Sub

Sub InitMemVars()
    cboModel.Text = ""
    cboColor.Text = ""
    txtCuscde = ""
    cboSA.Text = ""
    txtcust.Text = ""
    txtAmount.Text = ""
    txt_ProposedDate.Value = LOGDATE
    txt_ApprovedDate.Value = LOGDATE
    txtBAnk.Visible = False
    Label15.Visible = False
    cboConsultant.Text = ""
    txtAmount.Text = "0.00"
    cbomode.Text = ""
    cboPO.Text = ""
    txtBAnk.Text = ""
    FRAPO.Enabled = False
End Sub

Sub lockeME(ByVal X As Boolean)
    Frame1.Enabled = Not X
    Frame2.Enabled = Not X
    FRAPO.Enabled = Not X

End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM SMIS_ALLOCATION ORDER BY ID DESC", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    Dim rsCust                                                        As ADODB.Recordset
    If Not RS.EOF And Not RS.BOF Then

        labid = Null2String(RS!ID)
        cboPO.Text = Null2String(RS!SO_NO)
        cboModel.Text = Null2String(RS!Model)
        cboColor.Text = Null2String(RS!Color)
        txtcust.Text = Null2String(RS!Clientname)
        cboSA.Text = Null2String(RS!sc)
        cbomode.Text = Null2String(RS!modeofpayment)
        txtAmount.Text = FormatNumber(NumericVal(RS!AMOUNT))
        txt_RecieveDate.Value = Null2String(RS!datereceived)
        cboConsultant.Text = Null2String(RS!receivedby)
        txt_ProposedDate.Value = Null2String(RS!proposedaterel)
        txt_ApprovedDate.Value = Null2String(RS!dateApprove)
        txtMonths = NumericVal(RS!months)
        Set rsCust = gconDMIS.Execute("SELECT CUSCDE FROM SMIS_PO WHERE PO_NO='" & Null2String(RS!SO_NO) & "'")
        If Not rsCust.EOF Or Not rsCust.BOF Then
            txtCuscde = Null2String(rsCust!CUSCDE)
        End If
        If cbomode = "Bank Purchased" Then
            txtBAnk.Visible = True
            Label15.Visible = True
            txtBAnk.Text = Null2String(RS!bankapprover)
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True

    End If
End Sub

Sub loadPO()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset

    SQL = "SELECT PO_NO FROM SMIS_PO WHERE PO_NO NOT IN (SELECT SO_NO FROM SMIS_ALLOCATION) order by po_no asc"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cboPO.Clear

    Do While Not RS.EOF
        cboPO.AddItem Null2String(RS!po_no)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub FillSearchGrid()
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim XXX                                                           As String

    If Option1.Value = True Then

        Set TEMPRS = gconDMIS.Execute("SELECT SO_NO , CLIENTNAME, ID FROM SMIS_ALLOCATION WHERE SO_NO Like " & N2Str2Null(ReplaceQuote(txtSEARCH & "%")))
    Else

        Set TEMPRS = gconDMIS.Execute("SELECT CLIENTNAME ,SO_NO, ID FROM SMIS_ALLOCATION WHERE  CLIENTNAME Like " & N2Str2Null(ReplaceQuote(txtSEARCH & "%")))
    End If


    Listview_Loadval lstPO.ListItems, TEMPRS

    Set TEMPRS = Nothing
    LV_AutoSizeColumn lstPO
End Sub

Private Sub cbomode_Click()
    If cbomode = "Bank Purchased" Then
        txtBAnk.Visible = True
        Label15.Visible = True
    Else
        txtBAnk.Visible = False
        Label15.Visible = False
    End If
End Sub

Private Sub cboPO_Change()
    If AddorEdit = "" Then Exit Sub
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset

    cboModel = ""
    cboColor = ""
    txtcust = ""
    SQL = "SELECT Modeldescript,color , datereceived from SMIS_PO Where PO_NO='" & cboPO.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        cboModel.Text = Null2String(RS!ModelDescript)
        cboColor.Text = Null2String(RS!Color)
        txt_RecieveDate = Null2String(RS!datereceived)
        lockeME False
    End If
    Set RS = Nothing
End Sub

Private Sub cboPO_Click()
    cboPO_Change
End Sub

Private Sub cboPO_Validate(Cancel As Boolean)
    cboPO.ListIndex = SelectCombo(cboPO, cboPO)
End Sub

Private Sub cmdAdd_Click()
    InitMemVars
    picSaves.Visible = True: picAdds.Visible = False
    AddorEdit = "ADD"
    loadPO
    lockeME (True): PICSEARCH.Enabled = False
    FRAPO.Enabled = True

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub cmdAddFromPO_Click()

    frm.SearchForCustomers
    frm.Show 1
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picSaves.Visible = False
    picAdds.Visible = True
    lockeME True
    StoreMemVars
    PICSEARCH.Enabled = True
End Sub

Private Sub cmdDelete_Click()
    Dim ans                                                           As String

    ans = MsgBox("Are you sure do you want to delete this record?", vbQuestion + vbYesNo)

    If ans = vbYes Then
        SQL_STATEMENT = "DELETE  FROM SMIS_ALLOCATION WHERE So_no='" & cboPO.Text & "'"
        '*************NEW LOG AUDIT***************
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "X", "ALLOCATION SLIP", SQL_STATEMENT, N2Str2Null(labid), "", "PO NO:" & cboPO, "", ""
        '*************NEW LOG AUDIT***************

        '*************RESET THE VARIABLE SQL_STATEMENT***************
        SQL_STATEMENT = ""
        '*************RESET THE VARIABLE SQL_STATEMENT***************

        SQL_STATEMENT = ("UPDATE SMIS_PO SET CUSCDE=NULL WHERE PO_NO=" & N2Str2Null(RS!SO_NO))
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "EE", "ALLOCATION SLIP", SQL_STATEMENT, N2Str2Null(labid), "", "PO NO:" & cboPO, "", ""

        MsgBox "Allocation Information has been deleted..", vbInformation, "Information"
        rsRefresh
        StoreMemVars
        FillSearchGrid
    End If
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    picAdds.Visible = False
    picSaves.Visible = True
    lockeME (False)

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    lockeME True
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

Private Sub cmdPrevious_Click()
    On Error Resume Next
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String

    Dim theReservation                                                As String
    Dim company                                                       As String
    Dim bank                                                          As String

    SQL = "SELECT * FROM SMIS_ALLOCATION WHERE SO_No='" & cboPO.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If cbomode.Text = "Reservation" Then
        theReservation = txtAmount.Text
    End If

    If cbomode.Text = "Company Purchased" Then
        company = txtAmount.Text
    End If

    If cbomode.Text = "Bank Purchased" Then
        bank = txtAmount.Text
    End If

    If Not RS.EOF And Not RS.BOF Then

        rptAllocation.Formulas(1) = "isreservation='" & theReservation & "'"
        rptAllocation.Formulas(2) = "iscompany='" & company & "'"
        rptAllocation.Formulas(3) = "isbankpo='" & bank & "'"
        rptAllocation.Formulas(4) = "CompanyName='" & COMPANY_NAME & "'"
        rptAllocation.Formulas(5) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptAllocation, SMIS_REPORT_PATH & "Allocation.rpt", "{SMIS_ALLOCATION.SO_no} = '" & cboPO.Text & "' ", DMIS_REPORT_Connection, 1
    Else
        MsgBox "No Records..", vbInformation, "No Record"
    End If
End Sub

Private Sub cmdSave_Click()
    SaveAllocation
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (ALLOCATION SLIP)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "ALLOCATION SLIP")
            'End If
    End Select

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set frm = New frmSMIS_Mis_SearchMaster
    picAdds.Visible = True: picSaves.Visible = False
    InitMemVars
    loadPO
    LoadCbo
    lockeME (True)
    rsRefresh
    StoreMemVars
    FillSearchGrid
End Sub

Private Sub frm_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    txtcust = Null2String(oCusRs!AcctName)
    txtCuscde = Null2String(oCusRs!CUSCDE)
    Unload frm
End Sub

Private Sub lstPO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPO
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

Private Sub lstPO_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    RS.MoveFirst
    RS.Find ("id=" & Item.ListSubItems(2).Text)
    StoreMemVars
End Sub

Private Sub Option1_Click()
    lstPO.ColumnHeaders(1).Text = "PO NO": lstPO.ColumnHeaders(2).Text = "CLIENT NAME"
End Sub

Private Sub Option2_Click()
    lstPO.ColumnHeaders(1).Text = "CLIENT NAME": lstPO.ColumnHeaders(2).Text = "PO NO"
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid
End Sub

Private Sub txtAmount_GotFocus()
    If IsNumeric(txtAmount) = True Then
        txtAmount = FormatNumber(txtAmount)
    Else
        txtAmount = "0.00"
    End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtAmount_Validate(Cancel As Boolean)
    If IsNumeric(txtAmount) = True Then
        txtAmount = FormatNumber(txtAmount)
    Else
        txtAmount = "0.00"
    End If
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "wizBox.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmHRMSEditCards 
   BackColor       =   &H00D7C6B5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viewing Employee Attendance"
   ClientHeight    =   10575
   ClientLeft      =   1530
   ClientTop       =   1170
   ClientWidth     =   13410
   ForeColor       =   &H8000000F&
   Icon            =   "EditCards.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EditCards.frx":030A
   ScaleHeight     =   10575
   ScaleWidth      =   13410
   Begin FlexCell.Grid Grid1 
      Height          =   4785
      Left            =   4200
      TabIndex        =   596
      Top             =   480
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   8440
      BackColorBkg    =   -2147483645
      BackColorSel    =   13811126
      Cols            =   12
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   17
      ScrollBars      =   0
   End
   Begin VB.CommandButton cmdAdd_2ndCutoff 
      Caption         =   "Fill 2nd Cut-Off Time  "
      Enabled         =   0   'False
      Height          =   795
      Left            =   1740
      MouseIcon       =   "EditCards.frx":3C034C
      MousePointer    =   99  'Custom
      Picture         =   "EditCards.frx":3C049E
      Style           =   1  'Graphical
      TabIndex        =   595
      ToolTipText     =   "Fill 2nd Cut-Off Time  "
      Top             =   8760
      Width           =   1125
   End
   Begin VB.TextBox txtgrid 
      Height          =   375
      Left            =   13860
      TabIndex        =   590
      Top             =   900
      Width           =   855
   End
   Begin VB.TextBox txtin_am 
      Height          =   375
      Left            =   13860
      TabIndex        =   589
      Top             =   1500
      Width           =   855
   End
   Begin VB.TextBox txtout_pm 
      Height          =   375
      Left            =   13860
      TabIndex        =   588
      Top             =   2100
      Width           =   975
   End
   Begin VB.PictureBox picNAME 
      Appearance      =   0  'Flat
      BackColor       =   &H00F77C48&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   60
      ScaleHeight     =   2865
      ScaleWidth      =   4155
      TabIndex        =   573
      Top             =   930
      Visible         =   0   'False
      Width           =   4185
      Begin MSComctlLib.ListView lsvSEARCH 
         Height          =   2325
         Left            =   60
         TabIndex        =   576
         Top             =   450
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
            Text            =   "Employee Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox Text3 
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
         Left            =   60
         TabIndex        =   575
         Top             =   60
         Width           =   4035
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00F77C48&
      Caption         =   "EMPLOYEE NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   570
      Top             =   720
      Width           =   1815
   End
   Begin VB.PictureBox picEMPNO 
      Appearance      =   0  'Flat
      BackColor       =   &H00F77C48&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   60
      ScaleHeight     =   1215
      ScaleWidth      =   4155
      TabIndex        =   568
      Top             =   930
      Width           =   4185
      Begin VB.CommandButton Command3 
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2550
         TabIndex        =   572
         Top             =   120
         Width           =   1305
      End
      Begin VB.TextBox TxtEmpName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   60
         TabIndex        =   571
         Top             =   600
         Width           =   4035
      End
      Begin VB.ComboBox cboEmpNumber 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   60
         TabIndex        =   569
         Text            =   "cboEmpNumber"
         Top             =   90
         Width           =   2355
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00F77C48&
      Caption         =   "EMPLOYEE NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   567
      Top             =   720
      Value           =   -1  'True
      Width           =   1485
   End
   Begin VB.PictureBox picNOTES2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   14520
      ScaleHeight     =   2385
      ScaleWidth      =   75
      TabIndex        =   367
      Top             =   6510
      Visible         =   0   'False
      Width           =   105
      Begin VB.ComboBox cboType2 
         Appearance      =   0  'Flat
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   523
         Top             =   1140
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.TextBox txtOT2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   2820
         MaxLength       =   3
         TabIndex        =   487
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtNOTES2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   370
         Top             =   330
         Width           =   3225
      End
      Begin VB.CommandButton cmdSaveNote2 
         Caption         =   "Save"
         Height          =   345
         Left            =   1320
         TabIndex        =   369
         Top             =   1950
         Width           =   1005
      End
      Begin VB.CommandButton cmdCancelNote2 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   2340
         TabIndex        =   368
         Top             =   1950
         Width           =   1005
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   0
         Left            =   1590
         TabIndex        =   555
         Top             =   2820
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   554
         Top             =   3180
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   2
         Left            =   1590
         TabIndex        =   553
         Top             =   3570
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   552
         Top             =   3990
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   551
         Top             =   4410
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   5
         Left            =   1620
         TabIndex        =   550
         Top             =   4830
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   6
         Left            =   1590
         TabIndex        =   549
         Top             =   5250
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   7
         Left            =   1620
         TabIndex        =   548
         Top             =   5700
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   8
         Left            =   1590
         TabIndex        =   547
         Top             =   6180
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   9
         Left            =   1590
         TabIndex        =   546
         Top             =   6630
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   10
         Left            =   1620
         TabIndex        =   545
         Top             =   6990
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   11
         Left            =   1590
         TabIndex        =   544
         Top             =   7380
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   12
         Left            =   1560
         TabIndex        =   543
         Top             =   7830
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   13
         Left            =   1560
         TabIndex        =   542
         Top             =   8220
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   14
         Left            =   1560
         TabIndex        =   541
         Top             =   8580
         Width           =   555
      End
      Begin VB.Label lblOTCode2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   15
         Left            =   1560
         TabIndex        =   540
         Top             =   8940
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   15
         Left            =   840
         TabIndex        =   521
         Top             =   8940
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   14
         Left            =   840
         TabIndex        =   520
         Top             =   8580
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   13
         Left            =   840
         TabIndex        =   519
         Top             =   8220
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   12
         Left            =   840
         TabIndex        =   518
         Top             =   7830
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   11
         Left            =   870
         TabIndex        =   517
         Top             =   7380
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   10
         Left            =   900
         TabIndex        =   516
         Top             =   6990
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   9
         Left            =   870
         TabIndex        =   515
         Top             =   6630
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   8
         Left            =   870
         TabIndex        =   514
         Top             =   6180
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   7
         Left            =   900
         TabIndex        =   513
         Top             =   5700
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   6
         Left            =   870
         TabIndex        =   512
         Top             =   5250
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   5
         Left            =   900
         TabIndex        =   511
         Top             =   4830
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   510
         Top             =   4410
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   509
         Top             =   3990
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   2
         Left            =   870
         TabIndex        =   508
         Top             =   3570
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   507
         Top             =   3180
         Width           =   555
      End
      Begin VB.Label lblOTno2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   0
         Left            =   870
         TabIndex        =   506
         Top             =   2820
         Width           =   555
      End
      Begin VB.Label lblNoOfOT2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1830
         TabIndex        =   489
         Top             =   1620
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblCAP2 
         BackColor       =   &H000000FF&
         Height          =   195
         Left            =   2700
         TabIndex        =   485
         Top             =   90
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblNoteTitle2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   90
         TabIndex        =   484
         Top             =   90
         Width           =   465
      End
      Begin VB.Label lblIndex2 
         BackColor       =   &H000000FF&
         Height          =   225
         Left            =   1200
         TabIndex        =   483
         Top             =   60
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   482
         Top             =   2820
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   481
         Top             =   3180
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   480
         Top             =   3570
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   3
         Left            =   60
         TabIndex        =   479
         Top             =   3990
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   4
         Left            =   60
         TabIndex        =   478
         Top             =   4410
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   477
         Top             =   4830
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   6
         Left            =   90
         TabIndex        =   476
         Top             =   5250
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   475
         Top             =   5700
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   8
         Left            =   90
         TabIndex        =   474
         Top             =   6180
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   9
         Left            =   90
         TabIndex        =   473
         Top             =   6630
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   472
         Top             =   6990
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   11
         Left            =   90
         TabIndex        =   471
         Top             =   7380
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   12
         Left            =   60
         TabIndex        =   470
         Top             =   7830
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   13
         Left            =   60
         TabIndex        =   469
         Top             =   8220
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   14
         Left            =   60
         TabIndex        =   468
         Top             =   8580
         Width           =   675
      End
      Begin VB.Label lblOT2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   15
         Left            =   60
         TabIndex        =   467
         Top             =   8940
         Width           =   675
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   0
         Left            =   2310
         TabIndex        =   466
         Top             =   2790
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   1
         Left            =   2310
         TabIndex        =   465
         Top             =   3150
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   2
         Left            =   2310
         TabIndex        =   464
         Top             =   3540
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   3
         Left            =   2310
         TabIndex        =   463
         Top             =   3960
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   4
         Left            =   2310
         TabIndex        =   462
         Top             =   4410
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   461
         Top             =   4830
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   460
         Top             =   5220
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   459
         Top             =   5700
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   8
         Left            =   2280
         TabIndex        =   458
         Top             =   6120
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   9
         Left            =   2220
         TabIndex        =   457
         Top             =   6570
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   10
         Left            =   2250
         TabIndex        =   456
         Top             =   6960
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   11
         Left            =   2220
         TabIndex        =   455
         Top             =   7380
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   12
         Left            =   2220
         TabIndex        =   454
         Top             =   7830
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   13
         Left            =   2220
         TabIndex        =   453
         Top             =   8220
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   14
         Left            =   2220
         TabIndex        =   452
         Top             =   8610
         Width           =   1305
      End
      Begin VB.Label lblUT2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   15
         Left            =   2190
         TabIndex        =   451
         Top             =   8940
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   0
         Left            =   3690
         TabIndex        =   450
         Top             =   2790
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   1
         Left            =   3690
         TabIndex        =   449
         Top             =   3150
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   2
         Left            =   3690
         TabIndex        =   448
         Top             =   3570
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   3
         Left            =   3750
         TabIndex        =   447
         Top             =   3930
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   4
         Left            =   3750
         TabIndex        =   446
         Top             =   4380
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   5
         Left            =   3720
         TabIndex        =   445
         Top             =   4800
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   6
         Left            =   3750
         TabIndex        =   444
         Top             =   5220
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   7
         Left            =   3720
         TabIndex        =   443
         Top             =   5670
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   8
         Left            =   3780
         TabIndex        =   442
         Top             =   6150
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   9
         Left            =   3750
         TabIndex        =   441
         Top             =   6570
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   10
         Left            =   3750
         TabIndex        =   440
         Top             =   6930
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   11
         Left            =   3720
         TabIndex        =   439
         Top             =   7410
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   12
         Left            =   3750
         TabIndex        =   438
         Top             =   7830
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   13
         Left            =   3750
         TabIndex        =   437
         Top             =   8250
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   14
         Left            =   3720
         TabIndex        =   436
         Top             =   8610
         Width           =   1305
      End
      Begin VB.Label lblLEV2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   15
         Left            =   3690
         TabIndex        =   435
         Top             =   8940
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   434
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   1
         Left            =   5130
         TabIndex        =   433
         Top             =   3210
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   432
         Top             =   3600
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   3
         Left            =   5160
         TabIndex        =   431
         Top             =   3930
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   430
         Top             =   4410
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   5
         Left            =   5220
         TabIndex        =   429
         Top             =   4800
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   6
         Left            =   5190
         TabIndex        =   428
         Top             =   5220
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   7
         Left            =   5220
         TabIndex        =   427
         Top             =   5640
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   8
         Left            =   5250
         TabIndex        =   426
         Top             =   6120
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   9
         Left            =   5160
         TabIndex        =   425
         Top             =   6600
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   10
         Left            =   5190
         TabIndex        =   424
         Top             =   6960
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   11
         Left            =   5190
         TabIndex        =   423
         Top             =   7440
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   12
         Left            =   5160
         TabIndex        =   422
         Top             =   7860
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   13
         Left            =   5190
         TabIndex        =   421
         Top             =   8250
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   14
         Left            =   5190
         TabIndex        =   420
         Top             =   8580
         Width           =   1305
      End
      Begin VB.Label lblNO2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   15
         Left            =   5190
         TabIndex        =   419
         Top             =   8970
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   0
         Left            =   6600
         TabIndex        =   418
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   417
         Top             =   3180
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   2
         Left            =   6600
         TabIndex        =   416
         Top             =   3540
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   3
         Left            =   6630
         TabIndex        =   415
         Top             =   3870
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   4
         Left            =   6690
         TabIndex        =   414
         Top             =   4380
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   5
         Left            =   6600
         TabIndex        =   413
         Top             =   4800
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   6
         Left            =   6600
         TabIndex        =   412
         Top             =   5220
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   7
         Left            =   6660
         TabIndex        =   411
         Top             =   5640
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   8
         Left            =   6630
         TabIndex        =   410
         Top             =   6150
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   10
         Left            =   6600
         TabIndex        =   409
         Top             =   6990
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   11
         Left            =   6630
         TabIndex        =   408
         Top             =   7380
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   12
         Left            =   6600
         TabIndex        =   407
         Top             =   7800
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   13
         Left            =   6600
         TabIndex        =   406
         Top             =   8250
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   9
         Left            =   6630
         TabIndex        =   405
         Top             =   6600
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   14
         Left            =   6630
         TabIndex        =   404
         Top             =   8610
         Width           =   1305
      End
      Begin VB.Label lblHOL2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   15
         Left            =   6660
         TabIndex        =   403
         Top             =   9000
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   0
         Left            =   8010
         TabIndex        =   402
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   1
         Left            =   8040
         TabIndex        =   401
         Top             =   3210
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   2
         Left            =   8070
         TabIndex        =   400
         Top             =   3570
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   3
         Left            =   8100
         TabIndex        =   399
         Top             =   3930
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   4
         Left            =   8100
         TabIndex        =   398
         Top             =   4440
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   5
         Left            =   8130
         TabIndex        =   397
         Top             =   4890
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   6
         Left            =   8160
         TabIndex        =   396
         Top             =   5250
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   7
         Left            =   8190
         TabIndex        =   395
         Top             =   5610
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   8
         Left            =   8070
         TabIndex        =   394
         Top             =   6150
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   9
         Left            =   8040
         TabIndex        =   393
         Top             =   6570
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   10
         Left            =   8100
         TabIndex        =   392
         Top             =   6930
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   11
         Left            =   8100
         TabIndex        =   391
         Top             =   7410
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   12
         Left            =   8100
         TabIndex        =   390
         Top             =   7890
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   13
         Left            =   8130
         TabIndex        =   389
         Top             =   8280
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   14
         Left            =   8130
         TabIndex        =   388
         Top             =   8670
         Width           =   1305
      End
      Begin VB.Label lblIT2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   15
         Left            =   8160
         TabIndex        =   387
         Top             =   9000
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   0
         Left            =   9570
         TabIndex        =   386
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   1
         Left            =   9570
         TabIndex        =   385
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   2
         Left            =   9510
         TabIndex        =   384
         Top             =   3600
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   3
         Left            =   9510
         TabIndex        =   383
         Top             =   4080
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   4
         Left            =   9540
         TabIndex        =   382
         Top             =   4440
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   5
         Left            =   9540
         TabIndex        =   381
         Top             =   4920
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   6
         Left            =   9540
         TabIndex        =   380
         Top             =   5280
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   7
         Left            =   9540
         TabIndex        =   379
         Top             =   5670
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   8
         Left            =   9600
         TabIndex        =   378
         Top             =   6180
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   9
         Left            =   9600
         TabIndex        =   377
         Top             =   6540
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   10
         Left            =   9570
         TabIndex        =   376
         Top             =   6900
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   11
         Left            =   9570
         TabIndex        =   375
         Top             =   7380
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   12
         Left            =   9540
         TabIndex        =   374
         Top             =   7800
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   13
         Left            =   9540
         TabIndex        =   373
         Top             =   8280
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   14
         Left            =   9570
         TabIndex        =   372
         Top             =   8640
         Width           =   1305
      End
      Begin VB.Label lblMIL2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   15
         Left            =   9570
         TabIndex        =   371
         Top             =   8970
         Width           =   1305
      End
   End
   Begin VB.PictureBox picNOTES 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   14520
      ScaleHeight     =   2355
      ScaleWidth      =   75
      TabIndex        =   248
      Top             =   2640
      Visible         =   0   'False
      Width           =   105
      Begin VB.ComboBox cboType1 
         Appearance      =   0  'Flat
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   522
         Top             =   1140
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.TextBox txtOT1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   486
         Top             =   1560
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton cmdCancelNote 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   2310
         TabIndex        =   251
         Top             =   1950
         Width           =   1005
      End
      Begin VB.CommandButton cmdSaveNote 
         Caption         =   "Save"
         Height          =   345
         Left            =   1290
         TabIndex        =   250
         Top             =   1950
         Width           =   1005
      End
      Begin VB.TextBox txtNOTES 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   249
         Top             =   330
         Width           =   3195
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "notes for mil"
         Height          =   315
         Index           =   8
         Left            =   9690
         TabIndex        =   564
         Top             =   2760
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "notes for it"
         Height          =   315
         Index           =   7
         Left            =   8100
         TabIndex        =   563
         Top             =   2730
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "notes for koliday"
         Height          =   435
         Index           =   6
         Left            =   6750
         TabIndex        =   562
         Top             =   2730
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "notes for nigth ot"
         Height          =   315
         Index           =   5
         Left            =   5310
         TabIndex        =   561
         Top             =   2790
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "notes for liv"
         Height          =   315
         Index           =   4
         Left            =   3990
         TabIndex        =   560
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "notes for ut"
         Height          =   315
         Index           =   3
         Left            =   2700
         TabIndex        =   559
         Top             =   2730
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "no of ot hr"
         Height          =   435
         Index           =   2
         Left            =   1710
         TabIndex        =   558
         Top             =   2700
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "ot type"
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   557
         Top             =   2700
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LblCaption 
         BackColor       =   &H000000FF&
         Caption         =   "notes for OT"
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   556
         Top             =   2700
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   15
         Left            =   1680
         TabIndex        =   539
         Top             =   9360
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   14
         Left            =   1680
         TabIndex        =   538
         Top             =   9000
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   13
         Left            =   1680
         TabIndex        =   537
         Top             =   8640
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   1680
         TabIndex        =   536
         Top             =   8250
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   11
         Left            =   1710
         TabIndex        =   535
         Top             =   7800
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   1740
         TabIndex        =   534
         Top             =   7410
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   1710
         TabIndex        =   533
         Top             =   7050
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   8
         Left            =   1710
         TabIndex        =   532
         Top             =   6600
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   1740
         TabIndex        =   531
         Top             =   6120
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1710
         TabIndex        =   530
         Top             =   5670
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1740
         TabIndex        =   529
         Top             =   5250
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   528
         Top             =   4830
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   527
         Top             =   4410
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1710
         TabIndex        =   526
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   525
         Top             =   3600
         Width           =   525
      End
      Begin VB.Label lblOTCode1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1710
         TabIndex        =   524
         Top             =   3240
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   0
         Left            =   990
         TabIndex        =   505
         Top             =   3240
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   504
         Top             =   3600
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   2
         Left            =   990
         TabIndex        =   503
         Top             =   3990
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   502
         Top             =   4410
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   501
         Top             =   4830
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   5
         Left            =   1020
         TabIndex        =   500
         Top             =   5250
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   6
         Left            =   990
         TabIndex        =   499
         Top             =   5670
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   7
         Left            =   1020
         TabIndex        =   498
         Top             =   6120
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   8
         Left            =   990
         TabIndex        =   497
         Top             =   6600
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   9
         Left            =   990
         TabIndex        =   496
         Top             =   7050
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   10
         Left            =   1020
         TabIndex        =   495
         Top             =   7410
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   11
         Left            =   990
         TabIndex        =   494
         Top             =   7800
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   12
         Left            =   960
         TabIndex        =   493
         Top             =   8250
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   13
         Left            =   960
         TabIndex        =   492
         Top             =   8640
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   14
         Left            =   960
         TabIndex        =   491
         Top             =   9000
         Width           =   525
      End
      Begin VB.Label lblOTno1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   15
         Left            =   960
         TabIndex        =   490
         Top             =   9360
         Width           =   525
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   15
         Left            =   9720
         TabIndex        =   366
         Top             =   9420
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   14
         Left            =   9720
         TabIndex        =   365
         Top             =   9090
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   13
         Left            =   9690
         TabIndex        =   364
         Top             =   8730
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   12
         Left            =   9690
         TabIndex        =   363
         Top             =   8250
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   11
         Left            =   9720
         TabIndex        =   362
         Top             =   7830
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   10
         Left            =   9720
         TabIndex        =   361
         Top             =   7350
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   9
         Left            =   9750
         TabIndex        =   360
         Top             =   6990
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   8
         Left            =   9750
         TabIndex        =   359
         Top             =   6630
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   7
         Left            =   9690
         TabIndex        =   358
         Top             =   6120
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   6
         Left            =   9690
         TabIndex        =   357
         Top             =   5730
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   5
         Left            =   9690
         TabIndex        =   356
         Top             =   5370
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   4
         Left            =   9690
         TabIndex        =   355
         Top             =   4890
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   3
         Left            =   9660
         TabIndex        =   354
         Top             =   4530
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   2
         Left            =   9660
         TabIndex        =   353
         Top             =   4050
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   1
         Left            =   9720
         TabIndex        =   352
         Top             =   3690
         Width           =   1305
      End
      Begin VB.Label lblMIL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MIL"
         Height          =   285
         Index           =   0
         Left            =   9720
         TabIndex        =   351
         Top             =   3210
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   15
         Left            =   8310
         TabIndex        =   350
         Top             =   9450
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   14
         Left            =   8280
         TabIndex        =   349
         Top             =   9120
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   13
         Left            =   8280
         TabIndex        =   348
         Top             =   8730
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   12
         Left            =   8250
         TabIndex        =   347
         Top             =   8340
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   11
         Left            =   8250
         TabIndex        =   346
         Top             =   7860
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   10
         Left            =   8250
         TabIndex        =   345
         Top             =   7380
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   9
         Left            =   8190
         TabIndex        =   344
         Top             =   7020
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   8
         Left            =   8220
         TabIndex        =   343
         Top             =   6600
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   7
         Left            =   8340
         TabIndex        =   342
         Top             =   6060
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   6
         Left            =   8310
         TabIndex        =   341
         Top             =   5700
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   5
         Left            =   8280
         TabIndex        =   340
         Top             =   5340
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   4
         Left            =   8250
         TabIndex        =   339
         Top             =   4890
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   3
         Left            =   8250
         TabIndex        =   338
         Top             =   4380
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   2
         Left            =   8220
         TabIndex        =   337
         Top             =   4020
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   1
         Left            =   8190
         TabIndex        =   336
         Top             =   3660
         Width           =   1305
      End
      Begin VB.Label lblIT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "IT"
         Height          =   285
         Index           =   0
         Left            =   8160
         TabIndex        =   335
         Top             =   3210
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   15
         Left            =   6810
         TabIndex        =   334
         Top             =   9450
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   14
         Left            =   6780
         TabIndex        =   333
         Top             =   9060
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   9
         Left            =   6780
         TabIndex        =   332
         Top             =   7050
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   13
         Left            =   6750
         TabIndex        =   331
         Top             =   8700
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   12
         Left            =   6750
         TabIndex        =   330
         Top             =   8250
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   11
         Left            =   6780
         TabIndex        =   329
         Top             =   7830
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   10
         Left            =   6750
         TabIndex        =   328
         Top             =   7440
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   8
         Left            =   6780
         TabIndex        =   327
         Top             =   6600
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   7
         Left            =   6810
         TabIndex        =   326
         Top             =   6090
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   6
         Left            =   6750
         TabIndex        =   325
         Top             =   5670
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   5
         Left            =   6750
         TabIndex        =   324
         Top             =   5250
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   4
         Left            =   6840
         TabIndex        =   323
         Top             =   4830
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   3
         Left            =   6780
         TabIndex        =   322
         Top             =   4320
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   2
         Left            =   6750
         TabIndex        =   321
         Top             =   3990
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   1
         Left            =   6750
         TabIndex        =   320
         Top             =   3630
         Width           =   1305
      End
      Begin VB.Label lblHOL 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HOL"
         Height          =   285
         Index           =   0
         Left            =   6750
         TabIndex        =   319
         Top             =   3210
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   15
         Left            =   5340
         TabIndex        =   318
         Top             =   9420
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   14
         Left            =   5340
         TabIndex        =   317
         Top             =   9030
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   13
         Left            =   5340
         TabIndex        =   316
         Top             =   8700
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   12
         Left            =   5310
         TabIndex        =   315
         Top             =   8310
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   11
         Left            =   5340
         TabIndex        =   314
         Top             =   7890
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   10
         Left            =   5340
         TabIndex        =   313
         Top             =   7410
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   9
         Left            =   5310
         TabIndex        =   312
         Top             =   7050
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   8
         Left            =   5400
         TabIndex        =   311
         Top             =   6570
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   7
         Left            =   5370
         TabIndex        =   310
         Top             =   6090
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   6
         Left            =   5340
         TabIndex        =   309
         Top             =   5670
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   5
         Left            =   5370
         TabIndex        =   308
         Top             =   5250
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   4
         Left            =   5430
         TabIndex        =   307
         Top             =   4860
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   3
         Left            =   5310
         TabIndex        =   306
         Top             =   4380
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   2
         Left            =   5310
         TabIndex        =   305
         Top             =   4050
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   304
         Top             =   3660
         Width           =   1305
      End
      Begin VB.Label lblNO 
         BackColor       =   &H00C0FFFF&
         Caption         =   "NO"
         Height          =   285
         Index           =   0
         Left            =   5310
         TabIndex        =   303
         Top             =   3210
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   15
         Left            =   3840
         TabIndex        =   302
         Top             =   9390
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   14
         Left            =   3870
         TabIndex        =   301
         Top             =   9060
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   13
         Left            =   3900
         TabIndex        =   300
         Top             =   8700
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   12
         Left            =   3900
         TabIndex        =   299
         Top             =   8280
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   11
         Left            =   3870
         TabIndex        =   298
         Top             =   7860
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   10
         Left            =   3900
         TabIndex        =   297
         Top             =   7380
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   9
         Left            =   3900
         TabIndex        =   296
         Top             =   7020
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   8
         Left            =   3930
         TabIndex        =   295
         Top             =   6600
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   7
         Left            =   3870
         TabIndex        =   294
         Top             =   6120
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   6
         Left            =   3900
         TabIndex        =   293
         Top             =   5670
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   5
         Left            =   3870
         TabIndex        =   292
         Top             =   5250
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   4
         Left            =   3900
         TabIndex        =   291
         Top             =   4830
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   3
         Left            =   3900
         TabIndex        =   290
         Top             =   4380
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   289
         Top             =   4020
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   288
         Top             =   3600
         Width           =   1305
      End
      Begin VB.Label lblLEV 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LIV"
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   287
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   15
         Left            =   2340
         TabIndex        =   286
         Top             =   9390
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   14
         Left            =   2370
         TabIndex        =   285
         Top             =   9060
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   13
         Left            =   2370
         TabIndex        =   284
         Top             =   8670
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   12
         Left            =   2370
         TabIndex        =   283
         Top             =   8280
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   11
         Left            =   2370
         TabIndex        =   282
         Top             =   7830
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   10
         Left            =   2400
         TabIndex        =   281
         Top             =   7410
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   9
         Left            =   2370
         TabIndex        =   280
         Top             =   7020
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   8
         Left            =   2430
         TabIndex        =   279
         Top             =   6570
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   7
         Left            =   2430
         TabIndex        =   278
         Top             =   6150
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   6
         Left            =   2430
         TabIndex        =   277
         Top             =   5670
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   5
         Left            =   2430
         TabIndex        =   276
         Top             =   5280
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   4
         Left            =   2460
         TabIndex        =   275
         Top             =   4860
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   3
         Left            =   2460
         TabIndex        =   274
         Top             =   4410
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   2
         Left            =   2460
         TabIndex        =   273
         Top             =   3990
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   1
         Left            =   2460
         TabIndex        =   272
         Top             =   3600
         Width           =   1305
      End
      Begin VB.Label lblUT 
         BackColor       =   &H00C0C0FF&
         Caption         =   "UT"
         Height          =   285
         Index           =   0
         Left            =   2460
         TabIndex        =   271
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   15
         Left            =   30
         TabIndex        =   270
         Top             =   9360
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   14
         Left            =   30
         TabIndex        =   269
         Top             =   9000
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   13
         Left            =   30
         TabIndex        =   268
         Top             =   8640
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   12
         Left            =   30
         TabIndex        =   267
         Top             =   8250
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   11
         Left            =   60
         TabIndex        =   266
         Top             =   7800
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   10
         Left            =   90
         TabIndex        =   265
         Top             =   7410
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   9
         Left            =   60
         TabIndex        =   264
         Top             =   7050
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   8
         Left            =   60
         TabIndex        =   263
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   7
         Left            =   90
         TabIndex        =   262
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   6
         Left            =   60
         TabIndex        =   261
         Top             =   5670
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   260
         Top             =   5250
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   4
         Left            =   30
         TabIndex        =   259
         Top             =   4830
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   3
         Left            =   30
         TabIndex        =   258
         Top             =   4410
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   257
         Top             =   3990
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   1
         Left            =   30
         TabIndex        =   256
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblOT 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OT"
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   255
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblIndex 
         BackColor       =   &H000000FF&
         Height          =   225
         Left            =   1170
         TabIndex        =   254
         Top             =   60
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblNoteTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   90
         TabIndex        =   253
         Top             =   90
         Width           =   465
      End
      Begin VB.Label lblCap 
         BackColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   2730
         TabIndex        =   252
         Top             =   60
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblNoOfOT1 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1740
         TabIndex        =   488
         Top             =   1620
         Visible         =   0   'False
         Width           =   930
      End
   End
   Begin VB.PictureBox picBaba 
      BackColor       =   &H00F29800&
      Enabled         =   0   'False
      Height          =   4275
      Left            =   14670
      ScaleHeight     =   4215
      ScaleWidth      =   375
      TabIndex        =   116
      Top             =   5700
      Width           =   435
      Begin VB.HScrollBar HScroll2 
         Height          =   225
         LargeChange     =   1000
         Left            =   90
         Max             =   5000
         SmallChange     =   1000
         TabIndex        =   215
         Top             =   3930
         Width           =   4935
      End
      Begin VB.PictureBox pic16to31 
         BackColor       =   &H00F29800&
         BorderStyle     =   0  'None
         Height          =   4185
         Left            =   -4950
         ScaleHeight     =   4185
         ScaleWidth      =   9015
         TabIndex        =   117
         Top             =   0
         Width           =   9015
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   15
            Left            =   6750
            TabIndex        =   247
            Top             =   3630
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   14
            Left            =   6750
            TabIndex        =   246
            Top             =   3390
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   13
            Left            =   6750
            TabIndex        =   245
            Top             =   3150
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   12
            Left            =   6750
            TabIndex        =   244
            Top             =   2910
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   11
            Left            =   6750
            TabIndex        =   243
            Top             =   2670
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   10
            Left            =   6750
            TabIndex        =   242
            Top             =   2430
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   9
            Left            =   6750
            TabIndex        =   241
            Top             =   2190
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   8
            Left            =   6750
            TabIndex        =   240
            Top             =   1950
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   7
            Left            =   6750
            TabIndex        =   239
            Top             =   1710
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   6
            Left            =   6750
            TabIndex        =   238
            Top             =   1470
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   5
            Left            =   6750
            TabIndex        =   237
            Top             =   1230
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   4
            Left            =   6750
            TabIndex        =   236
            Top             =   990
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   3
            Left            =   6750
            TabIndex        =   235
            Top             =   750
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   2
            Left            =   6750
            TabIndex        =   234
            Top             =   510
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   1
            Left            =   6750
            TabIndex        =   233
            Top             =   270
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA2 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   0
            Left            =   6750
            TabIndex        =   232
            Top             =   30
            Width           =   1845
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   15
            Left            =   5940
            TabIndex        =   213
            Top             =   3630
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   14
            Left            =   5940
            TabIndex        =   212
            Top             =   3390
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   13
            Left            =   5940
            TabIndex        =   211
            Top             =   3150
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   12
            Left            =   5940
            TabIndex        =   210
            Top             =   2910
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   11
            Left            =   5940
            TabIndex        =   209
            Top             =   2670
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   10
            Left            =   5940
            TabIndex        =   208
            Top             =   2430
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   9
            Left            =   5940
            TabIndex        =   207
            Top             =   2190
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   8
            Left            =   5940
            TabIndex        =   206
            Top             =   1950
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   7
            Left            =   5940
            TabIndex        =   205
            Top             =   1710
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   6
            Left            =   5940
            TabIndex        =   204
            Top             =   1470
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   5
            Left            =   5940
            TabIndex        =   203
            Top             =   1230
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   4
            Left            =   5940
            TabIndex        =   202
            Top             =   990
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   3
            Left            =   5940
            TabIndex        =   201
            Top             =   750
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   2
            Left            =   5940
            TabIndex        =   200
            Top             =   510
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   1
            Left            =   5940
            TabIndex        =   199
            Top             =   270
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT2 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   0
            Left            =   5940
            TabIndex        =   198
            Top             =   30
            Width           =   795
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   0
            Left            =   4860
            TabIndex        =   197
            Top             =   30
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   1
            Left            =   4860
            TabIndex        =   196
            Top             =   270
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   2
            Left            =   4860
            TabIndex        =   195
            Top             =   510
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   3
            Left            =   4860
            TabIndex        =   194
            Top             =   750
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   4
            Left            =   4860
            TabIndex        =   193
            Top             =   990
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   5
            Left            =   4860
            TabIndex        =   192
            Top             =   1230
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   6
            Left            =   4860
            TabIndex        =   191
            Top             =   1470
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   7
            Left            =   4860
            TabIndex        =   190
            Top             =   1710
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   8
            Left            =   4860
            TabIndex        =   189
            Top             =   1950
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   9
            Left            =   4860
            TabIndex        =   188
            Top             =   2190
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   10
            Left            =   4860
            TabIndex        =   187
            Top             =   2430
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   11
            Left            =   4860
            TabIndex        =   186
            Top             =   2670
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   12
            Left            =   4860
            TabIndex        =   185
            Top             =   2910
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   13
            Left            =   4860
            TabIndex        =   184
            Top             =   3150
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   14
            Left            =   4860
            TabIndex        =   183
            Top             =   3390
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol2 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   15
            Left            =   4860
            TabIndex        =   182
            Top             =   3630
            Width           =   1035
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   0
            Left            =   3570
            TabIndex        =   181
            Top             =   30
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   1
            Left            =   3570
            TabIndex        =   180
            Top             =   270
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   2
            Left            =   3570
            TabIndex        =   179
            Top             =   510
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   3
            Left            =   3570
            TabIndex        =   178
            Top             =   750
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   4
            Left            =   3570
            TabIndex        =   177
            Top             =   990
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   5
            Left            =   3570
            TabIndex        =   176
            Top             =   1230
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   6
            Left            =   3570
            TabIndex        =   175
            Top             =   1470
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   7
            Left            =   3570
            TabIndex        =   174
            Top             =   1710
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   8
            Left            =   3570
            TabIndex        =   173
            Top             =   1950
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   9
            Left            =   3570
            TabIndex        =   172
            Top             =   2190
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   10
            Left            =   3570
            TabIndex        =   171
            Top             =   2430
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   11
            Left            =   3570
            TabIndex        =   170
            Top             =   2670
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   12
            Left            =   3570
            TabIndex        =   169
            Top             =   2910
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   13
            Left            =   3570
            TabIndex        =   168
            Top             =   3150
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   14
            Left            =   3570
            TabIndex        =   167
            Top             =   3390
            Width           =   1245
         End
         Begin VB.CheckBox ChkND2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   15
            Left            =   3570
            TabIndex        =   166
            Top             =   3630
            Width           =   1245
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   0
            Left            =   2490
            TabIndex        =   165
            Top             =   30
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   1
            Left            =   2490
            TabIndex        =   164
            Top             =   270
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   2
            Left            =   2490
            TabIndex        =   163
            Top             =   510
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   3
            Left            =   2490
            TabIndex        =   162
            Top             =   750
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   4
            Left            =   2490
            TabIndex        =   161
            Top             =   990
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   5
            Left            =   2490
            TabIndex        =   160
            Top             =   1230
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   6
            Left            =   2490
            TabIndex        =   159
            Top             =   1470
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   7
            Left            =   2490
            TabIndex        =   158
            Top             =   1710
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   8
            Left            =   2490
            TabIndex        =   157
            Top             =   1950
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   9
            Left            =   2490
            TabIndex        =   156
            Top             =   2190
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   10
            Left            =   2490
            TabIndex        =   155
            Top             =   2430
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   11
            Left            =   2490
            TabIndex        =   154
            Top             =   2670
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   12
            Left            =   2490
            TabIndex        =   153
            Top             =   2910
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   13
            Left            =   2490
            TabIndex        =   152
            Top             =   3150
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   14
            Left            =   2490
            TabIndex        =   151
            Top             =   3390
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL2 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   15
            Left            =   2490
            TabIndex        =   150
            Top             =   3630
            Width           =   1065
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   149
            Top             =   30
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   30
            TabIndex        =   148
            Top             =   270
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   2
            Left            =   30
            TabIndex        =   147
            Top             =   510
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   3
            Left            =   30
            TabIndex        =   146
            Top             =   750
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   4
            Left            =   30
            TabIndex        =   145
            Top             =   990
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   5
            Left            =   30
            TabIndex        =   144
            Top             =   1230
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   6
            Left            =   30
            TabIndex        =   143
            Top             =   1470
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   7
            Left            =   30
            TabIndex        =   142
            Top             =   1710
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   8
            Left            =   30
            TabIndex        =   141
            Top             =   1950
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   9
            Left            =   30
            TabIndex        =   140
            Top             =   2190
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   10
            Left            =   30
            TabIndex        =   139
            Top             =   2430
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   11
            Left            =   30
            TabIndex        =   138
            Top             =   2670
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   12
            Left            =   30
            TabIndex        =   137
            Top             =   2910
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   13
            Left            =   30
            TabIndex        =   136
            Top             =   3150
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   14
            Left            =   30
            TabIndex        =   135
            Top             =   3390
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   15
            Left            =   30
            TabIndex        =   134
            Top             =   3630
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   15
            Left            =   1260
            TabIndex        =   133
            Top             =   3630
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   14
            Left            =   1260
            TabIndex        =   132
            Top             =   3390
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   13
            Left            =   1260
            TabIndex        =   131
            Top             =   3150
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   12
            Left            =   1260
            TabIndex        =   130
            Top             =   2910
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   11
            Left            =   1260
            TabIndex        =   129
            Top             =   2670
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   10
            Left            =   1260
            TabIndex        =   128
            Top             =   2430
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   9
            Left            =   1260
            TabIndex        =   127
            Top             =   2190
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   8
            Left            =   1260
            TabIndex        =   126
            Top             =   1950
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   7
            Left            =   1260
            TabIndex        =   125
            Top             =   1710
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   6
            Left            =   1260
            TabIndex        =   124
            Top             =   1470
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   5
            Left            =   1260
            TabIndex        =   123
            Top             =   1230
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   4
            Left            =   1260
            TabIndex        =   122
            Top             =   990
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   1260
            TabIndex        =   121
            Top             =   750
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   2
            Left            =   1260
            TabIndex        =   120
            Top             =   510
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   1
            Left            =   1260
            TabIndex        =   119
            Top             =   270
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT2 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   0
            Left            =   1260
            TabIndex        =   118
            Top             =   30
            Width           =   1245
         End
      End
   End
   Begin VB.PictureBox picTaas 
      BackColor       =   &H00F29800&
      Enabled         =   0   'False
      Height          =   4275
      Left            =   14670
      ScaleHeight     =   4215
      ScaleWidth      =   405
      TabIndex        =   18
      Top             =   750
      Width           =   465
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         LargeChange     =   1000
         Left            =   30
         Max             =   5000
         SmallChange     =   1000
         TabIndex        =   214
         Top             =   3930
         Width           =   4965
      End
      Begin VB.PictureBox pic1To15 
         BackColor       =   &H00F29800&
         BorderStyle     =   0  'None
         Height          =   3945
         Left            =   -4950
         ScaleHeight     =   3945
         ScaleWidth      =   8625
         TabIndex        =   19
         Top             =   -60
         Width           =   8625
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   15
            Left            =   6780
            TabIndex        =   231
            Top             =   3690
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   14
            Left            =   6780
            TabIndex        =   230
            Top             =   3450
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   13
            Left            =   6780
            TabIndex        =   229
            Top             =   3210
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   12
            Left            =   6780
            TabIndex        =   228
            Top             =   2970
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   11
            Left            =   6780
            TabIndex        =   227
            Top             =   2730
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   10
            Left            =   6780
            TabIndex        =   226
            Top             =   2490
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   9
            Left            =   6780
            TabIndex        =   225
            Top             =   2250
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   8
            Left            =   6780
            TabIndex        =   224
            Top             =   2010
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   7
            Left            =   6780
            TabIndex        =   223
            Top             =   1770
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   6
            Left            =   6780
            TabIndex        =   222
            Top             =   1530
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   5
            Left            =   6780
            TabIndex        =   221
            Top             =   1290
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   4
            Left            =   6780
            TabIndex        =   220
            Top             =   1050
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   3
            Left            =   6780
            TabIndex        =   219
            Top             =   810
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   2
            Left            =   6780
            TabIndex        =   218
            Top             =   570
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   1
            Left            =   6780
            TabIndex        =   217
            Top             =   330
            Width           =   1845
         End
         Begin VB.CheckBox ChkwMA1 
            BackColor       =   &H00F29800&
            Caption         =   "with Meal Allowance"
            Enabled         =   0   'False
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   0
            Left            =   6780
            TabIndex        =   216
            Top             =   90
            Width           =   1845
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   0
            Left            =   1260
            TabIndex        =   115
            Top             =   90
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   1
            Left            =   1260
            TabIndex        =   114
            Top             =   330
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   2
            Left            =   1260
            TabIndex        =   113
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   1260
            TabIndex        =   112
            Top             =   810
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   4
            Left            =   1260
            TabIndex        =   111
            Top             =   1050
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   5
            Left            =   1260
            TabIndex        =   110
            Top             =   1290
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   6
            Left            =   1260
            TabIndex        =   109
            Top             =   1530
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   7
            Left            =   1260
            TabIndex        =   108
            Top             =   1770
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   8
            Left            =   1260
            TabIndex        =   107
            Top             =   2010
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   9
            Left            =   1260
            TabIndex        =   106
            Top             =   2250
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   10
            Left            =   1260
            TabIndex        =   105
            Top             =   2490
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   11
            Left            =   1260
            TabIndex        =   104
            Top             =   2730
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   12
            Left            =   1260
            TabIndex        =   103
            Top             =   2970
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   13
            Left            =   1260
            TabIndex        =   102
            Top             =   3210
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   14
            Left            =   1260
            TabIndex        =   101
            Top             =   3450
            Width           =   1245
         End
         Begin VB.CheckBox ChkUT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize UT"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   15
            Left            =   1260
            TabIndex        =   100
            Top             =   3690
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   15
            Left            =   30
            TabIndex        =   99
            Top             =   3690
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   14
            Left            =   30
            TabIndex        =   98
            Top             =   3450
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   13
            Left            =   30
            TabIndex        =   97
            Top             =   3210
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   12
            Left            =   30
            TabIndex        =   96
            Top             =   2970
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   11
            Left            =   30
            TabIndex        =   95
            Top             =   2730
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   10
            Left            =   30
            TabIndex        =   94
            Top             =   2490
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   9
            Left            =   30
            TabIndex        =   93
            Top             =   2250
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   8
            Left            =   30
            TabIndex        =   92
            Top             =   2010
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   7
            Left            =   30
            TabIndex        =   91
            Top             =   1770
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   6
            Left            =   30
            TabIndex        =   90
            Top             =   1530
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   5
            Left            =   30
            TabIndex        =   89
            Top             =   1290
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   4
            Left            =   30
            TabIndex        =   88
            Top             =   1050
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   3
            Left            =   30
            TabIndex        =   87
            Top             =   810
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   2
            Left            =   30
            TabIndex        =   86
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   30
            TabIndex        =   85
            Top             =   330
            Width           =   1245
         End
         Begin VB.CheckBox ChkOT1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize OT"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   84
            Top             =   90
            Width           =   1245
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   15
            Left            =   2490
            TabIndex        =   83
            Top             =   3690
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   14
            Left            =   2490
            TabIndex        =   82
            Top             =   3450
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   13
            Left            =   2490
            TabIndex        =   81
            Top             =   3210
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   12
            Left            =   2490
            TabIndex        =   80
            Top             =   2970
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   11
            Left            =   2490
            TabIndex        =   79
            Top             =   2730
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   10
            Left            =   2490
            TabIndex        =   78
            Top             =   2490
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   9
            Left            =   2490
            TabIndex        =   77
            Top             =   2250
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   8
            Left            =   2490
            TabIndex        =   76
            Top             =   2010
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   7
            Left            =   2490
            TabIndex        =   75
            Top             =   1770
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   6
            Left            =   2490
            TabIndex        =   74
            Top             =   1530
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   5
            Left            =   2490
            TabIndex        =   73
            Top             =   1290
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   4
            Left            =   2490
            TabIndex        =   72
            Top             =   1050
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   3
            Left            =   2490
            TabIndex        =   71
            Top             =   810
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   2
            Left            =   2490
            TabIndex        =   70
            Top             =   570
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   1
            Left            =   2490
            TabIndex        =   69
            Top             =   330
            Width           =   1065
         End
         Begin VB.CheckBox ChkOL1 
            BackColor       =   &H00F29800&
            Caption         =   "On Leave"
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   0
            Left            =   2490
            TabIndex        =   68
            Top             =   90
            Width           =   1065
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   15
            Left            =   3570
            TabIndex        =   67
            Top             =   3690
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   14
            Left            =   3570
            TabIndex        =   66
            Top             =   3450
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   13
            Left            =   3570
            TabIndex        =   65
            Top             =   3210
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   12
            Left            =   3570
            TabIndex        =   64
            Top             =   2970
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   11
            Left            =   3570
            TabIndex        =   63
            Top             =   2730
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   10
            Left            =   3570
            TabIndex        =   62
            Top             =   2490
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   9
            Left            =   3570
            TabIndex        =   61
            Top             =   2250
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   8
            Left            =   3570
            TabIndex        =   60
            Top             =   2010
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   7
            Left            =   3570
            TabIndex        =   59
            Top             =   1770
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   6
            Left            =   3570
            TabIndex        =   58
            Top             =   1530
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   5
            Left            =   3570
            TabIndex        =   57
            Top             =   1290
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   4
            Left            =   3570
            TabIndex        =   56
            Top             =   1050
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   3
            Left            =   3570
            TabIndex        =   55
            Top             =   810
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   2
            Left            =   3570
            TabIndex        =   54
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   1
            Left            =   3570
            TabIndex        =   53
            Top             =   330
            Width           =   1245
         End
         Begin VB.CheckBox ChkND1 
            BackColor       =   &H00F29800&
            Caption         =   "Authorize ND"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   225
            Index           =   0
            Left            =   3570
            TabIndex        =   52
            Top             =   90
            Width           =   1245
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   15
            Left            =   4860
            TabIndex        =   51
            Top             =   3690
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   14
            Left            =   4860
            TabIndex        =   50
            Top             =   3450
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   13
            Left            =   4860
            TabIndex        =   49
            Top             =   3210
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   12
            Left            =   4860
            TabIndex        =   48
            Top             =   2970
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   11
            Left            =   4860
            TabIndex        =   47
            Top             =   2730
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   10
            Left            =   4860
            TabIndex        =   46
            Top             =   2490
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   9
            Left            =   4860
            TabIndex        =   45
            Top             =   2250
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   8
            Left            =   4860
            TabIndex        =   44
            Top             =   2010
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   7
            Left            =   4860
            TabIndex        =   43
            Top             =   1770
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   6
            Left            =   4860
            TabIndex        =   42
            Top             =   1530
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   5
            Left            =   4860
            TabIndex        =   41
            Top             =   1290
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   4
            Left            =   4860
            TabIndex        =   40
            Top             =   1050
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   3
            Left            =   4860
            TabIndex        =   39
            Top             =   810
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   2
            Left            =   4860
            TabIndex        =   38
            Top             =   570
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   1
            Left            =   4860
            TabIndex        =   37
            Top             =   330
            Width           =   1035
         End
         Begin VB.CheckBox ChkHol1 
            BackColor       =   &H00F29800&
            Caption         =   "IS Holiday"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   0
            Left            =   4860
            TabIndex        =   36
            Top             =   90
            Width           =   1035
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   0
            Left            =   5940
            TabIndex        =   35
            Top             =   90
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   1
            Left            =   5940
            TabIndex        =   34
            Top             =   330
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   2
            Left            =   5940
            TabIndex        =   33
            Top             =   570
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   3
            Left            =   5940
            TabIndex        =   32
            Top             =   810
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   4
            Left            =   5940
            TabIndex        =   31
            Top             =   1050
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   5
            Left            =   5940
            TabIndex        =   30
            Top             =   1290
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   6
            Left            =   5940
            TabIndex        =   29
            Top             =   1530
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   7
            Left            =   5940
            TabIndex        =   28
            Top             =   1770
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   8
            Left            =   5940
            TabIndex        =   27
            Top             =   2010
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   9
            Left            =   5940
            TabIndex        =   26
            Top             =   2250
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   10
            Left            =   5940
            TabIndex        =   25
            Top             =   2490
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   11
            Left            =   5940
            TabIndex        =   24
            Top             =   2730
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   12
            Left            =   5940
            TabIndex        =   23
            Top             =   2970
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   13
            Left            =   5940
            TabIndex        =   22
            Top             =   3210
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   14
            Left            =   5940
            TabIndex        =   21
            Top             =   3450
            Width           =   795
         End
         Begin VB.CheckBox ChkwIT1 
            BackColor       =   &H00F29800&
            Caption         =   "with IT"
            Enabled         =   0   'False
            ForeColor       =   &H00400040&
            Height          =   225
            Index           =   15
            Left            =   5940
            TabIndex        =   20
            Top             =   3690
            Width           =   795
         End
      End
   End
   Begin VB.PictureBox FrmProgBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   90
      ScaleHeight     =   945
      ScaleWidth      =   4035
      TabIndex        =   13
      Top             =   7680
      Width           =   4065
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   450
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label LblProgBar2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   3510
         TabIndex        =   15
         Top             =   60
         Width           =   450
      End
      Begin XtremeShortcutBar.ShortcutCaption LblProgBar1 
         Height          =   315
         Left            =   0
         TabIndex        =   587
         Top             =   0
         Width           =   4035
         _Version        =   655364
         _ExtentX        =   7117
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Updating actual hour..."
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   8388608
      End
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2670
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   30
      Width           =   1515
   End
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   30
      Width           =   2505
   End
   Begin VB.TextBox TxtMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3570
      TabIndex        =   10
      Top             =   180
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5280
      TabIndex        =   3
      Top             =   1200
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5595
      Left            =   90
      Picture         =   "EditCards.frx":3C07B1
      ScaleHeight     =   5595
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   4890
      Width           =   4065
      Begin VB.CommandButton cmdAdd_1stCutoff 
         Caption         =   "Fill 1st Cut-Off Time  "
         Enabled         =   0   'False
         Height          =   795
         Left            =   540
         MouseIcon       =   "EditCards.frx":3C86BA
         MousePointer    =   99  'Custom
         Picture         =   "EditCards.frx":3C880C
         Style           =   1  'Graphical
         TabIndex        =   594
         ToolTipText     =   "Fill 1st Cut-Off Time  "
         Top             =   3870
         Width           =   1125
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D7C6B5&
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   0
         ScaleHeight     =   3765
         ScaleWidth      =   4035
         TabIndex        =   579
         Top             =   0
         Width           =   4065
         Begin wizBox.Box Box4 
            Height          =   1875
            Left            =   2160
            Top             =   330
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   3307
         End
         Begin VB.Label lblEmpno 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   270
            Left            =   90
            TabIndex        =   586
            Top             =   1740
            Width           =   1830
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Emp. No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   585
            Top             =   1530
            Width           =   705
         End
         Begin VB.Image imgDispPic 
            Height          =   1755
            Left            =   2220
            Picture         =   "EditCards.frx":3C8B1F
            Stretch         =   -1  'True
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Shift"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   2
            Left            =   90
            TabIndex        =   584
            Top             =   2580
            Width           =   840
         End
         Begin VB.Label lblShift 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   270
            Left            =   90
            TabIndex        =   583
            Top             =   2790
            Width           =   3840
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   1
            Left            =   90
            TabIndex        =   582
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label lblName 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   270
            Left            =   90
            TabIndex        =   581
            Top             =   2250
            Width           =   3840
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   315
            Left            =   0
            TabIndex        =   580
            Top             =   0
            Width           =   4035
            _Version        =   655364
            _ExtentX        =   7117
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Employee Information"
            ForeColor       =   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   4194304
         End
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "then click VIEW to refresh data."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   150
         TabIndex        =   593
         Top             =   5280
         Width           =   3345
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   30
         TabIndex        =   592
         Top             =   4830
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Double click  OUT-PM for shortcuts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   150
         TabIndex        =   591
         Top             =   5070
         Width           =   3525
      End
   End
   Begin VB.CommandButton CmdLineFill1 
      Caption         =   "Line Fill"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6060
      TabIndex        =   1
      ToolTipText     =   "Fill Lines"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton CmdFillAll1 
      Caption         =   "Fill-All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7200
      TabIndex        =   2
      ToolTipText     =   "Fill-All Lines"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1110
   End
   Begin Crystal.CrystalReport rptDepartment 
      Left            =   12990
      Top             =   9990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2940
      MouseIcon       =   "EditCards.frx":3C9E1C
      MousePointer    =   99  'Custom
      Picture         =   "EditCards.frx":3C9F6E
      Style           =   1  'Graphical
      TabIndex        =   574
      ToolTipText     =   "Print this Record"
      Top             =   3870
      Width           =   765
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2190
      Picture         =   "EditCards.frx":3CAFF0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Exit Window"
      Top             =   3870
      Width           =   765
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   1440
      Picture         =   "EditCards.frx":3CC072
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancel Entry"
      Top             =   3870
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   690
      Picture         =   "EditCards.frx":3CD0F4
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Save Entry"
      Top             =   3870
      Width           =   765
   End
   Begin FlexCell.Grid Grid2 
      Height          =   4785
      Left            =   4200
      TabIndex        =   597
      Top             =   5760
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   8440
      Appearance      =   0
      BackColorBkg    =   -2147483645
      BackColorSel    =   13811126
      Cols            =   12
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   17
      ScrollBars      =   0
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4470
      TabIndex        =   6
      Top             =   6630
      Width           =   1635
   End
   Begin VB.CommandButton CmdLineFill2 
      Caption         =   "Line Fill"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6270
      TabIndex        =   5
      ToolTipText     =   "Fill Lines"
      Top             =   7230
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton CmdFillAll2 
      Caption         =   "Fill-All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7410
      TabIndex        =   4
      ToolTipText     =   "Fill-All Lines"
      Top             =   6660
      Visible         =   0   'False
      Width           =   1110
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   578
      Top             =   5280
      Width           =   9255
      _Version        =   655364
      _ExtentX        =   16325
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "SECOND CUT-OFF SCHEDULE"
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
      ForeColor       =   4194304
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   577
      Top             =   60
      Width           =   9255
      _Version        =   655364
      _ExtentX        =   16325
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "FIRST CUT-OFF SCHEDULE"
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
      ForeColor       =   4194304
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SECOND CUT-OFF SCHEDULE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   345
      Left            =   4260
      TabIndex        =   566
      Top             =   5280
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FIRST CUT-OFF SCHEDULE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   465
      Left            =   12120
      TabIndex        =   565
      Top             =   1080
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Check Boxes to Authorize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Left            =   14670
      TabIndex        =   17
      Top             =   5130
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Check Boxes to Authorize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Left            =   14670
      TabIndex        =   16
      Top             =   120
      Width           =   435
   End
   Begin VB.Menu menuoption 
      Caption         =   "Option"
      Visible         =   0   'False
      Begin VB.Menu menuovertime 
         Caption         =   "Add Overtime"
      End
      Begin VB.Menu menuleave 
         Caption         =   "Add Leave"
      End
   End
   Begin VB.Menu mainoption 
      Caption         =   "Option"
      Visible         =   0   'False
      Begin VB.Menu mainot 
         Caption         =   "Add Overtime"
      End
      Begin VB.Menu mainleave 
         Caption         =   "Add Leave"
      End
   End
End
Attribute VB_Name = "frmHRMSEditCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TransactionDate                                                   As Date
Dim rsEmpInfo                                                         As ADODB.Recordset
Dim rsCard                                                            As ADODB.Recordset
Dim rsAttend                                                          As ADODB.Recordset
Dim varEmpNo                                                          As String
Const DefaultColor = &HF29800
Dim DontClickOption                                                   As Boolean
Dim rsNOTES                                                           As New ADODB.Recordset
Dim WithEvents frms                                                   As frmHRMS_RequestForLeave_short
Attribute frms.VB_VarHelpID = -1
Dim WithEvents FRM                                                    As frmHRMS_OT
Attribute FRM.VB_VarHelpID = -1
Dim xbuwan                                                            As Date
Dim xin                                                               As Date
Dim xout                                                              As Date
Dim X                                                                 As Integer

Function IsHolyday(xmon As Integer, yday As Integer) As String
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim list                                                          As String
    
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_HOLIDAY_LIST WHERE MANTH = " & xmon & " AND DEYT = " & yday & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
       list = Null2String(RSTMP!Description)
    Else
        list = IsLeaves(Format(TransactionDate, "mm/dd/yy"))
    End If
    
    IsHolyday = list

End Function
Function IsLeaves(xmon As String) As String
    Dim leave As New ADODB.Recordset
    Dim reqdesc As String
    
    
    Set leave = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT where '" & xmon & "' between dte_from  and  dte_to  and status = 'A' and empno = '" & cboEmpNumber.Text & "'")

    If Not (leave.BOF And leave.EOF) Then
        reqdesc = Null2String(leave!reqdesc)
    Else
       reqdesc = IsOT(Format(TransactionDate, "mm/dd/yy"))
    End If

    IsLeaves = reqdesc

End Function

Function IsOT(xmon As String) As String
    Dim rsot As New ADODB.Recordset
    
    Dim OVERTIME As String
    Set rsot = gconDMIS.Execute("SELECT * FROM HRMS_OVERTIME where deyt = '" & xmon & "' and empno = '" & cboEmpNumber.Text & "'")
    
    If Not (rsot.BOF And rsot.EOF) Then
        OVERTIME = "Overtime"
    Else
        OVERTIME = ""
    End If

    IsOT = OVERTIME

End Function


Function xNumber_of_Days(xChkedTimeIn, xChkedDateAndTimeNow)    'Date_And_Time_Punched_In, Date_And_Time_Punched_In)
    Dim In_Minute                                                     As Double
    Dim chk_Time_In
    Dim chk_Date_And_Time_Out
    Dim Number_hours
    Dim Number_Days

    chk_Time_In = CDate(xChkedTimeIn)
    chk_Date_And_Time_Out = CDate(xChkedDateAndTimeNow)
    Number_hours = DateDiff("h", chk_Time_In, chk_Date_And_Time_Out)

    ' Calculate number of days between punched in and now(date).
    ' ------------------------------------------------------------
    Number_Days = DateDiff("d", chk_Time_In, chk_Date_And_Time_Out)

    If Number_Days > 1 Then
        xNumber_of_Days = Number_Days
    End If
End Function


'OLD FILLALL PLS DONT DELETE
'''''''''            Set rsCard = New ADODB.Recordset
'''''''''            'rsCard.Open "Select * from HRMS_Attend where right(Empno,4) = " & N2Str2Null(cboEmpNumber.Text) & " AND Month(DateToday) = " & Month(TransactionDate) & "AND Year(DateToday) = " & cboYear.Text, gconDMIS
'''''''''            rsCard.Open "Select * from HRMS_Attend where Empno = " & N2Str2Null(varEmpNo) & " AND Month(DateToday) = " & Month(TransactionDate) & "AND Year(DateToday) = " & cboYear.Text, gconDMIS
'''''''''            If Not rsCard.EOF And Not rsCard.BOF Then
'''''''''                rsCard.MoveFirst
'''''''''                Do Until rsCard.EOF
'''''''''                    r = Day(rsCard!DATETODAY)
'''''''''                    If r < 16 Then
'''''''''                        With Grid1
'''''''''                            .Row = r
'''''''''                            .Col = 0: .CellForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(rsCard!DATETODAY, "mm/dd/yy")
''''''''''                            If Null2String(rsCard!INAM) = "" And Null2String(rsCard!OutAm) = "" Then
''''''''''                                .Col = 1: .CellBackColor = DefaultColor
''''''''''                                .Col = 2: .CellBackColor = DefaultColor
''''''''''                                ChkOL1(.Row - 1).Enabled = True
''''''''''                                ChkOL1(.Row - 1).BackColor = vbYellow
''''''''''                            End If
''''''''''                            If Null2String(rsCard!InPm) = "" And Null2String(rsCard!OutPM) = "" Then
''''''''''                                .Col = 3: .CellBackColor = DefaultColor
''''''''''                                .Col = 4: .CellBackColor = DefaultColor
''''''''''                                ChkOL1(.Row - 1).Enabled = True
''''''''''                                ChkOL1(.Row - 1).BackColor = vbYellow
''''''''''                            End If
'''''''''                            .Col = 1:                             '.CellForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(Null2String(rsCard!INAM), "hh:mm AM/PM")
''''''''''                            If Trim(.Text) <> "" Then
''''''''''                                If .Text < #7:00:00 AM# And .Text > #6:00:00 AM# Then
''''''''''                                    .CellForeColor = vbBlue: .CellBackColor = vbYellow
''''''''''                                    ChkOT1(.Row - 1).Enabled = True
''''''''''                                    ChkOT1(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwMA1(.Row - 1).Enabled = True
''''''''''                                    ChkwMA1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text > #9:00:00 AM# And .Text < #11:00:00 AM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkUT1(.Row - 1).Enabled = True
''''''''''                                    ChkUT1(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwIT1(.Row - 1).Enabled = True
''''''''''                                    ChkwIT1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text > #8:00:00 AM# And .Text < #9:00:00 AM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkwIT1(.Row - 1).Enabled = True
''''''''''                                    ChkwIT1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                            End If
'''''''''                            .Col = 2:                             '.ForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(Null2String(rsCard!OutAm), "hh:mm AM/PM")
''''''''''                            If Trim(.Text) <> "" Then
''''''''''                                If .Text < #11:00:00 AM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkUT1(.Row - 1).Enabled = True
''''''''''                                    ChkUT1(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwIT1(.Row - 1).Enabled = True
''''''''''                                    ChkwIT1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                            End If
'''''''''                            .Col = 3:                             '.ForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(Null2String(rsCard!InPm), "hh:mm AM/PM")
''''''''''                            If Trim(.Text) <> "" Then
''''''''''                                If .Text > #2:00:00 PM# And .Text < #4:00:00 PM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkUT1(.Row - 1).Enabled = True
''''''''''                                    ChkUT1(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwIT1(.Row - 1).Enabled = True
''''''''''                                    ChkwIT1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text > #1:00:00 PM# And .Text < #2:00:00 PM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkwIT1(.Row - 1).Enabled = True
''''''''''                                    ChkwIT1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                            End If
'''''''''                            .Col = 4:                             '.ForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(Null2String(rsCard!OutPM), "hh:mm AM/PM")
''''''''''                            If Trim(.Text) <> "" Then
''''''''''                                If .Text > #6:00:00 PM# Then
''''''''''                                    .CellForeColor = vbBlue: .CellBackColor = vbYellow
''''''''''                                    ChkOT1(.Row - 1).Enabled = True
''''''''''                                    ChkOT1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text < #4:00:00 PM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkUT1(.Row - 1).Enabled = True
''''''''''                                    ChkUT1(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwIT1(.Row - 1).Enabled = True
''''''''''                                    ChkwIT1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text > #10:00:00 PM# Or .Text < #6:00:00 AM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkND1(.Row - 1).Enabled = True
''''''''''                                    ChkND1(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwMA1(.Row - 1).Enabled = True
''''''''''                                    ChkwMA1(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                            End If
''''''''''                            If Null2Bool(rsCard!AuthorizeOT) = True Then ChkOT1(.Row - 1).Value = 1 Else ChkOT1(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!AuthorizeUT) = True Then ChkUT1(.Row - 1).Value = 1 Else ChkUT1(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!AbsentWLeave) = True Then ChkOL1(.Row - 1).Value = 1 Else ChkOL1(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!AuthorizeND) = True Then ChkND1(.Row - 1).Value = 1 Else ChkND1(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!IS_Holiday) = True Then ChkHol1(.Row - 1).Value = 1 Else ChkHol1(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!WithIT) = True Then ChkwIT1(.Row - 1).Value = 1 Else ChkwIT1(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!AuthorizeMA) = True Then ChkwMA1(.Row - 1).Value = 1 Else ChkwMA1(.Row - 1).Value = 0
'''''''''
'''''''''
'''''''''                            .Col = 5:
'''''''''                            .Text = Format(Null2String(rsCard!SHIFTINAM), "hh:mm AM/PM")
'''''''''                            .Col = 6:
'''''''''                            .Text = Format(Null2String(rsCard!SHIFTOUTAM), "hh:mm AM/PM")
'''''''''                            .Col = 7:
'''''''''                            .Text = Format(Null2String(rsCard!SHIFTINPM), "hh:mm AM/PM")
'''''''''                            .Col = 8:
'''''''''                            .Text = Format(Null2String(rsCard!SHIFTOUTPM), "hh:mm AM/PM")
'''''''''                        End With
'''''''''                    Else
'''''''''                        With Grid2
'''''''''                            .Row = r - 15
'''''''''                            .Col = 0: .CellForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(rsCard!DATETODAY, "mm/dd/yy")
''''''''''                            If Null2String(rsCard!INAM) = "" And Null2String(rsCard!OutAm) = "" Then
''''''''''                                .Col = 1: .CellBackColor = DefaultColor
''''''''''                                .Col = 2: .CellBackColor = DefaultColor
''''''''''                                ChkOL2(.Row - 1).Enabled = True
''''''''''                                ChkOL2(.Row - 1).BackColor = vbYellow
''''''''''                            End If
''''''''''                            If Null2String(rsCard!InPm) = "" And Null2String(rsCard!OutPM) = "" Then
''''''''''                                .Col = 3: .CellBackColor = DefaultColor
''''''''''                                .Col = 4: .CellBackColor = DefaultColor
''''''''''                                ChkOL2(.Row - 1).Enabled = True
''''''''''                                ChkOL2(.Row - 1).BackColor = vbYellow
''''''''''                            End If
'''''''''                            .Col = 1:                             '.CellForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(Null2String(rsCard!INAM), "hh:mm AM/PM")
''''''''''                            If Trim(.Text) <> "" Then
''''''''''                                If .Text < #7:00:00 AM# And .Text > #6:00:00 AM# Then
''''''''''                                    .CellForeColor = vbBlue: .CellBackColor = vbYellow
''''''''''                                    ChkOT2(.Row - 1).Enabled = True
''''''''''                                    ChkOT2(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwMA2(.Row - 1).Enabled = True
''''''''''                                    ChkwMA2(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text > #9:00:00 AM# And .Text < #11:00:00 AM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkUT2(.Row - 1).Enabled = True
''''''''''                                    ChkUT2(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwIT2(.Row - 1).Enabled = True
''''''''''                                    ChkwIT2(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text > #8:00:00 AM# And .Text < #9:00:00 AM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                End If
''''''''''                            End If
'''''''''                            .Col = 2:                             '.CellForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(Null2String(rsCard!OutAm), "hh:mm AM/PM")
''''''''''                            If Trim(.Text) <> "" Then
''''''''''                                If .Text < #11:00:00 AM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkUT2(.Row - 1).Enabled = True
''''''''''                                    ChkUT2(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwIT2(.Row - 1).Enabled = True
''''''''''                                    ChkwIT2(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                            End If
'''''''''                            .Col = 3:                             '.CellForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(Null2String(rsCard!InPm), "hh:mm AM/PM")
''''''''''                            If Trim(.Text) <> "" Then
''''''''''                                If .Text > #2:00:00 PM# And .Text < #4:00:00 PM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkUT2(.Row - 1).Enabled = True
''''''''''                                    ChkUT2(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwIT2(.Row - 1).Enabled = True
''''''''''                                    ChkwIT2(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text > #1:00:00 PM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkwIT2(.Row - 1).Enabled = True
''''''''''                                    ChkwIT2(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                            End If
'''''''''                            .Col = 4:                             '.CellForeColor = vbBlack: .CellBackColor = vbWhite
'''''''''                            .Text = Format(Null2String(rsCard!OutPM), "hh:mm AM/PM")
''''''''''                            If Trim(.Text) <> "" Then
''''''''''                                If .Text > #6:00:00 PM# Then
''''''''''                                    .CellForeColor = vbBlue: .CellBackColor = vbYellow
''''''''''                                    ChkOT2(.Row - 1).Enabled = True
''''''''''                                    ChkOT2(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwMA2(.Row - 1).Enabled = True
''''''''''                                    ChkwMA2(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text < #4:00:00 PM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkUT2(.Row - 1).Enabled = True
''''''''''                                    ChkUT2(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwMA2(.Row - 1).Enabled = True
''''''''''                                    ChkwMA2(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                                If .Text > #10:00:00 PM# Or .Text < #6:00:00 AM# Then
''''''''''                                    .CellForeColor = vbRed: .CellBackColor = vbYellow
''''''''''                                    ChkND2(.Row - 1).Enabled = True
''''''''''                                    ChkND2(.Row - 1).BackColor = vbYellow
''''''''''                                    ChkwMA2(.Row - 1).Enabled = True
''''''''''                                    ChkwMA2(.Row - 1).BackColor = vbYellow
''''''''''                                End If
''''''''''                            End If
''''''''''                            If Null2Bool(rsCard!AuthorizeOT) = True Then
''''''''''                               ChkOT2(.Row - 1).Value = 1
''''''''''                            Else
''''''''''                               ChkOT2(.Row - 1).Value = 0
''''''''''                            End If
''''''''''                            If Null2Bool(rsCard!AuthorizeUT) = True Then ChkUT2(.Row - 1).Value = 1 Else ChkUT2(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!AbsentWLeave) = True Then ChkOL2(.Row - 1).Value = 1 Else ChkOL2(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!AuthorizeND) = True Then ChkND2(.Row - 1).Value = 1 Else ChkND2(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!IS_Holiday) = True Then ChkHol2(.Row - 1).Value = 1 Else ChkHol2(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!WithIT) = True Then ChkwIT2(.Row - 1).Value = 1 Else ChkwIT2(.Row - 1).Value = 0
''''''''''                            If Null2Bool(rsCard!AuthorizeMA) = True Then ChkwMA2(.Row - 1).Value = 1 Else ChkwMA2(.Row - 1).Value = 0
'''''''''
'''''''''                            .Col = 5:
'''''''''                            .Text = Format(Null2String(rsCard!SHIFTINAM), "hh:mm AM/PM")
'''''''''                            .Col = 6:
'''''''''                            .Text = Format(Null2String(rsCard!SHIFTOUTAM), "hh:mm AM/PM")
'''''''''                            .Col = 7:
'''''''''                            .Text = Format(Null2String(rsCard!SHIFTINPM), "hh:mm AM/PM")
'''''''''                            .Col = 8:
'''''''''                            .Text = Format(Null2String(rsCard!SHIFTOUTPM), "hh:mm AM/PM")
'''''''''                        End With
'''''''''                    End If
'''''''''                    rsCard.MoveNext
'''''''''                Loop
'''''''''            End If
''''''''''===========================================



Sub InitGrid()
    Dim X, Y                                                          As Integer
    With Grid1
        .DisplayFocusRect = False
        .AllowUserResizing = False

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Column(0).Width = 0
        .Column(1).Width = 70
        .Column(2).Width = 30
        .Column(8).Width = 0
        .Column(9).Width = 0
        .Column(11).Width = 130
        
        For X = 3 To 6
            .Column(X).Width = 60
        Next X
        .Column(7).Width = 65
        .Column(10).Width = 65

        .Cell(0, 1).Text = "DATE"
        .Cell(0, 2).Text = "DAY"
        .Cell(0, 3).Text = "IN-AM"
        .Cell(0, 4).Text = "OUT-AM"
        .Cell(0, 5).Text = "IN-PM"
        .Cell(0, 6).Text = "OUT-PM"
        .Cell(0, 7).Text = "SHIFT-IN"
        .Cell(0, 8).Text = "SHIFT-OUT-AM"
        .Cell(0, 9).Text = "SHIFT-IN-PM"
        .Cell(0, 10).Text = "SHIFT-OUT"
        .Cell(0, 11).Text = "REMARKS"

        .Column(1).CellType = cellCalendar
        .Column(1).Locked = True
        .Column(2).Locked = True

        .Column(7).Locked = True
        .Column(10).Locked = True
        .Column(11).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 1, .Rows - 1, 10).ClearText
        '            For x = 0 To 10
        '                For y = 1 To 16
        '                    .Cell(y, x).Text = ""
        '                Next y
        '            Next x
    End With

    With Grid2
        .DisplayFocusRect = False: .AllowUserResizing = False

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Column(0).Width = 0
        .Column(1).Width = 70
        .Column(2).Width = 30
        .Column(8).Width = 0
        .Column(9).Width = 0
        .Column(11).Width = 130
        

        For X = 3 To 6
            .Column(X).Width = 60
        Next X
        .Column(7).Width = 65
        .Column(10).Width = 65

        .Cell(0, 1).Text = "DATE"
        .Cell(0, 2).Text = "DAY"
        .Cell(0, 3).Text = "IN-AM"
        .Cell(0, 4).Text = "OUT-AM"
        .Cell(0, 5).Text = "IN-PM"
        .Cell(0, 6).Text = "OUT-PM"
        .Cell(0, 7).Text = "SHIFT-IN"
        .Cell(0, 8).Text = "SHIFT-OUT-AM"
        .Cell(0, 9).Text = "SHIFT-IN-PM"
        .Cell(0, 10).Text = "SHIFT-OUT"
        .Cell(0, 11).Text = "REMARKS"

        .Column(1).CellType = cellCalendar
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(7).Locked = True
        .Column(10).Locked = True
        .Column(11).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 1, .Rows - 1, 10).ClearText
    End With
End Sub

Sub FillAll()

    Dim Criteria                                                      As String
    Dim k                                                             As Integer
    Dim r                                                             As Integer
    Screen.MousePointer = 11
    If cboMonth.Text = Date2Month(OneMonth(Date, -2)) Then
        TransactionDate = OneMonth(Date, -2)
    ElseIf cboMonth.Text = Date2Month(OneMonth(Date, -1)) Then
        TransactionDate = OneMonth(Date, -1)
    ElseIf cboMonth.Text = Date2Month(Date) Then
        TransactionDate = Date
    ElseIf cboMonth.Text = Date2Month(OneMonth(Date, 1)) Then
        TransactionDate = OneMonth(Date, 1)
    End If

    If cboEmpNumber.Text = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    InitGrid
    
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "Select * from HRMS_EmpInfo Where EmpNo  =  " & N2Str2Null(cboEmpNumber.Text), gconDMIS
    If rsEmpInfo.EOF Then
        For k = 1 To 150: Beep: Next k
        MsgBox "Employee Number Not Found", vbInformation, "Empty Record"
        cboEmpNumber.Text = ""
        On Error Resume Next
        Screen.MousePointer = 0
        cboEmpNumber.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    Else
        
        varEmpNo = rsEmpInfo!EMPNO
        TxtEmpName.Text = rsEmpInfo!lastname & ", " & rsEmpInfo!FIRSTNAME
        If Null2String(rsEmpInfo!PICFILNAME) <> "" Then
            On Error Resume Next
            If Len(Dir(HRMS_PICTURES_PATH & Null2String(rsEmpInfo!PICFILNAME))) > 0 Then
                LoadPic imgDispPic, HRMS_PICTURES_PATH & Null2String(rsEmpInfo!PICFILNAME)
            Else
                LoadPic imgDispPic, HRMS_PICTURES_PATH & "BlankFace.JPG"
            End If
        Else
            LoadPic imgDispPic, HRMS_PICTURES_PATH & "BlankFace.JPG"
        End If

        If Null2String(rsEmpInfo!ACTIVEINACTIVE) = "I" Then
            For k = 1 To 150: Beep: Next k
            MsgBox "Employee Not Active", 0, "Inactive Employee"
            On Error Resume Next
            cboEmpNumber.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        '=======================
        'TO DISPLAY BY CUT OFF
        Screen.MousePointer = 11
        Dim BegDayFirstCutOff, BegDaySecondCutOff                     As Integer
        Dim EndDayFirstCutOff, EndDaySecondCutOff                     As Integer
        Dim INCREMENTING_DAY                                          As Integer
        Dim I                                                         As Long
        'InitGrid
        Dim rsPayrollSetup                                            As ADODB.Recordset
        Set rsPayrollSetup = New ADODB.Recordset
        Set rsPayrollSetup = gconDMIS.Execute("Select * from HRMS_PayrollSetup")
        If Not rsPayrollSetup.EOF And Not rsPayrollSetup.BOF Then
            BegDayFirstCutOff = N2Str2Zero(rsPayrollSetup!FROMDATE1)
            BegDaySecondCutOff = N2Str2Zero(rsPayrollSetup!FROMDATE2)
            EndDayFirstCutOff = N2Str2Zero(rsPayrollSetup!TODATE1)
            EndDaySecondCutOff = N2Str2Zero(rsPayrollSetup!TODATE2)

            If BegDayFirstCutOff > EndDayFirstCutOff Then
                TransactionDate = DateSerial(cboYear, What_month(Trim(cboMonth)) - 1, BegDayFirstCutOff)
            Else
                TransactionDate = DateSerial(cboYear, What_month(Trim(cboMonth)), BegDayFirstCutOff)
            End If
            INCREMENTING_DAY = BegDayFirstCutOff: r = 0
            Grid1.AutoRedraw = True
            
            'MsgBox DateDiff("d", CDate(Null2String(Format(TransactionDate, "mm/dd/yyyy"))), Null2String(CDate(Format(DateSerial(cboYear, What_month(RTrim(LTrim(cboMOnth))), EndDayFirstCutOff), "mm/dd/yyyy"))))
            Grid1.Rows = DateDiff("d", CDate(Null2String(Format(TransactionDate, "mm/dd/yyyy"))), Null2String(CDate(Format(DateSerial(cboYear, What_month(RTrim(LTrim(cboMonth))), EndDayFirstCutOff), "mm/dd/yyyy")))) + 2
            Do While CDate(Format(TransactionDate, "mm/dd/yyyy")) <= CDate(Format(DateSerial(cboYear, What_month(RTrim(LTrim(cboMonth))), EndDayFirstCutOff), "mm/dd/yyyy"))
                INCREMENTING_DAY = INCREMENTING_DAY + 1
                r = r + 1
                Set rsCard = New ADODB.Recordset
                rsCard.Open "Select * from HRMS_Attend where Empno = " & N2Str2Null(varEmpNo) & " AND convert(varchar,DateToday,101) = '" & Format(TransactionDate, "mm/dd/yyyy") & "'", gconDMIS
                'PIYA
                Grid1.Cell(r, 2).Text = UCase(WeekdayName(Weekday(TransactionDate), True))
                Grid1.Cell(r, 1).Text = Format(TransactionDate, "mm/dd/yy")
                
                
                Grid1.Cell(r, 11).Text = IsHolyday(MONTH(TransactionDate), Day(TransactionDate))
                
                Grid1.Cell(r, 7).Text = ""
                Grid1.Cell(r, 8).Text = ""
                Grid1.Cell(r, 9).Text = ""
                Grid1.Cell(r, 10).Text = ""
                
                
                
                If Not rsCard.EOF And Not rsCard.BOF Then
                    With Grid1
                        .Cell(r, 3).Text = Format(Null2String(rsCard!INAM), "hh:mm AM/PM")
                        .Cell(r, 4).Text = Format(Null2String(rsCard!OUTAM), "hh:mm AM/PM")
                        .Cell(r, 5).Text = Format(Null2String(rsCard!INPM), "hh:mm AM/PM")
                        .Cell(r, 6).Text = Format(Null2String(rsCard!OUTPM), "hh:mm AM/PM")
                        .Cell(r, 7).Text = Format(Null2String(rsCard!SHIFTINAM), "hh:mm AM/PM")
                        .Cell(r, 8).Text = Format(Null2String(rsCard!SHIFTOUTAM), "hh:mm AM/PM")
                        .Cell(r, 9).Text = Format(Null2String(rsCard!SHIFTINPM), "hh:mm AM/PM")
                        .Cell(r, 10).Text = Format(Null2String(rsCard!SHIFTOUTPM), "hh:mm AM/PM")
                        '.Cell(r, 11).Text = IsHolyday(Month(rsCard!datetoday), Day(rsCard!datetoday))
                    End With
                End If
                With Grid1
                    I = 0
                    If LTrim(RTrim(.Cell(r, 11).Text)) <> "" Then
                        .Range(r, 0, r, 11).BackColor = &H69CBFA
                    ElseIf UCase(Grid1.Cell(r, 2).Text) = "SAT" Or UCase(Grid1.Cell(r, 2).Text) = "SUN" Then
                        .Range(r, 0, r, 11).BackColor = &HC0FFC0
                    Else
                        For I = 0 To 11
                            If r Mod 2 = 0 Then
                                .Cell(r, I).BackColor = .BackColor1
                            Else
                                .Cell(r, I).BackColor = .BackColor2
                            End If
                        Next
                    End If
                End With
                TransactionDate = CDate(TransactionDate + 1)
            Loop
            Grid1.AutoRedraw = True: Grid1.Refresh
'------------------------------------------------------------------------------------------------
            I = 0
            If BegDaySecondCutOff > EndDaySecondCutOff Then
                TransactionDate = DateSerial(cboYear, What_month(Trim(cboMonth)) - 1, BegDaySecondCutOff)
            Else
                TransactionDate = DateSerial(cboYear, What_month(Trim(cboMonth)), BegDaySecondCutOff)
            End If
            INCREMENTING_DAY = BegDaySecondCutOff: r = 0
            
            Grid2.Rows = DateDiff("d", Null2String(CDate(TransactionDate)), Null2String(CDate(DateSerial(cboYear, What_month(Trim(cboMonth)), EndDaySecondCutOff)))) + 2
            Do While CDate(TransactionDate) <= CDate(DateSerial(cboYear, What_month(Trim(cboMonth)), EndDaySecondCutOff))
                INCREMENTING_DAY = INCREMENTING_DAY + 1: r = r + 1
                Set rsCard = New ADODB.Recordset
                rsCard.Open "Select * from HRMS_Attend where Empno = " & N2Str2Null(varEmpNo) & " AND DateToday = '" & TransactionDate & "'", gconDMIS
                'PIYA
                Grid2.Cell(r, 2).Text = UCase(WeekdayName(Weekday(TransactionDate), True))
                Grid2.Cell(r, 1).Text = Format(TransactionDate, "mm/dd/yy")
                Grid2.Cell(r, 11).Text = IsHolyday(MONTH(TransactionDate), Day(TransactionDate))
                
                Grid2.Cell(r, 7).Text = ""
                Grid2.Cell(r, 8).Text = ""
                Grid2.Cell(r, 9).Text = ""
                Grid2.Cell(r, 10).Text = ""
                
                
                
                If Not rsCard.EOF And Not rsCard.BOF Then
                    With Grid2
                        .Cell(r, 3).Text = Format(Null2String(rsCard!INAM), "hh:mm AM/PM")
                        .Cell(r, 4).Text = Format(Null2String(rsCard!OUTAM), "hh:mm AM/PM")
                        .Cell(r, 5).Text = Format(Null2String(rsCard!INPM), "hh:mm AM/PM")
                        .Cell(r, 6).Text = Format(Null2String(rsCard!OUTPM), "hh:mm AM/PM")
                        .Cell(r, 7).Text = Format(Null2String(rsCard!SHIFTINAM), "hh:mm AM/PM")
                        .Cell(r, 8).Text = Format(Null2String(rsCard!SHIFTOUTAM), "hh:mm AM/PM")
                        .Cell(r, 9).Text = Format(Null2String(rsCard!SHIFTINPM), "hh:mm AM/PM")
                        .Cell(r, 10).Text = Format(Null2String(rsCard!SHIFTOUTPM), "hh:mm AM/PM")
                        '.Cell(r, 11).Text = IsHolyday(Month(rsCard!datetoday), Day(rsCard!datetoday))
                    End With
                End If
                I = 0
                With Grid2                                   'piya
                    If LTrim(RTrim(.Cell(r, 11).Text)) <> "" Then
                        .Range(r, 0, r, 11).BackColor = &H69CBFA
                    ElseIf UCase(Grid2.Cell(r, 2).Text) = "SAT" Or UCase(Grid2.Cell(r, 2).Text) = "SUN" Then
                        .Range(r, 0, r, 11).BackColor = &HC0FFC0
                    Else
                        For I = 1 To 11
                            If r Mod 2 = 0 Then
                                .Cell(r, I).BackColor = .BackColor1
                            Else
                                .Cell(r, I).BackColor = .BackColor2
                            End If
                        Next
                    End If
                End With
                TransactionDate = CDate(TransactionDate + 1)
            Loop
        End If
        'END DISPLAY BY CUT OFF
    End If
    Grid1.Refresh
    CmdLineFill1.Enabled = True
    CmdLineFill2.Enabled = True
    CmdFillAll1.Enabled = True
    CmdFillAll2.Enabled = True
    Screen.MousePointer = 0
End Sub

Sub FIllCbo()
    'updated by:    IEBV    06221145
    'description: load only the non contractual employee
    'Combo_Loadval cboEmpNumber, gconDMIS.Execute("Select EMPNO FROM HRMS_EMPINFO ORDER BY EMPNO DESC")
    Combo_Loadval cboEmpNumber, gconDMIS.Execute("Select EMPNO FROM HRMS_EMPINFO where emplevel <> 'C' ORDER BY EMPNO DESC")
    Dim X                                                             As Integer
    Dim Thedate                                                       As Date
    Thedate = OneMonth(Date, -3)
    For X = 1 To 12
        Thedate = OneMonth(Thedate, 1)
        cboMonth.AddItem Date2Month(CDate(Thedate))
        'cboyear.AddItem YEAR(CDate(Thedate))
    Next
End Sub



'End If
'End Sub

Private Sub cboEmpNumber_Change()
    'cboEmpNumber = Repleys(cboEmpNumber)                     'Format(cboEmpNumber, "000000")
    'FillAll
End Sub

Private Sub cboEmpNumber_Click()
    'cboEmpNumber_Change
    'FillAll
End Sub

Private Sub cboMonth_Change()
    'InitGrid
    'FillAll
End Sub

Private Sub cboYEAR_Change()
    InitGrid
    FillAll
End Sub

Private Sub cboyear_Click()
    InitGrid
    FillAll
End Sub

Private Sub cmdAdd_1stCutoff_Click()
  
    If Function_Access(LOGID, "Acess_Add", "EMPLOYEE TIME CARD") = False Then Exit Sub

   
      If MsgBox(" First Cut-off will be filled by the default time (08:00 AM - 05:00 PM). Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
                Call fillfirstcutoffgrid
      Else
            Exit Sub
      End If

End Sub

Private Sub cmdAdd_2ndCutoff_Click()
    
    If Function_Access(LOGID, "Acess_Add", "EMPLOYEE TIME CARD") = False Then Exit Sub
 
        If MsgBox(" 2nd Cut-off will be filled by the default time (08:00 AM - 05:00 PM). Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
                Call fillsecondcutoffgrid
        Else
            Exit Sub
        End If
 

 
End Sub

Private Sub cmdCancel_Click()
    '    picNOTES.Visible = False
    '    picNOTES2.Visible = False
    '    picTaas.Enabled = True
    '    picBaba.Enabled = True
    'On Error Resume Next
    'cboEmpNumber.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim k                                                             As Integer
    Dim X                                                             As Integer
    Dim rowDate                                                       As String
    'If Val(cboEmpNumber) = 0 Then
    '    MsgBox ("Sorry, There's Nothing to Save")
    'End If
    Dim IS_Authorize_OT                                               As Byte
    Dim IS_Authorize_UT                                               As Byte
    Dim ABSENT_WLEAVE                                                 As Byte
    Dim IS_Authorize_ND                                               As Byte
    Dim IS_Holiday                                                    As Byte
    Dim With_IT                                                       As Byte
    Dim IS_Authorize_MA                                               As Byte
    Dim CONT_SAVING_NOTES                                             As Boolean

    
    frmMain.Enabled = False
    FrmProgBar.Visible = True
    Me.Refresh
    Screen.MousePointer = 11
    
    DoEvents
    LblProgBar1.Caption = "Checking Invalid Entries, Please Wait..."
    DoEvents
    LblProgBar1.Caption = "Updating Records, Please Wait..."
    Dim TM(6)

    'UPDATED BY FML (FML - 05252007) -> UPDATED PATS FOR ACTUAL HRS, ENTERED HRS, ACTUAL DAYS, ENTERED DAYS AND OVERTIME, LATE AND OTHER COMPUTATIONS
    '======================================================================================================
    Dim Return_Minute
    Dim Return_Hour
    Dim valrtn

    Dim Date_And_Time_Punched_In_Am
    Dim Date_And_Time_Punched_Out_Am
    Dim Date_And_Time_Punched_In_Pm
    Dim Date_And_Time_Punched_Out_Pm

    Dim Number_Of_Hour_Worked_Am
    Dim Number_Of_Hour_Worked_Pm

    Dim Minutecalculated_Am
    Dim Minutecalculated_Pm

    Dim Return_Hour_Am
    Dim Return_Hour_Pm

    Dim Return_Minute_Am
    Dim Return_Minute_Pm

    Dim valrtn_Am
    Dim valrtn_Pm
    Dim TotalHourWorked_Am
    Dim TotalHourWorked_Pm

    Dim RegularHour_Am
    Dim RegularHour_Pm

    Dim Overtime_Hour_Worked_Am
    Dim Overtime_Hour_Worked_Pm

    Dim OT_CODE                                                       As String

    Dim vActualDays, vEnteredDays                                     As Double
    Set rsAttend = New ADODB.Recordset
    'Dim VINAM, VOUTAM, vINPM, VOUTPM
    
    Dim VINAM                                                         As String
    Dim VOUTAM                                                        As String
    Dim VINPM                                                         As String
    Dim VOUTPM                                                        As String
    Dim VDATES                                                        As String
    Dim vcombamay                                                     As String
    Dim vcombibanggi                                                  As String
    Dim v2dates                                                       As String
    Dim v2combamay                                                    As String
    Dim v2combibanggi                                                 As String

    
    Dim I                                                             As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim TMP_DATE1 As String
    Dim TMP_DATE2 As String
    Dim TMP_DATE3 As String
    Dim TMP_DATE4 As String
    Dim rsKUTO                                                      As New ADODB.Recordset
    
    ProgressBar1.Max = 100
    For I = 1 To Grid1.Rows - 1
        DoEvents
        'UPDATE BY   : MJP 12102008
        'DESCRIPTION : TO RETAIN THE ORIGINAL DATE
            Set RSTMP = New ADODB.Recordset
            Set RSTMP = gconDMIS.Execute("SELECT INAM,OUTAM, INPM, OUTPM FROM HRMS_ATTEND WHERE " & _
                " EMPNO =  " & N2Str2Null(varEmpNo) & _
                " And convert(varchar, DateToday ,101) = '" & Format((Grid1.Cell(I, 1).Text), "mm/dd/yyyy") & "'")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                If Not Null2String(RSTMP!INAM) = "" Then TMP_DATE1 = DateValue(Null2String(RSTMP!INAM))
                If Not Null2String(RSTMP!OUTAM) = "" Then TMP_DATE2 = DateValue(Null2String(RSTMP!OUTAM))
                If Not Null2String(RSTMP!INPM) = "" Then TMP_DATE3 = DateValue(Null2String(RSTMP!INPM))
                If Not Null2String(RSTMP!OUTPM) = "" Then TMP_DATE4 = DateValue(Null2String(RSTMP!OUTPM))
            End If
            Set RSTMP = Nothing
        'UPDATE BY   : MJP 12102008
        
        VDATES = N2Str2Null(Grid1.Cell(I, 1).Text)
        VINAM = N2Str2Null(Grid1.Cell(I, 3).Text)
        VOUTAM = N2Str2Null(Grid1.Cell(I, 4).Text)
        VINPM = N2Str2Null(Grid1.Cell(I, 5).Text)
        VOUTPM = N2Str2Null(Grid1.Cell(I, 6).Text)

        'UPDATE BY   : JBF 09/16/2010
        'DESCRIPTION : TO UPDATE SAVE for FIll button
        ' ************************************************
            If VINAM = "NULL" Then
                vcombamay = "NULL"
            Else
                vcombamay = Format(Grid1.Cell(I, 1).Text, "mm/dd/yyyy") & " " & Format(Grid1.Cell(I, 3).Text, "hh:mm:ss")
            End If
              
            
            If VOUTPM = "NULL" Then
                vcombibanggi = "NULL"
            Else
                vcombibanggi = Format(Grid1.Cell(I, 1).Text, "mm/dd/yyyy") & " " & Format(Grid1.Cell(I, 6).Text, "hh:mm:ss")
            End If

         ' ************************************************
        
        
        
        'UPDATE BY   : MJP 12102008
        'DESCRIPTION : TO RETAIN THE ORIGINAL DATE
            If Not Null2String(VINAM) = "" Then VINAM = N2Str2Null(TMP_DATE1 & " " & VINAM)
            If Not Null2String(VOUTAM) = "" Then VOUTAM = N2Str2Null(TMP_DATE2 & " " & VOUTAM)
            If Not Null2String(VINPM) = "" Then VINPM = N2Str2Null(TMP_DATE3 & " " & VINPM)
            If Not Null2String(VOUTPM) = "" Then VOUTPM = N2Str2Null(TMP_DATE4 & " " & VOUTPM)
        'UPDATE BY   : MJP 12102008
        
        DoEvents
        If IsDate(Grid1.Cell(I, 1).Text) = True Then
            'UPDATE BY   : MJP 01052010
            'DESCRIPTION : TCN00029
                Set rsKUTO = gconDMIS.Execute("SELECT EMPNO FROM HRMS_ATTEND WHERE " & _
                    " EMPNO = " & N2Str2Null(varEmpNo) & _
                    " And convert(varchar, DateToday ,101) = '" & Format((Grid1.Cell(I, 1).Text), "mm/dd/yyyy") & "'")
                If (rsKUTO.BOF And rsKUTO.EOF) Then
                    'If "#" & Format(Date, "mm/dd/yyyy") & "#" <= "#" & Format((Grid1.Cell(I, 1).Text), "mm/dd/yyyy") & "#" Then
                    If DateDiff("d", Format(Date, "mm/dd/yyyy"), Format((Grid1.Cell(I, 1).Text), "mm/dd/yyyy")) < 0 Then
                        
'                        gconDMIS.Execute ("INSERT INTO HRMS_ATTEND " & _
'                            " (DATETODAY, EMPNO, INAM, OUTAM, INPM, OUTPM) " & _
'                            " VALUES('" & Format((Grid1.Cell(I, 1).Text), "mm/dd/yyyy") & _
'                            "', " & N2Str2Null(varEmpNo) & _
'                            ", " & VINAM & _
'                            ", " & VOUTAM & _
'                            ", " & VINPM & _
'                            ", " & VOUTPM & ")")
                    
                         'UPDATE BY   : JBF 09/16/2010
                    
                         gconDMIS.Execute ("INSERT INTO HRMS_ATTEND " & _
                            " (DATETODAY, EMPNO, INAM, OUTAM, INPM, OUTPM,SHIFTINAM,SHIFTOUTPM) " & _
                            " VALUES('" & Format((Grid1.Cell(I, 1).Text), "mm/dd/yyyy") & _
                            "', " & N2Str2Null(varEmpNo) & _
                            ", " & N2Str2Null(vcombamay) & _
                            ", " & VOUTAM & _
                            ", " & VINPM & _
                            ", " & N2Str2Null(vcombibanggi) & _
                            ", " & N2Str2Null(vcombamay) & _
                            ", " & N2Str2Null(vcombibanggi) & ")")
                    
                    End If
                Else
            'UPDATE BY   : MJP 01052010
                gconDMIS.Execute "update HRMS_Attend " & _
                    " set INAM = " & VINAM & _
                    ", OUTAM = " & VOUTAM & _
                    " ,INPM = " & VINPM & _
                    ", OUTPM = " & VOUTPM & _
                    " Where EmpNo = " & N2Str2Null(varEmpNo) & _
                    " And convert(varchar, DateToday ,101) = '" & Format((Grid1.Cell(I, 1).Text), "mm/dd/yyyy") & "'"
            End If
            Set rsKUTO = Nothing
        End If
    Next
    DoEvents
    ProgressBar1.Value = 50
    
    For I = 1 To Grid2.Rows - 1
        DoEvents
        'UPDATE BY   : MJP 12102008
        'DESCRIPTION : TO RETAIN THE ORIGINAL DATE
            Set RSTMP = New ADODB.Recordset
         
    
            Set RSTMP = gconDMIS.Execute("SELECT INAM,OUTAM, INPM, OUTPM FROM HRMS_ATTEND WHERE " & _
                " EMPNO =  " & N2Str2Null(varEmpNo) & _
                " And convert(varchar, DateToday ,101) = '" & (Grid2.Cell(I, 1).Text) & "'")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                If Not Null2String(RSTMP!INAM) = "" Then TMP_DATE1 = DateValue(Null2String(RSTMP!INAM))
                If Not Null2String(RSTMP!OUTAM) = "" Then TMP_DATE2 = DateValue(Null2String(RSTMP!OUTAM))
                If Not Null2String(RSTMP!INPM) = "" Then TMP_DATE3 = DateValue(Null2String(RSTMP!INPM))
                If Not Null2String(RSTMP!OUTPM) = "" Then TMP_DATE4 = DateValue(Null2String(RSTMP!OUTPM))
            End If
            Set RSTMP = Nothing
        'UPDATE BY   : MJP 12102008
        
        
        v2dates = N2Str2Null(Grid2.Cell(I, 1).Text)
        VINAM = N2Str2Null(Grid2.Cell(I, 3).Text)
        VOUTAM = N2Str2Null(Grid2.Cell(I, 4).Text)
        VINPM = N2Str2Null(Grid2.Cell(I, 5).Text)
        VOUTPM = N2Str2Null(Grid2.Cell(I, 6).Text)

        
        'UPDATE BY   : JBF 09/16/2010
        'DESCRIPTION : TO UPDATE SAVE for FIll button
        ' ************************************************
        
        If VINAM = "NULL" Then
            v2combamay = "NULL"
        Else
            v2combamay = Format(Grid2.Cell(I, 1).Text, "mm/dd/yyyy") & " " & Format(Grid2.Cell(I, 3).Text, "hh:mm:ss")
        End If
          
        
        If VOUTPM = "NULL" Then
            v2combibanggi = "NULL"
        Else
            v2combibanggi = Format(Grid2.Cell(I, 1).Text, "mm/dd/yyyy") & " " & Format(Grid2.Cell(I, 6).Text, "hh:mm:ss")
        End If

        ' ************************************************

        'UPDATE BY   : MJP 12102008
        'DESCRIPTION : TO RETAIN THE ORIGINAL DATE
            If Not Null2String(VINAM) = "" Then VINAM = N2Str2Null(TMP_DATE1 & " " & VINAM)
            If Not Null2String(VOUTAM) = "" Then VOUTAM = N2Str2Null(TMP_DATE2 & " " & VOUTAM)
            If Not Null2String(VINPM) = "" Then VINPM = N2Str2Null(TMP_DATE3 & " " & VINPM)
            If Not Null2String(VOUTPM) = "" Then VOUTPM = N2Str2Null(TMP_DATE4 & " " & VOUTPM)
        'UPDATE BY   : MJP 12102008
        
        DoEvents
        If IsDate(Grid2.Cell(I, 1).Text) = True Then
            'UPDATE BY   : MJP 01052010
            'DESCRIPTION : TCN00029
                Set rsKUTO = gconDMIS.Execute("SELECT EMPNO FROM HRMS_ATTEND WHERE " & _
                    " EMPNO = " & N2Str2Null(varEmpNo) & _
                    " And convert(varchar, DateToday ,101) = '" & Format((Grid2.Cell(I, 1).Text), "mm/dd/yyyy") & "'")
                If (rsKUTO.BOF And rsKUTO.EOF) Then
                    'If "#" & Format(Date, "mm/dd/yyyy") & "#" <= "#" & Format((Grid2.Cell(I, 1).Text), "mm/dd/yyyy") & "#" Then
                    If DateDiff("d", Format(Date, "mm/dd/yyyy"), Format((Grid2.Cell(I, 1).Text), "mm/dd/yyyy")) < 0 Then
                        
'                        gconDMIS.Execute ("INSERT INTO HRMS_ATTEND " & _
'                            " (DATETODAY, EMPNO, INAM, OUTAM, INPM, OUTPM) " & _
'                            " VALUES('" & Format((Grid2.Cell(I, 1).Text), "mm/dd/yyyy") & _
'                            "', " & N2Str2Null(varEmpNo) & _
'                            ", " & VINAM & _
'                            ", " & VOUTAM & _
'                            ", " & VINPM & _
'                            ", " & VOUTPM & ")")
                    
                        
                        gconDMIS.Execute ("INSERT INTO HRMS_ATTEND " & _
                            " (DATETODAY, EMPNO, INAM, OUTAM, INPM, OUTPM,SHIFTINAM,SHIFTOUTPM) " & _
                            " VALUES('" & Format((Grid2.Cell(I, 1).Text), "mm/dd/yyyy") & _
                            "', " & N2Str2Null(varEmpNo) & _
                            ", " & N2Str2Null(v2combamay) & _
                            ", " & VOUTAM & _
                            ", " & VINPM & _
                            ", " & N2Str2Null(v2combibanggi) & _
                            ", " & N2Str2Null(v2combamay) & _
                            ", " & N2Str2Null(v2combibanggi) & ")")
                    
                    End If
                Else
            'UPDATE BY   : MJP 01052010
                gconDMIS.Execute "update HRMS_Attend " & _
                    " set INAM = " & VINAM & _
                    ", OUTAM = " & VOUTAM & _
                    " ,INPM = " & VINPM & _
                    ", OUTPM = " & VOUTPM & _
                    " Where EmpNo = " & N2Str2Null(varEmpNo) & _
                    " And convert(varchar, DateToday ,101) = '" & Format((Grid2.Cell(I, 1).Text), "mm/dd/yyyy") & "'"
            End If
            Set rsKUTO = Nothing
        End If
    Next
    ProgressBar1.Value = 100
    DoEvents
    
    
    
    'COMMENT BY  : MJP 120908
    'DESCRIPTION : TO MAKE THE SAVING MORE FASTER
        'rsAttend.Open "Select * from HRMS_Attend Where EmpNo = " & N2Str2Null(varEmpNo), gconDMIS
    'COMMENT BY  : MJP 120908
    
    'UPDATE BY   : MJP 120908
    'DESCRIPTION : SELECT THE DATE RANGE ONLY FOR THIS CUT OFF
        rsAttend.Open "Select * from HRMS_Attend Where " & _
            " EmpNo = " & N2Str2Null(varEmpNo) & _
            " AND DATETODAY BETWEEN '" & Format((Grid1.Cell(1, 1).Text), "mm/dd/yyyy") & _
            "' and '" & Format((Grid2.Cell(1, 1).Text), "mm/dd/yyyy") & "'", gconDMIS
    'UPDATE BY   : MJP 120908
    If Not rsAttend.EOF And Not rsAttend.BOF Then
        rsAttend.MoveFirst
        DoEvents
        ProgressBar1.Max = rsAttend.RecordCount
        ProgressBar1.Value = 0
        LblProgBar1.Caption = "Updating Actual Hours..."
        Do While Not rsAttend.EOF
            If N2Str2Zero(rsAttend!ActualHrsAm) + N2Str2Zero(rsAttend!ActualHrsPm) = 8 Then
                vActualDays = 1
            Else
                If N2Str2Zero(rsAttend!ActualHrsAm) + N2Str2Zero(rsAttend!ActualHrsPm) = 4 Then
                    vActualDays = 0.5
                Else
                    vActualDays = 0
                End If
            End If
            If N2Str2Zero(rsAttend!TotalHrsAm) + N2Str2Zero(rsAttend!TotalHrsPm) > 6 Then
                vEnteredDays = 1
            Else
                If N2Str2Zero(rsAttend!TotalHrsAm) + N2Str2Zero(rsAttend!TotalHrsPm) >= 4 Then
                    vEnteredDays = 0.5
                Else
                    vEnteredDays = 0
                End If
            End If

            gconDMIS.Execute "update HRMS_Attend " & _
                " Set ActualDays = " & vActualDays & _
                ", EnteredDays = " & vEnteredDays & _
                " Where EmpNo = " & N2Str2Null(varEmpNo) & _
                " And convert(varchar, DateToday ,101) = '" & Format((rsAttend!datetoday), "mm/dd/yyyy") & "'"
                
            DoEvents
            LblProgBar2.Caption = Format((ProgressBar1.Value / ProgressBar1.Max) * 100, "##0") & " %"
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
            
            rsAttend.MoveNext
        Loop
    End If
    
    frmMain.Enabled = True
    Screen.MousePointer = 0
    
    Call ShowSuccessFullyUpdated
    Call FillAll
    FrmProgBar.Visible = False
    DisableAllObject True
    picNOTES.Visible = False
    picNOTES2.Visible = False
    picTaas.Enabled = True
    picBaba.Enabled = True
End Sub

Private Sub Command2_Click()
    GENFROM = DateSerial(cboYear.Text, What_month(Trim(cboMonth.Text)), PAYROLLCODE_FROM1)
    GENTO = DateSerial(cboYear.Text, What_month(Trim(cboMonth.Text)), PAYROLLCODE_TO2)
    If PAYROLLCODE_TO2 < PAYROLLCODE_FROM1 Then
        GENFROM = DateSerial(cboYear.Text, What_month(Trim(cboMonth.Text)) - 1, PAYROLLCODE_FROM1)
        If What_month(Trim(cboMonth.Text)) = 1 Then
            GENFROM = DateSerial(NumericVal(cboYear.Text) - 1, 12, PAYROLLCODE_FROM1)
        End If
    End If
    rptDepartment.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptDepartment.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptDepartment.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptDepartment.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
    PrintSQLReport rptDepartment, HRMS_REPORT_PATH & "attendance.rpt", "(date({hrms_attend.datetoday}) >=  CDate('" & MONTH(GENFROM) & "/" & Day(GENFROM) & "/" & YEAR(GENFROM) & "')) AND {HRMS_ATTEND.empno} = '" & cboEmpNumber & "' and (date({hrms_attend.datetoday}) <=  CDate('" & MONTH(GENTO) & "/" & Day(GENTO) & "/" & YEAR(GENTO) & "'))", DMIS_REPORT_Connection, 1
    GENFROM = ""
    GENTO = ""
End Sub

Private Sub Command3_Click()
    Call FillAll
 
 
 'Grid1.Column(3).Locked = True
 Grid1.Column(4).Locked = True
 Grid1.Column(5).Locked = True
 
 'Grid2.Column(3).Locked = True
 Grid2.Column(4).Locked = True
 Grid2.Column(5).Locked = True
 
 cmdAdd_1stCutoff.Enabled = True
 cmdAdd_2ndCutoff.Enabled = True
 
 
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    
    FrmProgBar.Visible = False
    FIllCbo
'    Call fillcbomonth(cboMOnth)
'    Call FillcboYear(cboYear)
    InitGrid
    TxtMonth.Text = Date2Month(Date)
    cboMonth.Text = Date2Month(Date)
    
    'FillcboYear cboyear
    fillcombo_up cboYear
    cboYear.Text = YEAR(Date)
    Me.Top = Me.Top - 200
    pic1To15.Left = 0
    pic16to31.Left = 0
    'Call FillTheTypeOfOTKinds
End Sub

Private Sub Grid1_DblClick()

If Grid1.ActiveCell.Col = 6 And Grid1.Cell(Grid1.ActiveCell.Row, 6).Text = "" Then
        menuovertime.Enabled = False
        Call PopupMenu(menuoption)
ElseIf Grid1.ActiveCell.Col = 6 And Grid1.Cell(Grid1.ActiveCell.Row, 6).Text <> "" Then
    
        menuovertime.Enabled = True
        Call PopupMenu(menuoption)
Else
    'do thing
End If

End Sub

Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Then
        If Grid1.ActiveCell.Text = "" Then Exit Sub
        If IsDate(Grid1.ActiveCell.Text) = False Then
            Cancel = True
        Else
            Grid1.ActiveCell.Text = Format(Grid1.ActiveCell.Text, "HH:MM AM/PM")
        End If
    End If
End Sub

Private Sub Grid2_DblClick()
If Grid2.ActiveCell.Col = 6 And Grid2.Cell(Grid2.ActiveCell.Row, 6).Text = "" Then

    mainot.Enabled = False
    Call PopupMenu(mainoption)
ElseIf Grid2.ActiveCell.Col = 6 And Grid2.Cell(Grid2.ActiveCell.Row, 6).Text <> "" Then
    
    mainot.Enabled = True
    Call PopupMenu(mainoption)
Else
    'do thing
End If
End Sub

Private Sub Grid2_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Then
        If Grid2.ActiveCell.Text = "" Then Exit Sub
        If IsDate(Grid2.ActiveCell.Text) = False Then
            Cancel = True
        Else
            Grid2.ActiveCell.Text = Format(Grid2.ActiveCell.Text, "HH:MM AM/PM")
        End If
    End If
End Sub

Sub DisableAllObject(COND As Boolean)
    Grid1.Enabled = COND
    Grid2.Enabled = COND
    picTaas.Enabled = COND
    picBaba.Enabled = COND
    cboMonth.Enabled = COND
    cboYear.Enabled = COND
    cboEmpNumber.Enabled = COND
    TxtEmpName.Enabled = COND
    Command1.Enabled = COND
    CmdCancel.Enabled = COND
    CmdExit.Enabled = COND
    Text1.Enabled = COND
    CmdLineFill1.Enabled = COND
    CmdFillAll1.Enabled = COND
    Text2.Enabled = COND
    CmdLineFill2.Enabled = COND
    CmdFillAll2.Enabled = COND
End Sub

Private Sub HScroll1_Change()
    pic1To15.Left = 0 - HScroll1.Value
End Sub

Private Sub HScroll2_Change()
    pic16to31.Left = 0 - HScroll2.Value
End Sub

Private Sub cboEmpNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboMonth_Click()
    InitGrid
    FillAll
End Sub

Private Sub lsvSEARCH_DblClick()
    If lsvSEARCH.ListItems.count = 0 Then Exit Sub
    
    Dim Index As Integer
    
    Index = lsvSEARCH.SelectedItem.Index
    cboEmpNumber.Text = lsvSEARCH.ListItems(Index).ListSubItems(1)
    lblEmpno.Caption = lsvSEARCH.ListItems(Index).ListSubItems(1)
    lblName.Caption = lsvSEARCH.ListItems(Index).Text
    lblShift.Caption = GetEmployeeShift(lsvSEARCH.ListItems(Index).ListSubItems(1), "from1") & " - " & GetEmployeeShift(lsvSEARCH.ListItems(Index).ListSubItems(1), "to1")
    Call FillAll
End Sub

Function GetEmployeeShift(XEMPNO As String, xFIELD As String) As String
    Dim RSTMP   As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT SHIFT, " & _
        " ( " & _
        "   SELECT CONVERT(VARCHAR," & xFIELD & ",8) FROM HRMS_Time_Shift_Code WHERE SHIFTCODE = H.SHIFT " & _
        " ) AS RESULT " & _
        " FROM HRMS_EMPINFO H WHERE EMPNO = " & N2Str2Null(XEMPNO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        On Error GoTo Error_Time
        GetEmployeeShift = TimeValue(Null2String(RSTMP!RESULT))
    End If
    Set RSTMP = Nothing
    
    Exit Function
Error_Time:
    GetEmployeeShift = "Not Set"
End Function

Private Sub lsvSEARCH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvSEARCH_DblClick
End Sub

Private Sub mainleave_Click()
Screen.MousePointer = 11
        Set frms = New frmHRMS_RequestForLeave_short
        Call frms.SelectSQl("Select * from hrms_requestleave_ot where empno = " & N2Str2Null(cboEmpNumber) & "", cboEmpNumber.Text)
        frms.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub mainot_Click()

 Screen.MousePointer = 11
    
       xbuwan = Format(Grid2.Cell(Grid2.ActiveCell.Row, 1).Text, "mm/dd/yyyy")
       xin = Format(Grid2.Cell(Grid2.ActiveCell.Row, 3).Text, "hh:mm AM/PM")
       xout = Format(Grid2.Cell(Grid2.ActiveCell.Row, 6).Text, "hh:mm AM/PM")
        
       txtgrid.Text = xbuwan
       txtin_am.Text = xin
       txtout_pm.Text = xout
       
       'cboMonth.Text = Date2Month(Date)
    
        Set FRM = New frmHRMS_OT
        Call FRM.SelectSQl("Select * from HRMS_OVERTIME where empno = " & N2Str2Null(cboEmpNumber) & "", cboEmpNumber.Text, txtgrid.Text, cboMonth.Text, txtin_am, txtout_pm.Text)
        EMP_TYPE = "EMPLOYEE"
        FRM.Show 1
 
 Screen.MousePointer = 0
End Sub

Private Sub menuleave_Click()
 Screen.MousePointer = 11
        Set frms = New frmHRMS_RequestForLeave_short
        Call frms.SelectSQl("Select * from hrms_requestleave_ot where empno = " & N2Str2Null(cboEmpNumber) & "", cboEmpNumber.Text)
        frms.Show 1
Screen.MousePointer = 0
End Sub

Private Sub menuovertime_Click()

 Screen.MousePointer = 11
    
         
       
       xbuwan = Format(Grid1.Cell(Grid1.ActiveCell.Row, 1).Text, "mm/dd/yyyy")
       xin = Format(Grid1.Cell(Grid1.ActiveCell.Row, 3).Text, "hh:mm AM/PM")
       xout = Format(Grid1.Cell(Grid1.ActiveCell.Row, 6).Text, "hh:mm AM/PM")
        
       txtgrid.Text = xbuwan
       txtin_am.Text = xin
       txtout_pm.Text = xout
       
    
       Set FRM = New frmHRMS_OT
       Call FRM.SelectSQl("Select * from HRMS_OVERTIME where empno = " & N2Str2Null(cboEmpNumber) & "", cboEmpNumber.Text, txtgrid.Text, cboMonth.Text, txtin_am, txtout_pm.Text)
       EMP_TYPE = "EMPLOYEE"
       FRM.Show 1
 
 Screen.MousePointer = 0


End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        picEMPNO.Visible = True
        picNAME.Visible = False
        
        cboEmpNumber.SetFocus
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        picEMPNO.Visible = False
        picNAME.Visible = True
        
        Text3.SetFocus
    End If
End Sub

Private Sub Text3_Change()
    If Text3.Text = "" Then
        Call FillGrid
    Else
        Call FillSearchGrid(Text3)
    End If
End Sub

Sub FillGrid()
    Dim RSTMP  As New ADODB.Recordset
    Dim ITEM As ListItem
    Set RSTMP = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME AS NAME,EMPNO FROM HRMS_EMPINFO ORDER BY LASTNAME + ', ' + FIRstNAME")
    lsvSEARCH.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvSEARCH.ListItems.Add(, , Null2String(RSTMP!NAME))
            ITEM.SubItems(1) = RSTMP!EMPNO
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSTMP  As New ADODB.Recordset
    Dim ITEM As ListItem
    Dim KKK As String
    
    KKK = Replace(XXX, "'", "")
    Set RSTMP = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME AS NAME,EMPNO FROM HRMS_EMPINFO WHERE LASTNAME + ', ' + FIRstNAME LIKE '%" & KKK & "%' ORDER BY LASTNAME + ', '  + FIRstNAME")
    lsvSEARCH.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvSEARCH.ListItems.Add(, , Null2String(RSTMP!NAME))
            ITEM.SubItems(1) = RSTMP!EMPNO
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lsvSEARCH.SetFocus
End Sub

Sub initgrid1()

Dim X, Y                                                          As Integer
    With Grid1
        .DisplayFocusRect = False
        .AllowUserResizing = False

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Column(0).Width = 0
        .Column(1).Width = 70
        .Column(2).Width = 30
        .Column(8).Width = 0
        .Column(9).Width = 0
        .Column(11).Width = 130

        For X = 3 To 6
            .Column(X).Width = 60
        Next X
        .Column(7).Width = 65
        .Column(10).Width = 65

        .Cell(0, 1).Text = "DATE"
        .Cell(0, 2).Text = "DAY"
        .Cell(0, 3).Text = "IN-AM"
        .Cell(0, 4).Text = "OUT-AM"
        .Cell(0, 5).Text = "IN-PM"
        .Cell(0, 6).Text = "OUT-PM"
        .Cell(0, 7).Text = "SHIFT-IN"
        .Cell(0, 8).Text = "SHIFT-OUT-AM"
        .Cell(0, 9).Text = "SHIFT-IN-PM"
        .Cell(0, 10).Text = "SHIFT-OUT"
        .Cell(0, 11).Text = "REMARKS"

        .Column(1).CellType = cellCalendar
        .Column(1).Locked = True
        .Column(2).Locked = True

        .Column(7).Locked = True
        .Column(10).Locked = True
        .Column(11).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 1, .Rows - 1, 10).ClearText
    End With

End Sub

Sub initgrid2()

    With Grid2
        .DisplayFocusRect = False: .AllowUserResizing = False

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Column(0).Width = 0
        .Column(1).Width = 70
        .Column(2).Width = 30
        .Column(8).Width = 0
        .Column(9).Width = 0
        .Column(11).Width = 130
        

        For X = 3 To 6
            .Column(X).Width = 60
        Next X
        .Column(7).Width = 65
        .Column(10).Width = 65

        .Cell(0, 1).Text = "DATE"
        .Cell(0, 2).Text = "DAY"
        .Cell(0, 3).Text = "IN-AM"
        .Cell(0, 4).Text = "OUT-AM"
        .Cell(0, 5).Text = "IN-PM"
        .Cell(0, 6).Text = "OUT-PM"
        .Cell(0, 7).Text = "SHIFT-IN"
        .Cell(0, 8).Text = "SHIFT-OUT-AM"
        .Cell(0, 9).Text = "SHIFT-IN-PM"
        .Cell(0, 10).Text = "SHIFT-OUT"
        .Cell(0, 11).Text = "REMARKS"

        .Column(1).CellType = cellCalendar
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(7).Locked = True
        .Column(10).Locked = True
        .Column(11).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 1, .Rows - 1, 10).ClearText
    
    End With


End Sub

Sub fillfirstcutoffgrid()
Dim Criteria                                                          As String
    Dim k                                                             As Integer
    Dim r                                                             As Integer
    Screen.MousePointer = 11
    If cboMonth.Text = Date2Month(OneMonth(Date, -2)) Then
        TransactionDate = OneMonth(Date, -2)
    ElseIf cboMonth.Text = Date2Month(OneMonth(Date, -1)) Then
        TransactionDate = OneMonth(Date, -1)
    ElseIf cboMonth.Text = Date2Month(Date) Then
        TransactionDate = Date
    ElseIf cboMonth.Text = Date2Month(OneMonth(Date, 1)) Then
        TransactionDate = OneMonth(Date, 1)
    End If

    If cboEmpNumber.Text = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    initgrid1
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "Select * from HRMS_EmpInfo Where EmpNo  =  " & N2Str2Null(cboEmpNumber.Text), gconDMIS
    If rsEmpInfo.EOF Then
        For k = 1 To 150: Beep: Next k
        MsgBox "Employee Number Not Found", vbInformation, "Empty Record"
        cboEmpNumber.Text = ""
        On Error Resume Next
        Screen.MousePointer = 0
        cboEmpNumber.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    Else
        
        varEmpNo = rsEmpInfo!EMPNO
        TxtEmpName.Text = rsEmpInfo!lastname & ", " & rsEmpInfo!FIRSTNAME
        If Null2String(rsEmpInfo!PICFILNAME) <> "" Then
            On Error Resume Next
            If Len(Dir(HRMS_PICTURES_PATH & Null2String(rsEmpInfo!PICFILNAME))) > 0 Then
                LoadPic imgDispPic, HRMS_PICTURES_PATH & Null2String(rsEmpInfo!PICFILNAME)
            Else
                LoadPic imgDispPic, HRMS_PICTURES_PATH & "BlankFace.JPG"
            End If
        Else
            LoadPic imgDispPic, HRMS_PICTURES_PATH & "BlankFace.JPG"
        End If

        If Null2String(rsEmpInfo!ACTIVEINACTIVE) = "I" Then
            For k = 1 To 150: Beep: Next k
            MsgBox "Employee Not Active", 0, "Inactive Employee"
            On Error Resume Next
            cboEmpNumber.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        '=======================
        'TO DISPLAY BY CUT OFF
        Screen.MousePointer = 11
        Dim BegDayFirstCutOff, BegDaySecondCutOff                     As Integer
        Dim EndDayFirstCutOff, EndDaySecondCutOff                     As Integer
        Dim INCREMENTING_DAY                                          As Integer
        Dim I                                                         As Long
        'InitGrid
        Dim rsPayrollSetup                                            As ADODB.Recordset
        Set rsPayrollSetup = New ADODB.Recordset
        Set rsPayrollSetup = gconDMIS.Execute("Select * from HRMS_PayrollSetup")
        If Not rsPayrollSetup.EOF And Not rsPayrollSetup.BOF Then
            BegDayFirstCutOff = N2Str2Zero(rsPayrollSetup!FROMDATE1)
            BegDaySecondCutOff = N2Str2Zero(rsPayrollSetup!FROMDATE2)
            EndDayFirstCutOff = N2Str2Zero(rsPayrollSetup!TODATE1)
            EndDaySecondCutOff = N2Str2Zero(rsPayrollSetup!TODATE2)

            If BegDayFirstCutOff > EndDayFirstCutOff Then
                TransactionDate = DateSerial(cboYear, What_month(Trim(cboMonth)) - 1, BegDayFirstCutOff)
            Else
                TransactionDate = DateSerial(cboYear, What_month(Trim(cboMonth)), BegDayFirstCutOff)
            End If
            INCREMENTING_DAY = BegDayFirstCutOff: r = 0
            Grid1.AutoRedraw = True
            
            Grid1.Rows = DateDiff("d", CDate(Null2String(Format(TransactionDate, "mm/dd/yyyy"))), Null2String(CDate(Format(DateSerial(cboYear, What_month(RTrim(LTrim(cboMonth))), EndDayFirstCutOff), "mm/dd/yyyy")))) + 2
            Do While CDate(Format(TransactionDate, "mm/dd/yyyy")) <= CDate(Format(DateSerial(cboYear, What_month(RTrim(LTrim(cboMonth))), EndDayFirstCutOff), "mm/dd/yyyy"))
                INCREMENTING_DAY = INCREMENTING_DAY + 1
                r = r + 1
                Set rsCard = New ADODB.Recordset
                rsCard.Open "Select * from HRMS_Attend where Empno = " & N2Str2Null(varEmpNo) & " AND convert(varchar,DateToday,101) = '" & Format(TransactionDate, "mm/dd/yyyy") & "'", gconDMIS
                       
                Grid1.Cell(r, 1).Text = Format(TransactionDate, "mm/dd/yy")
                Grid1.Cell(r, 2).Text = UCase(WeekdayName(Weekday(TransactionDate), True))
                Grid1.Cell(r, 11).Text = IsHolyday(MONTH(TransactionDate), Day(TransactionDate))
                    
                    
                    If UCase(Grid1.Cell(r, 2).Text) = "SAT" Or UCase(Grid1.Cell(r, 2).Text) = "SUN" Then
                            
                            'do nothing
                    Else
                                
                                With Grid1
                                    .Cell(r, 3).Text = "08:00 AM"
                                    .Cell(r, 6).Text = "05:00 PM"
                                    .Cell(r, 7).Text = "08:00 AM"
                                    .Cell(r, 10).Text = "05:00 PM"
                                End With
                        
                    End If
                
                
                With Grid1
                    I = 0
                    If LTrim(RTrim(.Cell(r, 11).Text)) <> "" Then
                        .Range(r, 0, r, 11).BackColor = &H69CBFA
                    ElseIf UCase(Grid1.Cell(r, 2).Text) = "SAT" Or UCase(Grid1.Cell(r, 2).Text) = "SUN" Then
                        .Range(r, 0, r, 11).BackColor = &HC0FFC0
                    Else
                        For I = 0 To 11
                            If r Mod 2 = 0 Then
                                .Cell(r, I).BackColor = .BackColor1
                            Else
                                .Cell(r, I).BackColor = .BackColor2
                            End If
                        Next
                    End If
                End With
                TransactionDate = CDate(TransactionDate + 1)
            Loop
            Grid1.AutoRedraw = True: Grid1.Refresh

        End If
        'END DISPLAY BY CUT OFF
    End If

   Screen.MousePointer = 0
End Sub



Sub fillsecondcutoffgrid()

Dim Criteria                                                          As String
    Dim k                                                             As Integer
    Dim r                                                             As Integer
    Screen.MousePointer = 11
    If cboMonth.Text = Date2Month(OneMonth(Date, -2)) Then
        TransactionDate = OneMonth(Date, -2)
    ElseIf cboMonth.Text = Date2Month(OneMonth(Date, -1)) Then
        TransactionDate = OneMonth(Date, -1)
    ElseIf cboMonth.Text = Date2Month(Date) Then
        TransactionDate = Date
    ElseIf cboMonth.Text = Date2Month(OneMonth(Date, 1)) Then
        TransactionDate = OneMonth(Date, 1)
    End If

    If cboEmpNumber.Text = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    
    initgrid2
    
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "Select * from HRMS_EmpInfo Where EmpNo  =  " & N2Str2Null(cboEmpNumber.Text), gconDMIS
    If rsEmpInfo.EOF Then
        For k = 1 To 150: Beep: Next k
        MsgBox "Employee Number Not Found", vbInformation, "Empty Record"
        cboEmpNumber.Text = ""
        On Error Resume Next
        Screen.MousePointer = 0
        cboEmpNumber.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    Else
        
        varEmpNo = rsEmpInfo!EMPNO
        TxtEmpName.Text = rsEmpInfo!lastname & ", " & rsEmpInfo!FIRSTNAME
        If Null2String(rsEmpInfo!PICFILNAME) <> "" Then
            On Error Resume Next
            If Len(Dir(HRMS_PICTURES_PATH & Null2String(rsEmpInfo!PICFILNAME))) > 0 Then
                LoadPic imgDispPic, HRMS_PICTURES_PATH & Null2String(rsEmpInfo!PICFILNAME)
            Else
                LoadPic imgDispPic, HRMS_PICTURES_PATH & "BlankFace.JPG"
            End If
        Else
            LoadPic imgDispPic, HRMS_PICTURES_PATH & "BlankFace.JPG"
        End If

        If Null2String(rsEmpInfo!ACTIVEINACTIVE) = "I" Then
            For k = 1 To 150: Beep: Next k
            MsgBox "Employee Not Active", 0, "Inactive Employee"
            On Error Resume Next
            cboEmpNumber.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        '=======================
        'TO DISPLAY BY CUT OFF
        Screen.MousePointer = 11
        Dim BegDayFirstCutOff, BegDaySecondCutOff                     As Integer
        Dim EndDayFirstCutOff, EndDaySecondCutOff                     As Integer
        Dim INCREMENTING_DAY                                          As Integer
        Dim I                                                         As Long
        'InitGrid
        Dim rsPayrollSetup                                            As ADODB.Recordset
        Set rsPayrollSetup = New ADODB.Recordset
        Set rsPayrollSetup = gconDMIS.Execute("Select * from HRMS_PayrollSetup")
        If Not rsPayrollSetup.EOF And Not rsPayrollSetup.BOF Then
            BegDayFirstCutOff = N2Str2Zero(rsPayrollSetup!FROMDATE1)
            BegDaySecondCutOff = N2Str2Zero(rsPayrollSetup!FROMDATE2)
            EndDayFirstCutOff = N2Str2Zero(rsPayrollSetup!TODATE1)
            EndDaySecondCutOff = N2Str2Zero(rsPayrollSetup!TODATE2)
            
            I = 0
            If BegDaySecondCutOff > EndDaySecondCutOff Then
                TransactionDate = DateSerial(cboYear, What_month(Trim(cboMonth)) - 1, BegDaySecondCutOff)
            Else
                TransactionDate = DateSerial(cboYear, What_month(Trim(cboMonth)), BegDaySecondCutOff)
            End If
            INCREMENTING_DAY = BegDaySecondCutOff: r = 0

            Grid2.Rows = DateDiff("d", Null2String(CDate(TransactionDate)), Null2String(CDate(DateSerial(cboYear, What_month(Trim(cboMonth)), EndDaySecondCutOff)))) + 2
            Do While CDate(TransactionDate) <= CDate(DateSerial(cboYear, What_month(Trim(cboMonth)), EndDaySecondCutOff))
                INCREMENTING_DAY = INCREMENTING_DAY + 1: r = r + 1
                Set rsCard = New ADODB.Recordset
                rsCard.Open "Select * from HRMS_Attend where Empno = " & N2Str2Null(varEmpNo) & " AND DateToday = '" & TransactionDate & "'", gconDMIS
                Grid2.Cell(r, 2).Text = UCase(WeekdayName(Weekday(TransactionDate), True))
                Grid2.Cell(r, 1).Text = Format(TransactionDate, "mm/dd/yy")
                Grid2.Cell(r, 11).Text = IsHolyday(MONTH(TransactionDate), Day(TransactionDate))

                            
                 If UCase(Grid2.Cell(r, 2).Text) = "SAT" Or UCase(Grid2.Cell(r, 2).Text) = "SUN" Then
                            'do nothing
                    Else
                                With Grid2
                                    .Cell(r, 3).Text = "08:00 AM"
                                    .Cell(r, 6).Text = "05:00 PM"
                                    .Cell(r, 7).Text = "08:00 AM"
                                    .Cell(r, 10).Text = "05:00 PM"
                                End With

                    End If

        
                I = 0
                With Grid2
                    If LTrim(RTrim(.Cell(r, 11).Text)) <> "" Then
                        .Range(r, 0, r, 11).BackColor = &H69CBFA
                    ElseIf UCase(Grid2.Cell(r, 2).Text) = "SAT" Or UCase(Grid2.Cell(r, 2).Text) = "SUN" Then
                        .Range(r, 0, r, 11).BackColor = &HC0FFC0
                    Else
                        For I = 1 To 11
                            If r Mod 2 = 0 Then
                                .Cell(r, I).BackColor = .BackColor1
                            Else
                                .Cell(r, I).BackColor = .BackColor2
                            End If
                        Next
                    End If
                End With
                TransactionDate = CDate(TransactionDate + 1)
            Loop
        End If
        'END DISPLAY BY CUT OFF
    End If
    Screen.MousePointer = 0

End Sub


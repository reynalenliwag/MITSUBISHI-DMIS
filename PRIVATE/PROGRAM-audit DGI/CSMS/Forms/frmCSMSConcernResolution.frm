VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMSConcernResolution 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Concern Resolution"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmCSMSConcernResolution.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   7815
   Begin VB.PictureBox FrameInfo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   60
      ScaleHeight     =   6645
      ScaleWidth      =   7665
      TabIndex        =   31
      Top             =   60
      Width           =   7695
      Begin VB.TextBox cboCustomer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   1410
         Width           =   7455
      End
      Begin VB.TextBox txtveh 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   5
         Top             =   2670
         Width           =   4860
      End
      Begin VB.TextBox txtcontact 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   2670
         Width           =   2340
      End
      Begin VB.TextBox txtaddress 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   2010
         Width           =   7455
      End
      Begin VB.TextBox txtNature 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3960
         Width           =   7455
      End
      Begin VB.TextBox txtAction 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   5070
         Width           =   7455
      End
      Begin VB.TextBox txtnoted 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4050
         TabIndex        =   12
         Top             =   6150
         Width           =   3525
      End
      Begin VB.TextBox txtSubmitted 
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
         TabIndex        =   11
         Top             =   6150
         Width           =   3540
      End
      Begin VB.TextBox txtWeekNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6570
         TabIndex        =   1
         Top             =   660
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPNextCall 
         Height          =   345
         Left            =   2370
         TabIndex        =   7
         Top             =   3300
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56819713
         CurrentDate     =   39246
      End
      Begin MSComCtl2.DTPicker DTPReleasedate 
         Height          =   345
         Left            =   4770
         TabIndex        =   8
         Top             =   3300
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56819713
         CurrentDate     =   39246
      End
      Begin MSComCtl2.DTPicker DTPdateCalled 
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   3300
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56819713
         CurrentDate     =   39246
      End
      Begin Crystal.CrystalReport rptConcernReport 
         Left            =   7110
         Top             =   3270
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker DTPdate 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   630
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56819713
         CurrentDate     =   39246
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   -30
         TabIndex        =   50
         Top             =   0
         Width           =   9825
         _Version        =   655364
         _ExtentX        =   17330
         _ExtentY        =   661
         _StockProps     =   14
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         ForeColor       =   -2147483630
      End
      Begin VB.Label labid 
         Height          =   345
         Left            =   5250
         TabIndex        =   46
         Top             =   570
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   44
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contact no"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   43
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date Released"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4785
         TabIndex        =   42
         Top             =   3060
         Width           =   1230
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vehicle type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2880
         TabIndex        =   41
         Top             =   2460
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date Next Called"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2400
         TabIndex        =   40
         Top             =   3060
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   39
         Top             =   1770
         Width           =   720
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date Called"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   38
         Top             =   3060
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nature of Concern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   37
         Top             =   3690
         Width           =   1545
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Action Taken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   4860
         Width           =   1110
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Noted By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4095
         TabIndex        =   35
         Top             =   5880
         Width           =   750
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Submitted By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   75
         TabIndex        =   34
         Top             =   5880
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Week no"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6630
         TabIndex        =   33
         Top             =   390
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   32
         Top             =   390
         Width           =   390
      End
   End
   Begin VB.PictureBox picADD 
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
      Height          =   975
      Left            =   960
      ScaleHeight     =   975
      ScaleWidth      =   6945
      TabIndex        =   27
      Top             =   6780
      Width           =   6945
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   825
         Left            =   6030
         MouseIcon       =   "frmCSMSConcernResolution.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   825
         Left            =   5310
         MouseIcon       =   "frmCSMSConcernResolution.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   825
         Left            =   4590
         MouseIcon       =   "frmCSMSConcernResolution.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   825
         Left            =   3870
         MouseIcon       =   "frmCSMSConcernResolution.frx":1E6F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":1FC1
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   825
         Left            =   3150
         MouseIcon       =   "frmCSMSConcernResolution.frx":231D
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":246F
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   825
         Left            =   2430
         MouseIcon       =   "frmCSMSConcernResolution.frx":2782
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":28D4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   825
         Left            =   1710
         MouseIcon       =   "frmCSMSConcernResolution.frx":2BCE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":2D20
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   825
         Left            =   990
         MouseIcon       =   "frmCSMSConcernResolution.frx":3078
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":31CA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picEdeted 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   60
      ScaleHeight     =   7545
      ScaleWidth      =   7665
      TabIndex        =   25
      Top             =   60
      Visible         =   0   'False
      Width           =   7695
      Begin XtremeReportControl.ReportControl ListView1 
         Height          =   6585
         Left            =   60
         TabIndex        =   24
         Top             =   900
         Width           =   7545
         _Version        =   655364
         _ExtentX        =   13309
         _ExtentY        =   11615
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         EditOnClick     =   0   'False
      End
      Begin VB.CommandButton cmdCANCELSEARCH 
         Caption         =   "X"
         Height          =   405
         Left            =   7200
         TabIndex        =   49
         Top             =   450
         Width           =   405
      End
      Begin VB.TextBox txtkeyword 
         Height          =   360
         Left            =   1500
         TabIndex        =   23
         Top             =   480
         Width           =   5595
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Enter Keyword"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   48
         Top             =   570
         Width           =   1185
      End
      Begin XtremeShortcutBar.ShortcutCaption cap 
         Height          =   375
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   9795
         _Version        =   655364
         _ExtentX        =   17277
         _ExtentY        =   661
         _StockProps     =   14
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   30
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label14 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         MouseIcon       =   "frmCSMSConcernResolution.frx":3529
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   6360
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.PictureBox picSave 
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
      Height          =   990
      Left            =   5700
      ScaleHeight     =   990
      ScaleWidth      =   2115
      TabIndex        =   28
      Top             =   6780
      Visible         =   0   'False
      Width           =   2115
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   825
         Left            =   1290
         MouseIcon       =   "frmCSMSConcernResolution.frx":3833
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":3985
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   825
         Left            =   570
         MouseIcon       =   "frmCSMSConcernResolution.frx":3CC3
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSConcernResolution.frx":3E15
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblIDcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   29
         Top             =   210
         Width           =   825
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   765
      Index           =   0
      Left            =   8190
      TabIndex        =   45
      Top             =   3720
      Width           =   1515
      _Version        =   655364
      _ExtentX        =   2672
      _ExtentY        =   1349
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCSMSConcernResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thecode                                             As String
Dim theID                                               As String
Dim ADDOREDIT                                           As String
Dim rsConcern                                           As ADODB.Recordset

Sub rsRefresh()
    Set rsConcern = New ADODB.Recordset
    rsConcern.Open "SELECT * From CSMS_ConcernResolution  order by id", gconDMIS, adOpenKeyset, adLockOptimistic
End Sub

Sub StoreMemvars()
    If Not (rsConcern.BOF And rsConcern.EOF) Then
        DTPdate.Value = Null2String(rsConcern!Thedate)
        txtWeekNo.Text = NumericVal(rsConcern!WEEKNO)
        cbocustomer.Text = Null2String(rsConcern!customername)
        txtAddress.Text = Null2String(rsConcern!Address)
        txtcontact.Text = Null2String(rsConcern!CONTACTNO)
        txtveh.Text = Null2String(rsConcern!VEHICLETYPE)
        DTPdateCalled.Value = Null2String(rsConcern!DATECALLED)
        DTPNextCall.Value = Null2String(rsConcern!NEXTCALLDATE)
        DTPReleasedate.Value = Null2String(rsConcern!RELEASEDATE)
        txtNature.Text = Null2String(rsConcern!NATUREOFCONCERN)
        txtAction.Text = Null2String(rsConcern!ACTIONTAKEN)
        txtSubmitted.Text = Null2String(rsConcern!SUBMITTEDBY)
        txtnoted.Text = Null2String(rsConcern!NotedBy)
        labID.Caption = rsConcern!ID
    Else
        Call ShowNoRecord
        Call cmdAdd_Click
    End If
End Sub


Private Sub cmdCancelSearch_Click()
    picEdeted.Visible = False
    picEdeted.ZOrder 1
End Sub

Private Sub cmdNext_Click()
    rsConcern.MoveNext
    If rsConcern.EOF Then
        rsConcern.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsConcern.MovePrevious
    If rsConcern.BOF Then
        rsConcern.MoveFirst
        Call ShowLastRecordMsg
    End If
    Call StoreMemvars
End Sub

Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CONCERN RESOLUTION") = False Then Exit Sub

    If MsgBox("Are You Sure Do You Want to print this record?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        Screen.MousePointer = 11
        rptConcernReport.WindowTitle = "Concern Resolution"
        PrintSQLReport rptConcernReport, CSMS_REPORT_PATH & "CSMS_concernReport.rpt", "{CSMS_ConcernResolution.id}= " & labID & "", CSMS_REPORT_CONNECTION, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub displayCustomer()
    Dim RS                                             As New ADODB.Recordset

    Set RS = gconDMIS.Execute("SELECT Acctname FROM ALL_Customer")
    cbocustomer.Clear
    With RS
        Do While Not .EOF
            cbocustomer.AddItem RS!ACCTNAME
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub DisableMe(ByVal x As Boolean)
    FrameInfo.Enabled = Not x
    txtAction.Enabled = Not x
    txtNature.Enabled = Not x
    txtSubmitted.Enabled = Not x
    txtnoted.Enabled = Not x
    txtWeekNo.Enabled = Not x
End Sub

Sub initMemvars()
    DTPdate.Value = Date
    cbocustomer = "":
    txtAddress.Text = ""
    txtcontact.Text = "":
    txtveh.Text = ""
    DTPReleasedate.Value = Date
    DTPdateCalled.Value = Date
    DTPNextCall = Date
    txtNature.Text = ""
    txtAction.Text = ""
    txtSubmitted.Text = "":
    txtnoted.Text = ""
    DTPdate = Date
    txtWeekNo.Text = ""
    theID = ""
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CONCERN RESOLUTION") = False Then Exit Sub

    ADDOREDIT = "ADD"
    Call initMemvars
    FrameInfo.Enabled = True
    picADD.Visible = False
    picSave.Visible = True
    
    On Error Resume Next
    DTPdate.SetFocus
End Sub

Private Sub cmdCancel_Click()
    FrameInfo.Enabled = False
    picADD.Visible = True
    picSave.Visible = False
    FrameInfo.Enabled = False
    
    Call StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "CONCERN RESOLUTION") = False Then Exit Sub

    If MsgBox("delete this record, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        
    gconDMIS.Execute ("Delete From CSMS_ConcernResolution Where ID = " & labID & "")
    LogAudit "X", "CONCERN RESOLUTION", cbocustomer
    Call ShowDeletedMsg
        
    Call rsRefresh
    Call StoreMemvars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "CONCERN RESOLUTION") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    
    FrameInfo.Enabled = True
    picADD.Visible = False
    picSave.Visible = True
    picEdeted.Enabled = False
    
    DTPdate.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    Call FillConcern
    
    picEdeted.Visible = True
    picEdeted.ZOrder 0
    picEdeted.Enabled = True
    
    txtKeyword.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim SQL                                             As String
    Dim theCustomer                                     As String
    Dim theVehicle                                      As String
    Dim theAddress                                      As String
    Dim theContact                                      As String
    Dim theReleaseDate                                  As String
    Dim thedateCalled                                   As String
    Dim theNextCall                                     As String
    Dim theSubmitted                                    As String
    Dim thenoted                                        As String
    Dim theNature                                       As String
    Dim theAction                                       As String
    Dim Thedate                                         As String
    Dim theWeekNo                                       As String

    theCustomer = cbocustomer.Text
    theAddress = Trim(txtAddress.Text)
    theContact = Trim(txtcontact.Text)
    theVehicle = Trim(txtveh.Text)
    theReleaseDate = DTPReleasedate
    thedateCalled = DTPdateCalled
    theNextCall = DTPNextCall
    theNature = Trim(txtNature.Text)
    theAction = Trim(txtAction.Text)
    theSubmitted = Trim(txtSubmitted.Text)
    thenoted = Trim(txtnoted.Text)
    Thedate = DTPdate
    theWeekNo = Trim(txtWeekNo.Text)
    Dim ans                                            As String

    If Len(theCustomer) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Customer"
        On Error Resume Next
        cbocustomer.SetFocus
        Exit Sub
    End If

    If Len(theAddress) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Address"
        On Error Resume Next
        txtAddress.SetFocus
        Exit Sub
    End If

    If Len(theContact) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Contact No"
        On Error Resume Next
        txtcontact.SetFocus
        Exit Sub
    End If

    If Len(theVehicle) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Vehicle Model"
        On Error Resume Next
        txtveh.SetFocus
        Exit Sub
    End If


    If Len(theNature) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Nature Of Concern"
        On Error Resume Next
        txtNature.SetFocus
        Exit Sub
    End If

    If Len(theAction) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Action Taken"
        On Error Resume Next
        txtAction.SetFocus
        Exit Sub
    End If

    If Len(theSubmitted) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Submitted By"
        On Error Resume Next
        txtSubmitted.SetFocus
        Exit Sub
    End If

    If Len(thenoted) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Submitted By"
        On Error Resume Next
        txtnoted.SetFocus
        Exit Sub
    End If

    If Len(theWeekNo) = 0 Then
        ShowIsRequiredMsg "Missing Parameters..Week No"
        On Error Resume Next
        txtWeekNo.SetFocus
        Exit Sub
    End If
    
    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute ("INSERT INTO CSMS_concernResolution (CUSTOMERNAME, ADDRESS, CONTACTNO, VEHICLETYPE, RELEASEDATE, DATECALLED, NEXTCALLDATE, NATUREOFCONCERN, ACTIONTAKEN, SUBMITTEDBY, NOTEDBY, THEDATE, WEEKNO)" & _
            " VALUES('" & theCustomer & "','" & theAddress & _
            "', '" & theContact & _
            "', '" & theVehicle & _
            "', '" & theReleaseDate & _
            "', '" & thedateCalled & _
            "', '" & theNextCall & _
            "', '" & theNature & _
            "', '" & theAction & _
            "', '" & theSubmitted & _
            "', '" & thenoted & _
            "', '" & Thedate & _
            "', '" & theWeekNo & "')")
        'LogAudit "A", "CONCERN RESOLUTION", cboCustomer

        Call ShowSuccessFullyAdded
    Else
        gconDMIS.Execute ("UPDATE CSMS_concernResolution set " & _
            " customername = '" & theCustomer & _
            "', address = '" & theAddress & _
            "', contactNo = '" & theContact & _
            "', vehicleType = '" & theVehicle & _
            "', ReleaseDate = '" & theReleaseDate & _
            "', datecalled = '" & thedateCalled & _
            "', nextCalldate = '" & theNextCall & _
            "', NatureOfConcern = '" & theNature & _
            "', actiontaken = '" & theAction & _
            "', SubmittedBy = '" & theSubmitted & _
            "', notedBy = '" & thenoted & _
            "', thedate = '" & Thedate & _
            "', Weekno = '" & theWeekNo & _
            "' WHERE ID = '" & labID & "'")
        'LogAudit "E", "CONCERN RESOLUTION", cboCustomer

        Call ShowSuccessFullyUpdated
    End If
    
    picSave.Visible = False
    picADD.Visible = True
    Call cmdCancel_Click
    
    Call FillConcern
    Call rsRefresh
    Call StoreMemvars
    
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    With ListView1
        .Columns.Add 0, "Date", 100, True
        .Columns.Add 1, "Customer Name", 100, True
        .Columns.Add 2, "Nature of Concern", 100, True
        .Columns.Add 3, "Action Taken", 100, True
        .Columns.Add 4, "Submitted by", 100, True
        .Columns.Add 5, "Noted By", 100, True
        .Columns.Add 6, "ID", 0, False
        
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.GroupRowTextBold = True
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
        .SetCustomDraw xtpCustomBeforeDrawRow
    End With
    
    Call FillConcern
    Call rsRefresh
    Call StoreMemvars
End Sub

Sub FillConcern()
    Dim SQL                                            As String

    SQL = "SELECT THEDATE, CUSTOMERNAME, NATUREOFCONCERN, ACTIONTAKEN, SUBMITTEDBY, NOTEDBY, ID FROM CSMS_ConcernResolution order by id desc"

    Dim RecSet                                         As New ADODB.Recordset
    Dim fld                                            As Field
    Dim j                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord
    Set RecSet = gconDMIS.Execute(SQL)
    
    ListView1.Records.DeleteAll
    While Not RecSet.EOF
        j = j + 1
        Set REC = ListView1.Records.Add
        For Each fld In RecSet.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RecSet.MoveNext
    Wend
    
    ListView1.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RecSet = Nothing
End Sub

Private Sub Label14_Click()
    picEdeted.Visible = False
    picSave.Visible = False
    picADD.Visible = True
End Sub

Private Sub ListView1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then Exit Sub
    On Error GoTo ERRORCODE
    
    rsConcern.MoveFirst
    rsConcern.Find ("ID=" & Item.Record(6).Value)
    Call StoreMemvars
    Call cmdCancelSearch_Click
    Exit Sub

ERRORCODE:
    ShowVBError
End Sub

Private Sub txtKeyword_Change()
    ListView1.FilterText = txtKeyword
    ListView1.Populate
End Sub


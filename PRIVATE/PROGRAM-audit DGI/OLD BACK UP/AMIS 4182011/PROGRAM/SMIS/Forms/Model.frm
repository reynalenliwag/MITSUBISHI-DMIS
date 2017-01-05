VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Files_Model 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Models"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Model.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   6420
   Begin VB.CommandButton cmdVSpec 
      Caption         =   "Add Vehicle Spec"
      Height          =   345
      Left            =   2880
      TabIndex        =   20
      Top             =   1950
      Width           =   1575
   End
   Begin VB.CommandButton cmdVcost 
      Caption         =   "Add Cost / Price"
      Height          =   345
      Left            =   1320
      TabIndex        =   21
      Top             =   1950
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update List"
      Height          =   990
      Left            =   5040
      MaskColor       =   &H0000FFFF&
      Picture         =   "Model.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Update List"
      Top             =   270
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vehicle Model Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   6225
      Begin VB.ComboBox cbomodel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1260
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1020
         Width           =   3495
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   180
         Top             =   390
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cboVehicleMode_AssignedSC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1260
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   1470
         Width           =   3495
      End
      Begin VB.TextBox txtCode 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1260
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   270
         Width           =   1725
      End
      Begin VB.TextBox txtDescript 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1260
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   660
         Width           =   3465
      End
      Begin VB.TextBox txtModel 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   6240
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1890
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1950
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned SC"
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
         Left            =   150
         TabIndex        =   13
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   690
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "Label9"
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
         Left            =   1380
         TabIndex        =   9
         Top             =   660
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label32 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         TabIndex        =   7
         Top             =   660
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label31 
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label28 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   630
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame fraDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   45
      TabIndex        =   32
      Top             =   2400
      Width           =   6210
      Begin VB.OptionButton optModel 
         Caption         =   "Model"
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
         Left            =   3405
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   225
         Width           =   1350
      End
      Begin VB.OptionButton optCode 
         Caption         =   "&Code"
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
         Left            =   1125
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   255
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton optDescription 
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
         Height          =   285
         Left            =   2010
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   225
         Width           =   1350
      End
      Begin MSComctlLib.ListView lstVehicleModel 
         Height          =   2685
         Left            =   60
         TabIndex        =   38
         Top             =   1050
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   4736
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "Model.frx":0BD4
         NumItems        =   0
      End
      Begin VB.TextBox txtSearch 
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
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   660
         Width           =   6075
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search by:"
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
         Left            =   135
         TabIndex        =   33
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.PictureBox picCost 
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
      Height          =   2415
      Left            =   1320
      ScaleHeight     =   2385
      ScaleWidth      =   3705
      TabIndex        =   22
      Top             =   2250
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtLTO 
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
         ForeColor       =   &H00400000&
         Height          =   405
         Left            =   990
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   1350
         Width           =   2580
      End
      Begin VB.TextBox txtSRP 
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
         ForeColor       =   &H00400000&
         Height          =   405
         Left            =   990
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   900
         Width           =   2580
      End
      Begin VB.TextBox txtUnitPrice 
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
         ForeColor       =   &H00400000&
         Height          =   405
         Left            =   990
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   450
         Width           =   2580
      End
      Begin wizButton.cmd CmdCost 
         Height          =   375
         Left            =   1350
         TabIndex        =   30
         Top             =   1860
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         TX              =   "Ok"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Model.frx":0D36
      End
      Begin wizButton.cmd CmdCostCancel 
         Height          =   375
         Left            =   2490
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Model.frx":0D52
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "LTO "
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
         Left            =   -330
         TabIndex        =   28
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
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
         Left            =   -330
         TabIndex        =   26
         Top             =   990
         Width           =   1245
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   3705
         _Version        =   655364
         _ExtentX        =   6535
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Vehicle Cost / Price Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
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
         Left            =   120
         TabIndex        =   25
         Top             =   510
         Width           =   1245
      End
   End
   Begin VB.PictureBox picAdds 
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
      Height          =   945
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   6585
      TabIndex        =   39
      Top             =   6240
      Width           =   6585
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   5580
         MouseIcon       =   "Model.frx":0D6E
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":0EC0
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   4890
         MouseIcon       =   "Model.frx":1226
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":1378
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   4200
         MouseIcon       =   "Model.frx":16DE
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":1830
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3510
         MouseIcon       =   "Model.frx":1B5B
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":1CAD
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2820
         MouseIcon       =   "Model.frx":2009
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":215B
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   2130
         MouseIcon       =   "Model.frx":246E
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   1440
         MouseIcon       =   "Model.frx":28BA
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":2A0C
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   750
         MouseIcon       =   "Model.frx":2D64
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":2EB6
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Frame picSpecs 
      Caption         =   "Model Specification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7125
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6315
      Begin wizButton.cmd cmdaddSpec 
         Height          =   405
         Left            =   3270
         TabIndex        =   2
         Top             =   6630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   714
         TX              =   "Update"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Model.frx":3215
      End
      Begin wizButton.cmd cmd1 
         Height          =   405
         Left            =   4800
         TabIndex        =   3
         Top             =   6630
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   714
         TX              =   "Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Model.frx":3231
      End
      Begin VB.TextBox txtspec 
         Height          =   6285
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   270
         Width           =   6165
      End
   End
   Begin VB.PictureBox picSaves 
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
      Height          =   885
      Left            =   4875
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   48
      Top             =   6270
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "Model.frx":324D
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":339F
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Cancel"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "Model.frx":36DD
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":382F
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Save this Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   18
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
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
      Left            =   600
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmSMIS_Files_Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsModel                                                           As ADODB.Recordset
Dim rsMRRINV                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String
Attribute AddorEdit.VB_VarUserMemId = 1073938435

Function GetEmpCode(XXX)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT SAECODE FROM SMIS_vw_Srep  WHERE name=" & N2Str2Null(XXX))
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        GetEmpCode = Null2String(TEMPRS.Fields("SAECODE"))
    End If
    Set TEMPRS = Nothing
End Function

Function SetEmpCode(XXX)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT [name] FROM SMIS_vw_Srep WHERE SAECODE=" & N2Str2Null(XXX))
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        SetEmpCode = Null2String(TEMPRS.Fields("name"))
    End If
    Set TEMPRS = Nothing
End Function

Sub InitMemVars()
    TXTCODE.Text = ""
    txtDescript.Text = ""
    txtModel.Text = ""
    cboModel = ""
End Sub

Sub rsREFRESH()
    Set rsModel = New ADODB.Recordset
    rsModel.Open "SELECT  *  from ALL_MODEL  order by ID DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsModel.EOF And Not rsModel.BOF Then
        labid.Caption = rsModel!ID
        TXTCODE.Text = Null2String(rsModel!CODE)
        txtDescript.Text = Null2String(rsModel!DESCRIPT)
        txtModel.Text = UCase(Null2String(rsModel!Model))
        cboVehicleMode_AssignedSC = SetEmpCode(Null2String(rsModel!SAECODE))
        txtUnitPrice = FormatNumber(NumericVal(rsModel!costprice))
        txtSRP = FormatNumber(NumericVal(rsModel!unitcost))
        txtLto = FormatNumber(NumericVal(rsModel!LTO))
        txtspec = Replace(Null2String(rsModel!spec), "†", "'")
        cboModel.Text = UCase(Null2String(rsModel!Model))
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsTemp                                                        As ADODB.Recordset
    lstVehicleModel.Sorted = False
    lstVehicleModel.ListItems.Clear
    lstVehicleModel.Enabled = False
    Set rsTemp = New ADODB.Recordset

    If optCode.Value = True Then
        Set rsTemp = gconDMIS.Execute("SELECT CODE, UPPER(MODEL), descript, ID FROM ALL_MODEL WHERE CODE like'" & ReplaceQuote(XXX) & "%' order by CODE asc")
    ElseIf optDescription.Value = True Then
        Set rsTemp = gconDMIS.Execute("SELECT CODE, UPPER(MODEL), descript, ID FROM ALL_MODEL WHERE descript like'" & ReplaceQuote(XXX) & "%' order by descript asc")
    Else
        Set rsTemp = gconDMIS.Execute("SELECT CODE, UPPER(MODEL), descript, ID FROM ALL_MODEL WHERE MODEL like'" & ReplaceQuote(XXX) & "%' order by MODEL asc")
    End If

    If Not (rsTemp.EOF And rsTemp.BOF) Then
        Listview_Loadval Me.lstVehicleModel.ListItems, rsTemp
        lstVehicleModel.Refresh
        lstVehicleModel.Enabled = True
    End If

End Sub

Sub FillModel()
    '    Dim SQL                             As String
    '    Dim rsModel                         As New ADODB.Recordset
    '    SQL = "SELECT Description from All_modelcode"
    '    Set rsModel = New ADODB.Recordset
    '    Set rsModel = gconDMIS.Execute(SQL)
    '    cbomodel.Clear
    '    Do While Not rsModel.EOF
    '        cbomodel.AddItem Null2String(rsModel!Description)
    '        rsModel.MoveNext
    '    Loop
    '    Set rsModel = Nothing
    '***UPDATED by RDC Aug 19, 2008
    Combo_Loadval cboModel, gconDMIS.Execute("SELECT DISTINCT(MODEL) FROM ALL_MODEL")
    '******************************
End Sub

Private Sub cboModel_CLick()
    txtModel.Text = cboModel.Text
End Sub

Private Sub cboModel_LostFocus()
    '    cbomodel.ListIndex = SelectCombo(cbomodel, cbomodel)
    '    If cbomodel.ListIndex = -1 Then
    '    cbomodel = ""
    '    End If
End Sub

Private Sub cmd1_Click()
    picSpecs.Visible = False
    txtModel.Text = ""
    cmdVSpec.Value = False
End Sub

Private Sub cmd2_Click()

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "VEHICLE MODEL") = False Then Exit Sub
    On Error GoTo ErrorCode:
    AddorEdit = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    cmdVcost.Enabled = False: cmdVSpec.Enabled = False
    InitMemVars
    lstVehicleModel.Enabled = False
    txtSEARCH.Enabled = False
    On Error Resume Next
    TXTCODE.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdaddSpec_Click()
    Dim ans                                                           As String
    ans = MsgBox("Are you sure?", vbQuestion + vbYesNo)
    If ans = vbYes Then
        picSpecs.Visible = True
        If LTrim(RTrim(txtspec)) = "" Then
            MsgBox "Pls input vehicle specification!..", vbExclamation, "Warning"
            Exit Sub
        End If
        gconDMIS.Execute ("update all_model set spec='" & Replace(txtspec, "'", "†") & "' where id=" & labid)
        rsREFRESH
        rsModel.Find ("id=" & labid)
        StoreMemVars
        picSpecs.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    lstVehicleModel.Enabled = True
    txtSEARCH.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub CmdCost_Click()
    If IsNumeric(txtUnitPrice) = False Then
        MsgBox "Invalid Amount!..", vbExclamation, "Information"
        txtUnitPrice.SetFocus
        Exit Sub
    End If

    If IsNumeric(txtSRP.Text) = False Then
        MsgBox "Invalid Amount!..", vbExclamation, "Information"
        txtSRP.SetFocus
        Exit Sub
    End If

    If IsNumeric(txtLto.Text) = False Then
        MsgBox "Invalid Amount!..", vbExclamation, "Information"
        txtLto.SetFocus
        Exit Sub
    End If
    Dim ans
    ans = MsgBox("Are you sure do want to save..", vbQuestion + vbYesNo)

    If ans = vbYes Then
        gconDMIS.Execute ("update all_model set costprice=" & NumericVal(txtUnitPrice) & ", unitcost=" & NumericVal(txtSRP) & ", lto=" & NumericVal(txtLto) & " where id=" & labid)
        rsREFRESH
        rsModel.Find ("ID=" & labid)
        StoreMemVars
        ShowHidePictureBox2 picCost, False
    End If
End Sub

Private Sub CmdCostCancel_Click()
    picCost.Visible = False
    cmdVcost.Value = False
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "VEHICLE MODEL") = False Then Exit Sub
    On Error GoTo ErrorCode
    If Not rsModel.BOF Or Not rsModel.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from All_Model where id = " & labid.Caption
            gconDMIS.Execute (SQL_STATEMENT)
            NEW_LogAudit "X", "VEHICLE MODEL", SQL_STATEMENT, labid, "", "Code: " & TXTCODE, "", ""

            'LogAudit "X", "MODEL MASTER FILE", txtModel
            rsREFRESH
            StoreMemVars
            FillSearchGrid ""
        End If
    Else
        ShowNothingToDeleteMsg
    End If

    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "VEHICLE MODEL") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    cmdVcost.Enabled = True: cmdVSpec.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    fraDetails.Enabled = False
    On Error Resume Next
    TXTCODE.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSEARCH.SetFocus

End Sub

Private Sub cmdNext_Click()
    rsModel.MoveNext
    If rsModel.EOF Then
        rsModel.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsModel.MovePrevious
    If rsModel.BOF Then
        rsModel.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "VEHICLE MODEL") = False Then Exit Sub

    Screen.MousePointer = 11
    CrystalReport1.Formulas(0) = "COMPANYNAME='" & COMPANY_NAME & "'"
    CrystalReport1.Formulas(1) = "COMPANYADDRESS='" & COMPANY_ADDRESS & "'"
    PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "Listing\MODELS.rpt", "", DMIS_REPORT_Connection, 1
    
    Call NEW_LogAudit("V", "VEHICLE MODEL", "", labid, "", "CODE: " & TXTCODE, "", "")
    'LogAudit "V", "MODEL MASTER FILE", txtModel
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    Dim vtxtCode                                                      As String
    Dim vtxtdescript                                                  As String
    Dim vtxtmodel                                                     As String
    Dim VTXTSAECODE                                                   As String
    Dim lng                                                           As Integer
    Dim THEUNITCOST                                                   As Double

    On Error GoTo ErrorCode:

    vtxtCode = N2Str2Null(TXTCODE)
    vtxtdescript = N2Str2Null(txtDescript)
    vtxtmodel = N2Str2Null(cboModel)
    VTXTSAECODE = N2Str2Null(GetEmpCode(cboVehicleMode_AssignedSC))


    lng = gconDMIS.Execute("select Count(*) from All_Model WHERE code=" & N2Str2Null(TXTCODE)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            On Error Resume Next
            TXTCODE.SetFocus
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsModel!CODE)) <> UCase(TXTCODE) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            TXTCODE.SetFocus
            Exit Sub
        End If
    End If

    ' AXP-6520071053
    If AddorEdit = "ADD" Then
        Set rsModel = New ADODB.Recordset
        rsModel.Open "select * from All_Model order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsModel.EOF And Not rsModel.BOF Then
            rsModel.MoveLast
            labid.Caption = NumericVal(rsModel!ID) + 1
        End If

        'REM UPDATE BY BTT 12172007
        SQL_STATEMENT = "Insert into ALL_Model" & _
                      " (code,descript,model, SAECODE)" & _
                      " values (" & vtxtCode & ", " & vtxtdescript & ", " & vtxtmodel & "," & VTXTSAECODE & ")"


        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "A", "VEHICLE MODEL", SQL_STATEMENT, FindTransactionID(N2Str2Null(vtxtCode), "code", "ALL_MODEL"), "", "Code: " & vtxtCode, "", ""
        ShowSuccessFullyAdded
        'LogAudit "A", "MODEL MASTER FILE", txtModel
    Else
        SQL_STATEMENT = "update All_Model set" & _
                      " code = " & vtxtCode & "," & _
                      " SAECODE = " & VTXTSAECODE & "," & _
                      " descript = " & vtxtdescript & ", model =" & vtxtmodel & _
                      " where id = " & labid.Caption


        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "E", "VEHICLE MODEL", SQL_STATEMENT, Null2String(labid), "", "Code:  " & vtxtCode, "", ""
        'LogAudit "E", "MODEL MASTER FILE", txtModel
        ShowSuccessFullyUpdated
    End If
    
    rsREFRESH
    If AddorEdit = "EDIT" Then
        rsModel.Find "id =" & labid
    End If
    cmdCancel.Value = True
    FillSearchGrid ""

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdspec_Click()

End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo ErrorCode
    Dim rsUpModel                                                     As ADODB.Recordset
    Dim cnt                                                           As Integer
    Set rsMRRINV = New ADODB.Recordset
    Dim CODE As String
    rsMRRINV.Open "select * from SMIS_MrrInv order by model,descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
        rsMRRINV.MoveFirst
        Screen.MousePointer = 11
        cnt = 0
        Do While Not rsMRRINV.EOF
            Set rsUpModel = New ADODB.Recordset


            'rsUpModel.Open "Select * from All_Model WHERE ltrim(rtrim(replace(descript,' ',''))) = '" & Replace(LTrim(RTrim(rsMRRINV!DESCRIPT)), " ", "") & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            rsUpModel.Open "Select * from All_Model WHERE ltrim(rtrim(descript)) = '" & LTrim(RTrim(rsMRRINV!DESCRIPT)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly

            If rsUpModel.EOF And rsUpModel.BOF Then
               ' If MsgBox("UNIT DESCRIPTION:" & LTrim(RTrim(rsMRRINV!DESCRIPT)) & vbCrLf & "Not yet added in Database, " & _
                '          vbCrLf & "Do you want to Add It?", vbYesNo + vbInformation) = vbYes Then


                        CODE = N2Str2Null(GenerateCode("ALL_MODEL", "CODE", "000000"))
                    gconDMIS.Execute "insert into ALL_Model " & _
                                     "(code,descript,model) " & _
                                     "values (" & UCase(CODE) & ", '" & UCase(LTrim(RTrim(rsMRRINV!DESCRIPT))) & "'," & _
                                     "'" & UCase(LTrim(RTrim(rsMRRINV!Model))) & "')"

                    cnt = cnt + 1
                'End If
            End If
            rsMRRINV.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    If cnt = 0 Then
        MessagePop InfoFriend, "UPDATE", "THE DATABASE LIST IS UP TO DATE"
    Else
        MessagePop RecSaveInfo, "UPDATED", "DATABASE LIST UPDATED"
        LogAudit "E", "MODEL MASTER FILE UPDATED FROM MRR", Now
    End If

    rsREFRESH
FillSearchGrid ""
    If AddorEdit = "EDIT" Then
        rsModel.Find "id =" & labid.Caption
    End If

    StoreMemVars


    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If picAdds.Visible = True And KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE MODEL)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "VEHICLE MODEL")
            'End If
    End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsREFRESH
    Frame1.Enabled = False
    txtSEARCH.Text = ""
    InitMemVars
    StoreMemVars
    picSpecs.Visible = False
    picAdds.Visible = True
    picSaves.Visible = False
    picCost.Visible = False
    FillModel

    'UDPATING CODE    :AXP-6520071053
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT name from SMIS_vw_Srep")
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        'cboVehicleMode_AssignedSC
        Combo_Loadval cboVehicleMode_AssignedSC, TEMPRS
    End If


    Call AddColumnHeader(" CODE, MODEL, DESCRIPTION", lstVehicleModel)
    ResizeColumnHeader lstVehicleModel, "18,25,53"
    Listview_Loadval lstVehicleModel.ListItems, gconDMIS.Execute("SELECT CODE, MODEL, descript, ID FROM ALL_MODEL")
    Screen.MousePointer = 0
End Sub

Private Sub lstVehicleModel_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstVehicleModel
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

Private Sub lstVehicleModel_DblClick()
    If lstVehicleModel.SelectedItem Is Nothing Then: Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lstVehicleModel_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsModel.MoveFirst
    Call rsModel.Find("ID=" & ITEM.ListSubItems(3).Text)
    StoreMemVars
End Sub

Private Sub Option1_Click()

End Sub

Private Sub cmdVcost_Click()
    If Module_Access(LOGID, "VEHICLE COST", "SYSTEM") = False Then Exit Sub
    If cmdVcost.Value = True Then
        ShowHidePictureBox2 picCost, True
    End If
End Sub

Private Sub cmdVSpec_Click()
    If Module_Access(LOGID, "VEHICLE SPECIFICATION", "SYSTEM") = False Then Exit Sub
    If cmdVSpec.Value = True Then
        ShowHidePictureBox2 picSpecs, True
    End If
End Sub

Private Sub txtLto_GotFocus()
    If NumericVal(txtLto.Text) <= 0 Then txtLto = ""
End Sub

Private Sub txtLto_LostFocus()
    If NumericVal(txtLto.Text) <= 0 Then txtLto.Text = "0.00"
    txtLto.Text = FormatNumber(NumericVal(txtLto.Text))
End Sub

Private Sub txtUnitPrice_GotFocus()
    If NumericVal(txtUnitPrice.Text) <= 0 Then txtUnitPrice = ""
End Sub

Private Sub txtUnitPrice_LostFocus()
    If NumericVal(txtUnitPrice) <= 0 Then txtUnitPrice = "0.00"
    txtUnitPrice = FormatNumber(NumericVal(txtUnitPrice))
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSRP_GotFocus()
    If NumericVal(txtSRP.Text) <= 0 Then txtSRP.Text = ""
End Sub

Private Sub txtSRP_LostFocus()
    If NumericVal(txtSRP.Text) <= 0 Then txtSRP.Text = "0.00"
    txtSRP.Text = FormatNumber(NumericVal(txtSRP.Text))
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid (txtSEARCH)

End Sub


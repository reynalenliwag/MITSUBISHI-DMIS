VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Trans_TestVehicles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Drive Vehicle"
   ClientHeight    =   7560
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   7605
   ForeColor       =   &H00FCFCFC&
   Icon            =   "TestVehicles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   7605
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   7605
      TabIndex        =   49
      Top             =   6570
      Width           =   7605
      Begin Crystal.CrystalReport rptMRR 
         Left            =   840
         Top             =   90
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   420
         Top             =   120
      End
      Begin VB.PictureBox picSave 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5880
         ScaleHeight     =   885
         ScaleWidth      =   1800
         TabIndex        =   60
         Top             =   30
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
            Left            =   960
            MouseIcon       =   "TestVehicles.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   61
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
            Left            =   270
            MouseIcon       =   "TestVehicles.frx":0D5A
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":0EAC
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Save this Record"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdd 
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   810
         ScaleHeight     =   945
         ScaleWidth      =   8655
         TabIndex        =   50
         Top             =   -30
         Width           =   8655
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
            Left            =   6030
            MouseIcon       =   "TestVehicles.frx":11FC
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":134E
            Style           =   1  'Graphical
            TabIndex        =   59
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
            Left            =   5340
            MouseIcon       =   "TestVehicles.frx":16B4
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":1806
            Style           =   1  'Graphical
            TabIndex        =   58
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
            Left            =   4650
            MouseIcon       =   "TestVehicles.frx":1B6C
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":1CBE
            Style           =   1  'Graphical
            TabIndex        =   57
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
            Left            =   3960
            MouseIcon       =   "TestVehicles.frx":1FE9
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":213B
            Style           =   1  'Graphical
            TabIndex        =   54
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
            Left            =   3270
            MouseIcon       =   "TestVehicles.frx":2497
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":25E9
            Style           =   1  'Graphical
            TabIndex        =   55
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
            Left            =   2580
            MouseIcon       =   "TestVehicles.frx":28FC
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":2A4E
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdView 
            Caption         =   "Test Drive Monitoring"
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
            Left            =   1830
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "TestVehicles.frx":2D48
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":2E9A
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Post this Transaction"
            Top             =   60
            Width           =   765
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
            Left            =   1140
            MouseIcon       =   "TestVehicles.frx":31BF
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":3311
            Style           =   1  'Graphical
            TabIndex        =   52
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
            Left            =   450
            MouseIcon       =   "TestVehicles.frx":3669
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":37BB
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox picSearch 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6570
      Left            =   0
      ScaleHeight     =   6570
      ScaleWidth      =   2625
      TabIndex        =   9
      Top             =   0
      Width           =   2625
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
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Text            =   "TEXT1"
         Top             =   960
         Width           =   2505
      End
      Begin VB.OptionButton optByCode 
         Caption         =   "By Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   90
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.OptionButton optByModel 
         Caption         =   "By Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   2235
      End
      Begin VB.OptionButton optByDescription 
         Caption         =   "By Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   12
         Top             =   630
         Width           =   2235
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   5025
         Left            =   90
         TabIndex        =   14
         Top             =   1500
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   8864
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
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picDataEntry 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6570
      Left            =   2625
      ScaleHeight     =   6570
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox CboSa 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1290
         Style           =   1  'Simple Combo
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   2190
         Width           =   3555
      End
      Begin VB.TextBox txttime 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3210
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   4710
         Width           =   1635
      End
      Begin VB.ComboBox CboClient 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1290
         Style           =   1  'Simple Combo
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   1770
         Width           =   3555
      End
      Begin VB.TextBox txtSource 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   24
         Tag             =   "@R"
         Top             =   1335
         Width           =   3540
      End
      Begin VB.TextBox txtIGNKeyNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1305
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   3450
         Width           =   3555
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2625
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   20
         Tag             =   "@R"
         Text            =   "Text1"
         Top             =   480
         Width           =   2250
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   1245
      End
      Begin VB.ComboBox cboDescript 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1305
         Style           =   1  'Simple Combo
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   900
         Width           =   3555
      End
      Begin VB.ComboBox cboColor 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1305
         Style           =   1  'Simple Combo
         TabIndex        =   30
         Text            =   "cboColor"
         Top             =   2625
         Width           =   3555
      End
      Begin MSComCtl2.DTPicker dtdaterecieved 
         Height          =   390
         Left            =   1305
         TabIndex        =   40
         Top             =   4305
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   688
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
         Format          =   54788097
         CurrentDate     =   39141
      End
      Begin VB.TextBox txtNotes 
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
         ForeColor       =   &H00701E2A&
         Height          =   675
         Left            =   1290
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   5160
         Width           =   3555
      End
      Begin VB.TextBox txtEngineNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1305
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   3045
         Width           =   3555
      End
      Begin VB.TextBox txtVINo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1305
         TabIndex        =   38
         Tag             =   "@R"
         Text            =   "Text1"
         Top             =   3870
         Width           =   3555
      End
      Begin VB.TextBox txtSerialNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4950
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   3690
         Visible         =   0   'False
         Width           =   3555
      End
      Begin MSComCtl2.DTPicker dtDateReturned 
         Height          =   390
         Left            =   1320
         TabIndex        =   42
         Top             =   4710
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   688
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
         Format          =   54788097
         CurrentDate     =   39141
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1290
         TabIndex        =   46
         Top             =   5820
         Width           =   3585
         Begin VB.CheckBox chkApproved 
            Caption         =   "Approved"
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
            Left            =   180
            TabIndex        =   48
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox ChkDisApp 
            Caption         =   "DisApproved"
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
            Left            =   1410
            TabIndex        =   47
            Top             =   270
            Width           =   1905
         End
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   420
         TabIndex        =   27
         Top             =   2250
         Width           =   765
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "CLient Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date and time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   41
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   105
         TabIndex        =   16
         Top             =   90
         Width           =   1965
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label28 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   90
         TabIndex        =   18
         Top             =   510
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   29
         Top             =   2655
         Width           =   885
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   945
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Arrived"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   39
         Top             =   4380
         Width           =   1065
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   5115
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "CS NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -300
         TabIndex        =   35
         Top             =   3450
         Width           =   1515
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   270
         TabIndex        =   32
         Top             =   3075
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "VI No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   37
         Top             =   3870
         Width           =   885
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5130
         TabIndex        =   31
         Top             =   2970
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.PictureBox piclist 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7785
      ScaleWidth      =   7485
      TabIndex        =   0
      Top             =   -30
      Visible         =   0   'False
      Width           =   7515
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   330
         Width           =   7395
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   660
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   240
            Width           =   3765
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SAE :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   270
            Width           =   405
         End
      End
      Begin wizButton.cmd cmdclose 
         Height          =   435
         Left            =   5790
         TabIndex        =   8
         Top             =   7200
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "TestVehicles.frx":3B1A
      End
      Begin MSComctlLib.ListView ListTestDrive 
         Height          =   5775
         Left            =   60
         TabIndex        =   6
         Top             =   1020
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   10186
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
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vehicle Model"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CS no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Color"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SAE"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ProspectID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Information:Double Click to Select "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   7110
         Width           =   3345
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "::Test Drive For Approval::"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   2610
         TabIndex        =   2
         Top             =   60
         Width           =   2955
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   -60
         TabIndex        =   1
         Top             =   0
         Width           =   7845
         _Version        =   655364
         _ExtentX        =   13838
         _ExtentY        =   609
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_TestVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theProspectID                                                     As Integer
Public TestDriveID                                                    As Long
Private RS                                                            As Recordset

Function GetProspectName(ByVal XXX As Integer)
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset

    SQL = "SELECT Acctname FROM CRIS_PROSPECTS WHERE ProspectID=" & XXX & ""

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.BOF And Not RS.EOF Then
        CboClient.Text = Null2String(RS!AcctName)
    End If
    Set RS = Nothing
End Function

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset
    lvSearch.Sorted = False: lvSearch.ListItems.Clear
    lvSearch.Enabled = False
    Set TEMPRS = New ADODB.Recordset

    If optByCode.Value = True Then
        Set TEMPRS = gconDMIS.Execute("select  Code, Descript, ID from CRIS_MRRINV where CODE like'" & ReplaceQuote(XXX) & "%' order by 1 asc")
    ElseIf optByModel.Value = True Then
        Set TEMPRS = gconDMIS.Execute("select  Code, Descript, ID from CRIS_MRRINV where Model like'" & ReplaceQuote(XXX) & "%' order by 1 asc")
    ElseIf optByDescription.Value = True Then
        Set TEMPRS = gconDMIS.Execute("select  Code, Descript, ID from CRIS_MRRINV where Descript like'" & ReplaceQuote(XXX) & "%' order by 1 asc")

    End If

    If Not (TEMPRS.EOF And TEMPRS.BOF) Then
        Listview_Loadval lvSearch.ListItems, TEMPRS
        lvSearch.Refresh
        lvSearch.Enabled = True
    End If
End Sub

Sub InitVars()
    Dim cntl                                                          As Control
    For Each cntl In Me.ControlS
        If TypeOf cntl Is TextBox Or TypeOf cntl Is ComboBox Then
            cntl = vbNullString
        End If
    Next
    dtdaterecieved.Value = LOGDATE
    dtDateReturned.Value = LOGDATE
    lblSTATUS = ""
    FillCombo "SELECT ID, Color_Desc FROM ALL_Color", 0, 1, cboColor
    FillCombo "Select ID,  Descript from All_Model where LEN(code)<> 0 order by descript asc", 0, 1, cboDescript
End Sub

Sub rsRefresh()

    Set RS = gconDMIS.Execute("SELECT * FROM CRIS_MRRINV order by id desc")

End Sub

Sub StoreMemVars()
    If Not (RS.EOF And RS.BOF) Then
        TestDriveID = RS!ID
        TXTCODE = Null2String(RS!CODE)
        cboDescript.ListIndex = SelectCombo(cboDescript, Null2String(RS!DESCRIPT))
        txtModel = Null2String(RS!Model)
        cboColor = Null2String(RS!Color)
        txtSerialNo = Null2String(RS!SERIALNO)
        txtVINO = Null2String(RS!VINNUMBER)
        txtEngineNo = Null2String(RS!ENGINENUMBER)
        txtIGNKeyNo = Null2String(RS!IGNKEYNO)
        dtdaterecieved.Value = Null2String(RS!datereceived)
        txttime = Null2String(RS!Time)
        CboClient = Null2String(RS!Clientname)
        cboSA = Null2String(RS!saname)
        dtDateReturned.Value = Null2String(RS!DateReturned)

        txtNotes = Null2String(RS!Notes)
        txtSource = Null2String(RS!Source)
        If Null2String(RS!HITCOUNTER) <> "" Then
            cmdDelete.Enabled = False
        Else
            cmdDelete.Enabled = True
        End If
        If IsNull(RS!DateReturned) = True Then
            'lblStatus.Caption = "***Available***"
        Else
            'lblStatus.Caption = "***Returned***"
        End If

        If Null2String(RS!STATUS) = "Approved" Then
            chkApproved.Value = 1
            ChkDisApp.Value = 0
        Else
            ChkDisApp.Value = 1
            chkApproved.Value = 0
        End If


    Else
        ShowNoRecord
        piclist.Visible = False
        'cmdAdd.Value = True
    End If

End Sub

Sub FillClientName()
    Dim RsClient                                                      As New ADODB.Recordset
    Dim SQL                                                           As String

    SQL = "SELECT AcctName FROM CRIS_Prospects"

    Set RsClient = New ADODB.Recordset
    Set RsClient = gconDMIS.Execute(SQL)

    CboClient.Clear

    Do While Not RsClient.EOF
        CboClient.AddItem Null2String(RsClient!AcctName)
        RsClient.MoveNext
    Loop
    Set RsClient = Nothing
End Sub

Sub FillSAE()
    Dim RsSAE                                                         As New ADODB.Recordset
    Dim SQL                                                           As String

    SQL = "SELECT lname,fname,middle FROM Smis_SalesTeam"

    Set RsSAE = New ADODB.Recordset
    Set RsSAE = gconDMIS.Execute(SQL)

    cboSA.Clear

    Do While Not RsSAE.EOF
        cboSA.AddItem Null2String(RsSAE!lname) + "," + Null2String(RsSAE!fname) + "," + Null2String(RsSAE!MIDDLE)
        RsSAE.MoveNext
    Loop
    Set RsSAE = Nothing

End Sub

Sub LoadListTestDrive()
    On Error Resume Next

    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim cnt                                                           As Integer
    Dim Item                                                          As ListItem
    Dim xstatus                                                       As String

    xstatus = "For Approval"


    SQL = "SELECT prospectID,vehiclemodel,vehiclecode,color,status,SAE from CRIS_TestdriveSchedules where status='" & xstatus & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListTestDrive.ListItems.Clear

    cnt = 0
    '    rs.MoveFirst '***
    Do While Not RS.EOF
        cnt = cnt + 1
        Set Item = ListTestDrive.ListItems.Add(, , cnt)
        Item.SubItems(1) = Null2String(RS!vehiclemodel)
        Item.SubItems(2) = Null2String(RS!vehiclecode)
        Item.SubItems(3) = Null2String(RS!Color)
        Item.SubItems(4) = Null2String(RS!STATUS)
        Item.SubItems(5) = Null2String(RS!SAE)
        Item.SubItems(6) = Null2String(RS!PROSPECTID)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub UpdateTheLogTestDrive(ByVal XXX As Integer)
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim xstatus                                                       As String

    If chkApproved.Value = 1 Then
        xstatus = "Approved"
    End If

    If ChkDisApp.Value = 1 Then
        xstatus = "DisApproved"
    End If

    gconDMIS.Execute "UPDATE  CRIS_TestDriveSchedules set Status='" & xstatus & "' where ProspectID='" & XXX & "'"

End Sub

Private Sub cboDescript_Change()
    If cboDescript.ListIndex = -1 Then Exit Sub
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT * FROM ALL_MODEL WHERE ID=" & cboDescript.ItemData(cboDescript.ListIndex))
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        TXTCODE = Null2String(TEMPRS!CODE)
        txtModel = Null2String(TEMPRS!Model)
    End If
End Sub

Private Sub cboDescript_Click()
    cboDescript_Change
End Sub

Private Sub chkApproved_Click()
    If chkApproved.Value = 1 Then
        ChkDisApp.Enabled = False
    End If

    If chkApproved.Value = 0 Then
        ChkDisApp.Enabled = True
    End If


End Sub

Private Sub ChkDisApp_Click()
    If ChkDisApp.Value = 1 Then
        chkApproved.Enabled = False
    End If

    If ChkDisApp.Value = 0 Then
        chkApproved.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "TEST DRIVE VEHICLES") = False Then Exit Sub
    On Error GoTo ErrorCode:

    TestDriveID = 0
    InitVars
    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    picAdd.Visible = False
    picSave.Visible = True

    dtDateReturned.Enabled = False
    On Error Resume Next
    ' cboDescript.SetFocus
    LoadListTestDrive
    piclist.Visible = True
    piclist.ZOrder 0


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picDataEntry.Enabled = False
    PICSEARCH.Enabled = True
    picAdd.Visible = True
    picSave.Visible = False

    txtSEARCH.Enabled = True
    optByCode.Enabled = True
    optByModel.Enabled = True
    optByDescription.Enabled = True

    StoreMemVars
End Sub

Private Sub cmdClose_Click()
    piclist.Visible = False
End Sub

'Upating Code       : AXP-0707200713:30
Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "TEST DRIVE VEHICLES") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If gconDMIS.Execute("SELECT Count(*) From CRIS_TestDriveSchedules Where VehicleCode=" & N2Str2Null(TXTCODE)).Fields(0).Value > 0 Then
        MessagePop RecLocekd, "Record In Use", "Current Test Drive Information Has Been In Use .... Cannot Delete The Record"

        Exit Sub
    End If

    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "delete from CRIS_MRRINV where id = " & TestDriveID

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "X", "TEST DRIVE VEHICLES", SQL_STATEMENT, Null2String(TestDriveID), "", "Code :" & TXTCODE, "", ""

        ShowDeletedMsg
        FillSearchGrid txtSEARCH
        rsRefresh
        StoreMemVars
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "TEST DRIVE VEHICLES") = False Then Exit Sub
    On Error GoTo ErrorCode:

    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    picAdd.Visible = False
    picSave.Visible = True
    dtDateReturned.Enabled = True

    txtSEARCH.Enabled = False
    optByCode.Enabled = False
    optByModel.Enabled = False
    optByDescription.Enabled = False


    On Error Resume Next
    'cboDescript.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPost_Click()

End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "TEST DRIVE VEHICLES") = False Then Exit Sub
    On Error GoTo ErrorCode:
    Screen.MousePointer = 11




    rptMRR.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptMRR.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    PrintSQLReport rptMRR, SMIS_REPORT_PATH & "TestDrive.rpt", "{td.ID} = " & TestDriveID, DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:

End Sub

'Upating Code       : AXP-0707200713:30
Private Sub cmdSave_Click()
    Dim theStatus                                                     As String

    'On Error GoTo Errorcode:

    If TXTCODE = "" Then
        MsgBox "Missing Model Code..Please Advice to check to Model Master File", vbInformation, "Information"
        'ShowIsRequiredMsg "Vehicle Model Code"
        On Error Resume Next
        '        TXTCODE.SetFocus
        Exit Sub
    End If

    '    If cboDescript.ListIndex = -1 Then
    '        ShowIsRequiredMsg "Vehicle Model "
    '        On Error Resume Next
    '        cboDescript.SetFocus
    '        Exit Sub
    '    End If


    If cboColor = "" Then
        ShowIsRequiredMsg "Model Color"
        On Error Resume Next
        '        cboColor.SetFocus
        Exit Sub
    End If

    If txtVINO = "" Or txtEngineNo = "" Then
        ShowIsRequiredMsg " Vin Number and Engine Number"
        On Error Resume Next
        txtVINO.SetFocus
        Exit Sub
    End If


    If txtIGNKeyNo = "" Then
        ShowIsRequiredMsg "Conduction Sticker Number"
        On Error Resume Next
        txtIGNKeyNo.SetFocus
        Exit Sub
    End If

    If chkApproved.Value = 1 Then
        theStatus = "Approved"
    Else
        theStatus = "DisApproved"
    End If

    If chkApproved.Value = 0 And ChkDisApp.Value = 0 Then
        'MsgBox "You Forgot to put A status..", vbInformation, "Please select a status"
        MsgBox "Please Indicate Status..", vbInformation, "Status"
        Exit Sub
    End If

    Dim CODE                                                          As String
    Dim DESCRIPT                                                      As String
    Dim Source                                                        As String
    Dim Color                                                         As String
    Dim ENGINENUMBER                                                  As String
    Dim SERIALNO                                                      As String
    Dim VINNUMBER                                                     As String
    Dim TAGGEDPRICE                                                   As String
    Dim Notes                                                         As String
    Dim DATERECIEVED                                                  As String
    Dim Model, SQL                                                    As String
    Dim IGNKEYNO                                                      As String
    Dim theClint                                                      As String

    CODE = N2Str2Null(TXTCODE)
    DESCRIPT = N2Str2Null(cboDescript)
    Color = N2Str2Null(cboColor)
    ENGINENUMBER = N2Str2Null(txtEngineNo)
    SERIALNO = N2Str2Null(txtSerialNo)
    VINNUMBER = N2Str2Null(txtVINO)
    Notes = N2Str2Null(txtNotes)
    DATERECIEVED = Format(LOGDATE, "MM/dd/yyyy")
    Source = N2Str2Null(txtSource)
    Model = N2Str2Null(txtModel)
    IGNKEYNO = N2Str2Null(txtIGNKeyNo)
    theClint = CboClient.Text

    Dim TEMPRS                                                        As ADODB.Recordset

    If TestDriveID = 0 Then
        SQL_STATEMENT = " INSERT INTO  " & _
                      "  CRIS_MRRINV(Code, Descript, Model, Source, Color, EngineNumber, SerialNo, VinNumber, IGNKEYNO, Notes, DateReceived, clientname, saname, time, Status, hitcounter) " & _
                      "  VALUES( " & CODE & " , " & DESCRIPT & " , " & Model & " , " & Source & "," & Color & _
                      " , " & ENGINENUMBER & " , " & SERIALNO & " ," & VINNUMBER & "," & IGNKEYNO & " ," & Notes & " ,'" & DATERECIEVED & "','" & theClint & "','" & cboSA.Text & "','" & txttime.Text & "','" & theStatus & "' , 0) " & vbCrLf & " SELECT @@IDENTITY "

        gconDMIS.Execute (SQL_STATEMENT)

        NEW_LogAudit "A", "TEST DRIVE VEHICLES", SQL_STATEMENT, FindTransactionID(N2Str2Null(TXTCODE), "CODE", "CRIS_MRRINV"), "", "Code :" & TXTCODE, "", ""
    Else

        SQL_STATEMENT = " UPDATE CRIS_MRRINV " & _
                      " SET Code = " & CODE & " ," & _
                      " Descript= " & DESCRIPT & "," & _
                      " Model= " & Model & " ," & _
                      " Source= " & Source & " ," & _
                      " Color=" & Color & " ," & _
                      " EngineNumber= " & ENGINENUMBER & " ," & _
                      " SerialNo= " & SERIALNO & " ," & _
                      " VinNumber= " & VINNUMBER & "," & _
                      " IGNKEYNO=  " & IGNKEYNO & "," & _
                      " Notes= " & Notes & " , " & _
                      " ClientName= '" & CboClient.Text & "' , " & _
                      " saname= '" & cboSA.Text & "' , " & _
                      " time= '" & txttime & "' , " & _
                      " Status ='" & theStatus & "'," & _
                      " DateReceived = " & DATERECIEVED & _
                      " WHERE id= " & TestDriveID

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "A", "TEST DRIVE VEHICLES", SQL_STATEMENT, Null2String(TestDriveID), "", "Code :" & TXTCODE, "", ""

        If IsNull(dtDateReturned.Value) = False Then
            If MsgBox("Do you Want to Return This Vehicle. ", vbOKCancel + vbExclamation, "Confirm Posting") = vbOK Then
                SQL_STATEMENT = "update cris_mrrinv set DateReturned= " & N2Date2Null(dtDateReturned.Value) & " where id =" & TestDriveID

                gconDMIS.Execute (SQL_STATEMENT)
                NEW_LogAudit "E", "TEST DRIVE VEHICLES", SQL_STATEMENT, Null2String(TestDriveID), "", "Code :" & TXTCODE, "", ""

                MessagePop RecSaveOk, "Returned", "Test Drive Vehicle Returned"
            End If
        Else
            SQL_STATEMENT = "update cris_mrrinv set DateReturned= NULL where id =" & TestDriveID
            gconDMIS.Execute (SQL_STATEMENT)
            NEW_LogAudit "E", "TEST DRIVE VEHICLES", SQL_STATEMENT, Null2String(TestDriveID), "", "Code :" & TXTCODE, "", ""
        End If
    End If


    Set TEMPRS = New ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute(SQL_STATEMENT)
    If TestDriveID <= 0 Then
        MessagePop RecSave, "Record Added ", "New Test Vehicles Sucessfully "
    Else
        MessagePop RecSave, "Record Saved", "Test Vehicles  Sucessfully Updated"
    End If
    Set TEMPRS = TEMPRS.NextRecordset
    If Not TEMPRS Is Nothing Then
        TestDriveID = TEMPRS.Collect(0)
    End If
    Set TEMPRS = Nothing

    RS.Requery
    RS.Find ("ID=" & TestDriveID)
    cmdCancel.Value = True
    FillSearchGrid txtSEARCH
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub CmdView_Click()
    If Module_Access(LOGID, "TEST DRIVE VEHICLES", "INQUIRY") = False Then Exit Sub
    frmFile_TestDriveMonitoring.Show 1
End Sub

Private Sub dtdaterecieved_Change()
    'dtdaterecieved.Value = Format(Now, "MM/dd/yyyy")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TEST DRIVE VEHICLES)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(TestDriveID), "TEST DRIVE VEHICLES")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Call AddColumnHeader("Model,Description", lvSearch)
    Call ResizeColumnHeader(lvSearch, "25,70")
    InitVars
    rsRefresh

    picAdd.Visible = True
    picSave.Visible = False
    picDataEntry.Enabled = False

    StoreMemVars
    FillSearchGrid txtSEARCH
    'FillClientName
    'FillSAE
    LoadListTestDrive
    piclist.Visible = False
End Sub

Private Sub ListTestDrive_DblClick()
    On Error Resume Next

    Dim TEMPRS                                                        As ADODB.Recordset

    Dim SQL                                                           As String
    cboDescript.Text = ListTestDrive.SelectedItem.SubItems(1)

    cboColor.Text = ListTestDrive.SelectedItem.SubItems(3)
    cboSA.Text = ListTestDrive.SelectedItem.SubItems(5)
    theProspectID = ListTestDrive.SelectedItem.SubItems(6)

    piclist.Visible = False

    SQL = "SELECT * FROM ALL_MODEL WHERE descript= '" & cboDescript.Text & "'"
    Set TEMPRS = New ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute(SQL)

    If Not (TEMPRS.EOF And TEMPRS.BOF) Then
        TXTCODE = Null2String(TEMPRS!CODE)

        txtModel = Null2String(TEMPRS!Model)
    End If

    GetProspectName theProspectID

    chkApproved.Value = 0
    ChkDisApp.Value = 0
End Sub

Private Sub lvSearch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvSearch
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

Private Sub lvSearch_DblClick()
    If lvSearch.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find ("ID= " & Item.ListSubItems(lvSearch.ColumnHeaders.Count))
    StoreMemVars
End Sub

Private Sub optByCode_Click()
    FillSearchGrid txtSEARCH
End Sub

Private Sub optByDescription_Click()
    FillSearchGrid txtSEARCH
End Sub

Private Sub optByModel_Click()
    FillSearchGrid txtSEARCH
End Sub

Private Sub Text1_Change()

    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim cnt                                                           As Integer
    Dim Item                                                          As ListItem
    Dim Keyword                                                       As String

    If Text1.Text = "" Then
        LoadListTestDrive
        Exit Sub
    End If

    Keyword = Trim(Text1.Text)

    SQL = "SELECT prospectID,vehiclemodel,vehiclecode,color,status,SAE from CRIS_TestdriveSchedules where"

    If Len(Keyword) = 0 Then Exit Sub

    SQL = SQL & " SAE LIKE  '" & Keyword & "%'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListTestDrive.ListItems.Clear

    cnt = 0

    Do While Not RS.EOF
        cnt = cnt + 1
        Set Item = ListTestDrive.ListItems.Add(, , cnt)
        Item.SubItems(1) = Null2String(RS!vehiclemodel)
        Item.SubItems(2) = Null2String(RS!vehiclecode)
        Item.SubItems(3) = Null2String(RS!Color)
        Item.SubItems(4) = Null2String(RS!STATUS)
        Item.SubItems(5) = Null2String(RS!SAE)
        Item.SubItems(6) = Null2String(RS!PROSPECTID)
        RS.MoveNext
    Loop
    Set RS = Nothing


End Sub

Private Sub Timer1_Timer()
    '    If lblStatus.Caption <> "" Then
    '        If lblStatus.Visible = True Then
    '            lblStatus.Visible = False
    '        Else
    '            lblStatus.Visible = True
    '        End If
    '    End If
End Sub

Private Sub txtEngineNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)

End Sub

Private Sub txtIGNKeyNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid txtSEARCH
End Sub

Private Sub txtSerialNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSource_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtVINo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub


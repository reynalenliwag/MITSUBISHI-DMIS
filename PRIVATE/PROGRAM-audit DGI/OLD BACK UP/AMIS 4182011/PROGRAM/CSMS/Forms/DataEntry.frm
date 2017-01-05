VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMSDataEntry 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repair Order Data Entry"
   ClientHeight    =   7470
   ClientLeft      =   4605
   ClientTop       =   435
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "DataEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   13140
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8865
      Left            =   0
      ScaleHeight     =   8835
      ScaleWidth      =   13065
      TabIndex        =   26
      Top             =   30
      Visible         =   0   'False
      Width           =   13095
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
         Height          =   795
         Left            =   5460
         ScaleHeight     =   795
         ScaleWidth      =   4845
         TabIndex        =   222
         Top             =   6000
         Width           =   4845
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            CausesValidation=   0   'False
            Height          =   795
            Left            =   4140
            MouseIcon       =   "DataEntry.frx":1082
            MousePointer    =   99  'Custom
            Picture         =   "DataEntry.frx":11D4
            Style           =   1  'Graphical
            TabIndex        =   223
            ToolTipText     =   "Exit Window"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   3450
            MouseIcon       =   "DataEntry.frx":153A
            MousePointer    =   99  'Custom
            Picture         =   "DataEntry.frx":168C
            Style           =   1  'Graphical
            TabIndex        =   224
            ToolTipText     =   "Print this Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdReleaseRO 
            Caption         =   "Release"
            Height          =   795
            Left            =   2760
            MouseIcon       =   "DataEntry.frx":19F2
            MousePointer    =   99  'Custom
            Picture         =   "DataEntry.frx":1B44
            Style           =   1  'Graphical
            TabIndex        =   225
            ToolTipText     =   "Release Vehicle"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   2070
            MouseIcon       =   "DataEntry.frx":42E6
            MousePointer    =   99  'Custom
            Picture         =   "DataEntry.frx":4438
            Style           =   1  'Graphical
            TabIndex        =   226
            ToolTipText     =   "Edit Selected Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1380
            MouseIcon       =   "DataEntry.frx":4794
            MousePointer    =   99  'Custom
            Picture         =   "DataEntry.frx":48E6
            Style           =   1  'Graphical
            TabIndex        =   227
            ToolTipText     =   "Find a Record"
            Top             =   0
            Width           =   705
         End
      End
      Begin Crystal.CrystalReport rptHPC 
         Left            =   960
         Top             =   6000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   2115
         Left            =   8730
         ScaleHeight     =   2085
         ScaleWidth      =   4275
         TabIndex        =   277
         Top             =   7110
         Width           =   4305
         Begin XtremeShortcutBar.ShortcutCaption capInfo 
            Height          =   255
            Left            =   0
            TabIndex        =   281
            Top             =   0
            Width           =   4275
            _Version        =   655364
            _ExtentX        =   7541
            _ExtentY        =   450
            _StockProps     =   14
            Caption         =   "Repair Order Info"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
         Begin VB.Label labInfo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   1875
            Left            =   -30
            TabIndex        =   278
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Left            =   4380
         ScaleHeight     =   2085
         ScaleWidth      =   4305
         TabIndex        =   29
         Top             =   7110
         Width           =   4335
         Begin XtremeShortcutBar.ShortcutCaption capSUG 
            Height          =   255
            Left            =   0
            TabIndex        =   279
            Top             =   0
            Width           =   4305
            _Version        =   655364
            _ExtentX        =   7594
            _ExtentY        =   450
            _StockProps     =   14
            Caption         =   "Suggestion/ Recommendation:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
         Begin VB.Label labNotes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   1875
            Left            =   -15
            TabIndex        =   30
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.PictureBox pic3 
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
         Height          =   3465
         Left            =   10320
         ScaleHeight     =   3435
         ScaleWidth      =   2655
         TabIndex        =   40
         Top             =   30
         Width           =   2685
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F8 - Refresh Record"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   90
            TabIndex        =   285
            Top             =   906
            Width           =   2535
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
            Height          =   315
            Left            =   0
            TabIndex        =   276
            Top             =   0
            Width           =   2655
            _Version        =   655364
            _ExtentX        =   4683
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "RO Options"
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
            VisualTheme     =   3
            Alignment       =   1
         End
         Begin VB.Label labF7 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F7 - Input Discount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   90
            TabIndex        =   45
            Top             =   405
            Width           =   2535
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F11 - View CHG Limit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   90
            MouseIcon       =   "DataEntry.frx":4BE0
            MousePointer    =   99  'Custom
            TabIndex        =   44
            Top             =   2409
            Width           =   2535
         End
         Begin VB.Label Label79 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F12 - Edit Vehicle Info"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   90
            MouseIcon       =   "DataEntry.frx":4EEA
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   2910
            Width           =   2535
         End
         Begin VB.Label labF9 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F9 - Cancel Invoice"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   90
            TabIndex        =   42
            Top             =   1410
            Width           =   2535
         End
         Begin VB.Label labF10 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F10 - Unreleased"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   90
            TabIndex        =   41
            Top             =   1908
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture5 
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
         Height          =   3315
         Left            =   10320
         ScaleHeight     =   3285
         ScaleWidth      =   2655
         TabIndex        =   33
         Top             =   3480
         Width           =   2685
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Set to WARRANTY RO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "DataEntry.frx":51F4
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1560
            Width           =   90
         End
         Begin VB.CommandButton cmdNoCharge 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Set to NO CHARGE RO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "DataEntry.frx":5346
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   420
            Width           =   2445
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Set to PDI RO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "DataEntry.frx":5498
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   982
            Width           =   2445
         End
         Begin VB.CommandButton cmdROVatExempt 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Set to Zero Rated"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "DataEntry.frx":55EA
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2106
            Width           =   2445
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Force Delete Opened RO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   120
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "DataEntry.frx":573C
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   2670
            Width           =   2445
         End
         Begin VB.CommandButton cmdInternalRO 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Set to INTERNAL RO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "DataEntry.frx":588E
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1544
            Width           =   2445
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   345
            Left            =   0
            TabIndex        =   275
            Top             =   0
            Width           =   2655
            _Version        =   655364
            _ExtentX        =   4683
            _ExtentY        =   609
            _StockProps     =   14
            Caption         =   "RO Advance Options"
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
            VisualTheme     =   3
            Alignment       =   1
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
         Height          =   795
         Left            =   8730
         ScaleHeight     =   795
         ScaleWidth      =   1590
         TabIndex        =   228
         Top             =   6000
         Visible         =   0   'False
         Width           =   1590
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            CausesValidation=   0   'False
            Height          =   795
            Left            =   870
            MouseIcon       =   "DataEntry.frx":59E0
            MousePointer    =   99  'Custom
            Picture         =   "DataEntry.frx":5B32
            Style           =   1  'Graphical
            TabIndex        =   229
            ToolTipText     =   "Cancel"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   180
            MouseIcon       =   "DataEntry.frx":5E70
            MousePointer    =   99  'Custom
            Picture         =   "DataEntry.frx":5FC2
            Style           =   1  'Graphical
            TabIndex        =   230
            ToolTipText     =   "Save this Record"
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.PictureBox picFollowUpResult 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         Left            =   30
         ScaleHeight     =   2085
         ScaleWidth      =   4305
         TabIndex        =   27
         Top             =   7110
         Width           =   4335
         Begin XtremeShortcutBar.ShortcutCaption capFollow 
            Height          =   255
            Left            =   0
            TabIndex        =   280
            Top             =   0
            Width           =   4305
            _Version        =   655364
            _ExtentX        =   7594
            _ExtentY        =   450
            _StockProps     =   14
            Caption         =   "Notes after follow up "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
         Begin VB.Label labCalled_Result 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   0
            TabIndex        =   28
            Top             =   240
            Width           =   4320
         End
      End
      Begin VB.PictureBox Picture7 
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
         Height          =   270
         Left            =   10110
         ScaleHeight     =   240
         ScaleWidth      =   2895
         TabIndex        =   31
         Top             =   6810
         Width           =   2925
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   330
            Top             =   -450
         End
         Begin VB.Label labZeroRated 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BackStyle       =   0  'Transparent
            Caption         =   "ZERO RATED TRANSACTION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   60
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   2685
         End
      End
      Begin Crystal.CrystalReport rptRepairOrder 
         Left            =   60
         Top             =   6000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Repair Order Print Out"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin Crystal.CrystalReport rptDET 
         Left            =   510
         Top             =   6000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Repair Order Print Out"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   30
         TabIndex        =   48
         Top             =   -60
         Width           =   10275
         Begin VB.TextBox txttype 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   8400
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   318
            Top             =   2430
            Width           =   1035
         End
         Begin VB.ComboBox Cbo_Rotype 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   8400
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   319
            Top             =   2430
            Width           =   1065
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Change SA"
            Height          =   315
            Left            =   7170
            TabIndex        =   311
            Top             =   1710
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSComCtl2.DTPicker txtDte_recd 
            Height          =   315
            Left            =   1380
            TabIndex        =   310
            Top             =   2070
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   53149697
            CurrentDate     =   40102
         End
         Begin VB.TextBox cboModel 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   2760
            Width           =   1755
         End
         Begin VB.TextBox txtInvoiceNo 
            Alignment       =   2  'Center
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
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   5130
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   210
            Width           =   1365
         End
         Begin VB.CommandButton cmdCust 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2430
            TabIndex        =   6
            Top             =   570
            Width           =   345
         End
         Begin VB.TextBox txtAddress 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   930
            Width           =   8805
         End
         Begin VB.TextBox txtMake 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4740
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   2790
            Width           =   5445
         End
         Begin VB.TextBox txtCertific8 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6630
            MaxLength       =   9
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   210
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtParticipat 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   9
            Top             =   1290
            Width           =   1035
         End
         Begin VB.TextBox txtSektion 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4830
            MaxLength       =   3
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox txtKm_rdg 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            MaxLength       =   9
            TabIndex        =   13
            Top             =   1710
            Width           =   1755
         End
         Begin VB.TextBox txtTerm 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Left            =   8790
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   4
            Top             =   210
            Width           =   1365
         End
         Begin VB.TextBox txtSvc_No 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1050
            MaxLength       =   1
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   4080
            Width           =   1815
         End
         Begin VB.TextBox txtROType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1050
            MaxLength       =   1
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   3720
            Width           =   1815
         End
         Begin VB.TextBox txtAcct_No 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   5
            Top             =   570
            Width           =   1035
         End
         Begin VB.TextBox txtPlate_No 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   17
            Top             =   2400
            Width           =   1755
         End
         Begin VB.TextBox txtEstimateno 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   -480
            Width           =   1815
         End
         Begin VB.TextBox txtRep_Or 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "R-99999999"
            Top             =   210
            Width           =   1035
         End
         Begin VB.ComboBox cboRecd_by 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   4740
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1710
            Width           =   2385
         End
         Begin VB.CheckBox chkParticipat 
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
            ForeColor       =   &H80000007&
            Height          =   285
            Left            =   9930
            TabIndex        =   12
            Top             =   1320
            Width           =   225
         End
         Begin VB.TextBox txtDte_comp 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4740
            MaxLength       =   10
            TabIndex        =   15
            Top             =   2070
            Width           =   2385
         End
         Begin VB.TextBox txtDte_Rel 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   8400
            MaxLength       =   10
            TabIndex        =   16
            Top             =   2070
            Width           =   1785
         End
         Begin VB.TextBox txtVIN 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4740
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   18
            Top             =   2430
            Width           =   2385
         End
         Begin VB.TextBox txtParticipation 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1290
            Width           =   7155
         End
         Begin VB.CommandButton cmdPart 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2430
            TabIndex        =   10
            Top             =   1290
            Width           =   315
         End
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2430
            TabIndex        =   2
            Top             =   240
            Width           =   345
         End
         Begin VB.TextBox txtNiym 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   570
            Width           =   7425
         End
         Begin VB.Label lbl_rotype 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "RO Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7680
            TabIndex        =   320
            Top             =   2520
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Model Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   7
            Left            =   3195
            TabIndex        =   74
            Top             =   2850
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   6
            Left            =   765
            TabIndex        =   73
            Top             =   2850
            Width           =   510
         End
         Begin VB.Label Label17 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "    Warranty Certificate Number"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   7440
            TabIndex        =   72
            Top             =   3390
            Width           =   2865
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "VIN NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   42
            Left            =   4170
            TabIndex        =   71
            Top             =   2550
            Width           =   540
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Participation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   270
            TabIndex        =   70
            Top             =   1350
            Width           =   1020
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Released"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   10
            Left            =   7200
            TabIndex        =   69
            Top             =   2160
            Width           =   1170
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   41
            Left            =   3390
            TabIndex        =   68
            Top             =   2190
            Width           =   1320
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Recorded"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   38
            Left            =   90
            TabIndex        =   67
            Top             =   2130
            Width           =   1200
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Service Advisor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   37
            Left            =   3405
            TabIndex        =   66
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label10 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Section No."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   3780
            TabIndex        =   65
            Top             =   3390
            Width           =   1305
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "KM Reading"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   4
            Left            =   315
            TabIndex        =   64
            Top             =   1800
            Width           =   960
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Term"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   3
            Left            =   7890
            TabIndex        =   63
            Top             =   270
            Width           =   780
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Service"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   330
            TabIndex        =   62
            Top             =   4110
            Width           =   1035
         End
         Begin VB.Label Label6 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ROType"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   300
            TabIndex        =   61
            Top             =   3720
            Width           =   1035
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Plate No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   5
            Left            =   615
            TabIndex        =   60
            Top             =   2460
            Width           =   660
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Estimate No."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   90
            TabIndex        =   59
            Top             =   -450
            Width           =   1485
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   8
            Left            =   525
            TabIndex        =   58
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "RO Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   9
            Left            =   345
            TabIndex        =   57
            Top             =   300
            Width           =   930
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   56
            Top             =   660
            Width           =   840
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE NO."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   2
            Left            =   4080
            TabIndex        =   55
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label32 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "F12"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   7320
            TabIndex        =   54
            Top             =   3660
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   30
         TabIndex        =   46
         Top             =   3090
         Width           =   10275
         Begin TabDlg.SSTab SSTab1x 
            Height          =   1290
            Left            =   3990
            TabIndex        =   47
            Top             =   4110
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   2275
            _Version        =   393216
            Tabs            =   5
            Tab             =   4
            TabsPerRow      =   5
            TabHeight       =   1058
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   " Details"
            TabPicture(0)   =   "DataEntry.frx":6312
            Tab(0).ControlEnabled=   0   'False
            Tab(0).ControlCount=   0
            TabCaption(1)   =   " F3-Jobs"
            TabPicture(1)   =   "DataEntry.frx":6634
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   " F4-Parts"
            TabPicture(2)   =   "DataEntry.frx":694E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            TabCaption(3)   =   " F5-Materials"
            TabPicture(3)   =   "DataEntry.frx":6C68
            Tab(3).ControlEnabled=   0   'False
            Tab(3).ControlCount=   0
            TabCaption(4)   =   " F6-Accessories"
            TabPicture(4)   =   "DataEntry.frx":6F82
            Tab(4).ControlEnabled=   -1  'True
            Tab(4).ControlCount=   0
         End
         Begin XtremeSuiteControls.TabControl SSTab1 
            Height          =   2685
            Left            =   120
            TabIndex        =   263
            Top             =   120
            Width           =   10155
            _Version        =   655364
            _ExtentX        =   17912
            _ExtentY        =   4736
            _StockProps     =   64
            Appearance      =   2
            Color           =   4
            PaintManager.Layout=   2
            PaintManager.BoldSelected=   -1  'True
            PaintManager.HotTracking=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.LargeIcons=   -1  'True
            PaintManager.FixedTabWidth=   134
            ItemCount       =   5
            SelectedItem    =   4
            Item(0).Caption =   "Details"
            Item(0).ControlCount=   1
            Item(0).Control(0)=   "fraDetails"
            Item(1).Caption =   "F3 - Jobs"
            Item(1).ControlCount=   1
            Item(1).Control(0)=   "fraJobs"
            Item(2).Caption =   "F4 - Parts"
            Item(2).ControlCount=   1
            Item(2).Control(0)=   "fraParts"
            Item(3).Caption =   "F5 - Materials"
            Item(3).ControlCount=   1
            Item(3).Control(0)=   "fraMaterials"
            Item(4).Caption =   "F6 - Accessories"
            Item(4).ControlCount=   1
            Item(4).Control(0)=   "frmAccessories"
            Begin VB.Frame frmAccessories 
               BackColor       =   &H00404040&
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
               Height          =   2025
               Left            =   60
               TabIndex        =   273
               Top             =   600
               Width           =   10035
               Begin MSComctlLib.ListView lstAccessories 
                  Height          =   1995
                  Left            =   0
                  TabIndex        =   274
                  Top             =   0
                  Width           =   10005
                  _ExtentX        =   17648
                  _ExtentY        =   3519
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  HoverSelection  =   -1  'True
                  _Version        =   393217
                  ForeColor       =   0
                  BackColor       =   16777215
                  Appearance      =   1
                  MousePointer    =   99
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "DataEntry.frx":9734
                  NumItems        =   9
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "LINE #"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "ACC. CODE"
                     Object.Width           =   2646
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "DESCRIPTION"
                     Object.Width           =   3528
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   3
                     Text            =   "QTY"
                     Object.Width           =   1147
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   4
                     Text            =   "UNITPRICE"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   5
                     Text            =   "AMOUNT"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   6
                     Text            =   "WSC"
                     Object.Width           =   1147
                  EndProperty
                  BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   7
                     Text            =   "DISC."
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   8
                     Text            =   "ID"
                     Object.Width           =   2
                  EndProperty
               End
            End
            Begin VB.Frame fraMaterials 
               BackColor       =   &H00404040&
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
               Height          =   2025
               Left            =   -69940
               TabIndex        =   271
               Top             =   600
               Visible         =   0   'False
               Width           =   10035
               Begin MSComctlLib.ListView lstMaterials 
                  Height          =   1995
                  Left            =   0
                  TabIndex        =   272
                  Top             =   0
                  Width           =   10005
                  _ExtentX        =   17648
                  _ExtentY        =   3519
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  HoverSelection  =   -1  'True
                  _Version        =   393217
                  ForeColor       =   0
                  BackColor       =   16777215
                  Appearance      =   1
                  MousePointer    =   99
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "DataEntry.frx":9896
                  NumItems        =   9
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "LINE #"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "MAT. CODE"
                     Object.Width           =   2646
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "DESCRIPTION"
                     Object.Width           =   3528
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   3
                     Text            =   "QTY"
                     Object.Width           =   1147
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   4
                     Text            =   "UNITPRICE"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   5
                     Text            =   "AMOUNT"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   6
                     Text            =   "WSC"
                     Object.Width           =   1147
                  EndProperty
                  BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   7
                     Text            =   "DISC."
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   8
                     Text            =   "ID"
                     Object.Width           =   2
                  EndProperty
               End
            End
            Begin VB.Frame fraParts 
               BackColor       =   &H00404040&
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
               Height          =   2025
               Left            =   -69940
               TabIndex        =   269
               Top             =   600
               Visible         =   0   'False
               Width           =   10035
               Begin MSComctlLib.ListView lstParts 
                  Height          =   1995
                  Left            =   0
                  TabIndex        =   270
                  Top             =   0
                  Width           =   10005
                  _ExtentX        =   17648
                  _ExtentY        =   3519
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  HoverSelection  =   -1  'True
                  _Version        =   393217
                  ForeColor       =   0
                  BackColor       =   16777215
                  Appearance      =   1
                  MousePointer    =   99
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "DataEntry.frx":99F8
                  NumItems        =   9
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "LINE #"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "PART NUMBER"
                     Object.Width           =   2646
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "DESCRIPTION"
                     Object.Width           =   3528
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   3
                     Text            =   "QTY"
                     Object.Width           =   1147
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   4
                     Text            =   "UNITPRICE"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   5
                     Text            =   "AMOUNT"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   6
                     Text            =   "WSC"
                     Object.Width           =   1147
                  EndProperty
                  BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   7
                     Text            =   "DISC."
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   8
                     Text            =   "ID"
                     Object.Width           =   2
                  EndProperty
               End
            End
            Begin VB.Frame fraJobs 
               BackColor       =   &H00404040&
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
               Height          =   2025
               Left            =   -69940
               TabIndex        =   267
               Top             =   600
               Visible         =   0   'False
               Width           =   10035
               Begin MSComctlLib.ListView lstJobs 
                  Height          =   1995
                  Left            =   0
                  TabIndex        =   268
                  Top             =   0
                  Width           =   10005
                  _ExtentX        =   17648
                  _ExtentY        =   3519
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  HoverSelection  =   -1  'True
                  _Version        =   393217
                  ForeColor       =   0
                  BackColor       =   16777215
                  Appearance      =   1
                  MousePointer    =   99
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "DataEntry.frx":9B5A
                  NumItems        =   7
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "LINE #"
                     Object.Width           =   1764
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "JOB CODE"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "JOB DESCRIPTION"
                     Object.Width           =   5292
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   3
                     Text            =   "AMOUNT"
                     Object.Width           =   2646
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   2
                     SubItemIndex    =   4
                     Text            =   "WC"
                     Object.Width           =   2646
                  EndProperty
                  BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Alignment       =   1
                     SubItemIndex    =   5
                     Text            =   "DISCOUNT"
                     Object.Width           =   2328
                  EndProperty
                  BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   6
                     Text            =   "ID"
                     Object.Width           =   2
                  EndProperty
               End
            End
            Begin VB.Frame fraDetails 
               BackColor       =   &H00404040&
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
               Height          =   2055
               Left            =   -69970
               TabIndex        =   264
               Top             =   600
               Visible         =   0   'False
               Width           =   10035
               Begin MSFlexGridLib.MSFlexGrid grdDetails 
                  Height          =   1995
                  Left            =   30
                  TabIndex        =   265
                  Top             =   30
                  Width           =   9975
                  _ExtentX        =   17595
                  _ExtentY        =   3519
                  _Version        =   393216
                  Rows            =   5
                  Cols            =   9
                  ForeColor       =   0
                  ForeColorFixed  =   0
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483633
                  TextStyleFixed  =   3
                  BorderStyle     =   0
                  Appearance      =   0
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
                  MouseIcon       =   "DataEntry.frx":9CBC
               End
               Begin MSComctlLib.ListView ListView1 
                  Height          =   1395
                  Left            =   1440
                  TabIndex        =   266
                  Top             =   3300
                  Width           =   6135
                  _ExtentX        =   10821
                  _ExtentY        =   2461
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FlatScrollBar   =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "NO"
                     Object.Width           =   882
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Techcode"
                     Object.Width           =   2540
                  EndProperty
               End
            End
         End
      End
      Begin VB.PictureBox cmdAddJobs 
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
         Height          =   5505
         Left            =   3450
         ScaleHeight     =   5475
         ScaleWidth      =   5595
         TabIndex        =   101
         Top             =   270
         Width           =   5625
         Begin VB.Frame fraAddJobs 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   5055
            Left            =   0
            TabIndex        =   102
            Top             =   330
            Width           =   5595
            Begin VB.CheckBox optQUICK 
               Caption         =   "QUICK SERVICE"
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
               Height          =   285
               Left            =   3120
               TabIndex        =   122
               Top             =   450
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Edit HRs Work"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4230
               TabIndex        =   121
               Top             =   2970
               Width           =   1215
            End
            Begin VB.ComboBox cboAcctCodeLabor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1710
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   1500
               Width           =   3705
            End
            Begin VB.Frame Frame4 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
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
               Height          =   675
               Left            =   1110
               TabIndex        =   114
               Top             =   1890
               Width           =   2865
               Begin VB.OptionButton optByAmt 
                  Caption         =   "By Amount"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   0
                  TabIndex        =   118
                  Top             =   360
                  Width           =   1395
               End
               Begin VB.OptionButton optByPerc 
                  Caption         =   "By Percentage"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   0
                  TabIndex        =   117
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.TextBox txtJobDiscount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   3
                  TabIndex        =   116
                  Text            =   "0"
                  Top             =   0
                  Width           =   465
               End
               Begin VB.TextBox txtJobDiscountAmt 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1650
                  MaxLength       =   10
                  TabIndex        =   115
                  Text            =   "0"
                  Top             =   360
                  Width           =   1185
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   35
                  Left            =   2190
                  TabIndex        =   119
                  Top             =   60
                  Width           =   210
               End
            End
            Begin VB.TextBox txtJobLineNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1110
               TabIndex        =   113
               Text            =   "Text1"
               Top             =   60
               Width           =   585
            End
            Begin VB.ComboBox cboJobCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1110
               Sorted          =   -1  'True
               TabIndex        =   112
               Text            =   "cboJobCode"
               Top             =   780
               Width           =   4305
            End
            Begin VB.CommandButton cmdJobDelete 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Delete"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   90
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":9FD6
               Style           =   1  'Graphical
               TabIndex        =   111
               ToolTipText     =   "Delete Entry"
               Top             =   4260
               Width           =   825
            End
            Begin VB.ComboBox cboJcode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1110
               Sorted          =   -1  'True
               TabIndex        =   110
               Text            =   "cboJcode"
               Top             =   420
               Width           =   1425
            End
            Begin VB.ComboBox cboJobChargeTo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1110
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   109
               Top             =   1500
               Width           =   585
            End
            Begin VB.TextBox txtJobDetail 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   825
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   108
               Text            =   "DataEntry.frx":B058
               Top             =   3330
               Width           =   5325
            End
            Begin VB.TextBox txtJobRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1110
               MaxLength       =   10
               TabIndex        =   107
               Text            =   "0.00"
               Top             =   1140
               Width           =   1425
            End
            Begin VB.OptionButton optByCode 
               Caption         =   "By &Job Code"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2160
               TabIndex        =   106
               Top             =   60
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.TextBox txtDET_HRS 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   4230
               MaxLength       =   10
               TabIndex        =   105
               Text            =   "0.0"
               Top             =   2640
               Width           =   1185
            End
            Begin VB.ComboBox cboTechnician 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1110
               Sorted          =   -1  'True
               TabIndex        =   104
               Text            =   "cboJobCode"
               Top             =   2640
               Width           =   3075
            End
            Begin VB.TextBox txtJobPostCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2760
               TabIndex        =   103
               Text            =   "Text1"
               Top             =   450
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.CommandButton cmdJobCancel 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancel"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   4620
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":B05E
               Style           =   1  'Graphical
               TabIndex        =   123
               ToolTipText     =   "Cancel"
               Top             =   4260
               Width           =   825
            End
            Begin VB.CommandButton cmdJobSave 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Save"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   3840
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":C0E0
               Style           =   1  'Graphical
               TabIndex        =   124
               ToolTipText     =   "Save Jobs"
               Top             =   4260
               Width           =   795
            End
            Begin VB.Label lblTECHCODE_X 
               BackColor       =   &H000000FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   1500
               TabIndex        =   136
               Top             =   4350
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label81 
               BackColor       =   &H8000000D&
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
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   3300
               TabIndex        =   135
               Top             =   2070
               Width           =   225
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Line No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   11
               Left            =   375
               TabIndex        =   134
               Top             =   90
               Width           =   660
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Job Desc."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   13
               Left            =   240
               TabIndex        =   133
               Top             =   870
               Width           =   795
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Charge To"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   15
               Left            =   180
               TabIndex        =   132
               Top             =   1590
               Width           =   855
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Job Rate"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   14
               Left            =   330
               TabIndex        =   131
               Top             =   1230
               Width           =   705
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Discount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   16
               Left            =   315
               TabIndex        =   130
               Top             =   2100
               Width           =   720
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Enter Job Description"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   18
               Left            =   120
               TabIndex        =   129
               Top             =   3090
               Width           =   1785
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Job Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   12
               Left            =   255
               TabIndex        =   128
               Top             =   510
               Width           =   780
            End
            Begin VB.Label labJobDet_Vol 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               Caption         =   "det_vol"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   285
               Left            =   1530
               TabIndex        =   127
               Top             =   3540
               Width           =   1845
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Technician"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   17
               Left            =   150
               TabIndex        =   126
               Top             =   2670
               Width           =   885
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "No. of Hours Worked"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Index           =   36
               Left            =   4170
               TabIndex        =   125
               Top             =   2190
               Width           =   1305
            End
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Repair Order Job Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   55
            Left            =   0
            MouseIcon       =   "DataEntry.frx":E1B2
            MousePointer    =   99  'Custom
            TabIndex        =   137
            Top             =   0
            Width           =   5595
         End
      End
      Begin VB.PictureBox cmdAddParts 
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
         Height          =   5550
         Left            =   4305
         ScaleHeight     =   5520
         ScaleWidth      =   4545
         TabIndex        =   167
         Top             =   0
         Width           =   4575
         Begin VB.Frame fraAddParts 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Caption         =   "Add/Edit Parts"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   5085
            Left            =   0
            TabIndex        =   168
            Top             =   360
            Width           =   4545
            Begin VB.TextBox txtPartCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   3330
               MaxLength       =   2
               TabIndex        =   184
               Text            =   "Text1"
               Top             =   1530
               Width           =   345
            End
            Begin VB.TextBox txtPartAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   183
               Text            =   "0.00"
               Top             =   2220
               Width           =   1545
            End
            Begin VB.TextBox txtUnitPrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   182
               Text            =   "0.00"
               Top             =   1860
               Width           =   1545
            End
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1200
               MaxLength       =   5
               TabIndex        =   181
               Text            =   "0.0"
               Top             =   1500
               Width           =   555
            End
            Begin VB.ComboBox cboChargeTo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1200
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   180
               Top             =   2580
               Width           =   585
            End
            Begin VB.ComboBox cboPartNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1200
               Sorted          =   -1  'True
               TabIndex        =   179
               Text            =   "cboPartNo"
               Top             =   480
               Width           =   2685
            End
            Begin VB.ComboBox cboDescription 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   90
               Sorted          =   -1  'True
               TabIndex        =   178
               Text            =   "cboDescription"
               Top             =   1110
               Width           =   4335
            End
            Begin VB.CommandButton cmdPartsDelete 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Delete"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   90
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":E4BC
               Style           =   1  'Graphical
               TabIndex        =   177
               Top             =   4230
               Width           =   885
            End
            Begin VB.TextBox txtPartsLineNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1200
               MaxLength       =   4
               TabIndex        =   176
               Text            =   "Text1"
               Top             =   120
               Width           =   525
            End
            Begin VB.Frame Frame5 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
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
               Height          =   675
               Left            =   1200
               TabIndex        =   170
               Top             =   3360
               Width           =   3015
               Begin VB.OptionButton optPartsbyAmt 
                  Caption         =   "By Amount"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   0
                  TabIndex        =   174
                  Top             =   360
                  Width           =   1395
               End
               Begin VB.OptionButton optPartsByPerc 
                  Caption         =   "By Percentage"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   0
                  TabIndex        =   173
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1515
               End
               Begin VB.TextBox txtPartDiscount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1560
                  MaxLength       =   3
                  TabIndex        =   172
                  Text            =   "0"
                  Top             =   0
                  Width           =   465
               End
               Begin VB.TextBox txtPartDiscountAmt 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1560
                  MaxLength       =   10
                  TabIndex        =   171
                  Text            =   "0"
                  Top             =   360
                  Width           =   1305
               End
               Begin VB.Label Label18 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Index           =   67
                  Left            =   2100
                  TabIndex        =   175
                  Top             =   30
                  Width           =   225
               End
            End
            Begin VB.ComboBox cboAcctCodeParts 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   90
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   169
               Top             =   2940
               Width           =   4335
            End
            Begin VB.CommandButton cmdPartsCancel 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancel"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   3570
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":F53E
               Style           =   1  'Graphical
               TabIndex        =   185
               Top             =   4230
               Width           =   885
            End
            Begin VB.CommandButton cmdPartsSave 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Save"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   2760
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":105C0
               Style           =   1  'Graphical
               TabIndex        =   186
               Top             =   4230
               Width           =   855
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Amount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   64
               Left            =   420
               TabIndex        =   194
               Top             =   2220
               Width           =   660
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Charge To"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   65
               Left            =   225
               TabIndex        =   193
               Top             =   2610
               Width           =   855
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Discount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   66
               Left            =   360
               TabIndex        =   192
               Top             =   3510
               Width           =   720
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   61
               Left            =   135
               TabIndex        =   191
               Top             =   870
               Width           =   945
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Line No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   59
               Left            =   420
               TabIndex        =   190
               Top             =   150
               Width           =   660
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Part No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   60
               Left            =   450
               TabIndex        =   189
               Top             =   510
               Width           =   630
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   62
               Left            =   405
               TabIndex        =   188
               Top             =   1500
               Width           =   675
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Price"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   63
               Left            =   300
               TabIndex        =   187
               Top             =   1860
               Width           =   780
            End
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Issued Parts in Repair Order"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   58
            Left            =   0
            MouseIcon       =   "DataEntry.frx":12692
            MousePointer    =   99  'Custom
            TabIndex        =   195
            Top             =   0
            Width           =   4545
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4455
         Left            =   4560
         ScaleHeight     =   4425
         ScaleWidth      =   4005
         TabIndex        =   84
         Top             =   1140
         Visible         =   0   'False
         Width           =   4035
         Begin VB.CommandButton cmdPartClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   2040
            TabIndex        =   98
            Top             =   3720
            Width           =   915
         End
         Begin VB.TextBox txtLOAAmount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   1980
            TabIndex        =   97
            Text            =   "0.00"
            Top             =   450
            Width           =   1755
         End
         Begin VB.CheckBox chkAllowManDist 
            Caption         =   "Enable Manual Distribution"
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
            Height          =   240
            Left            =   120
            TabIndex        =   96
            Top             =   960
            Width           =   3015
         End
         Begin VB.Frame fraParticipation 
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
            Height          =   2355
            Left            =   120
            TabIndex        =   85
            Top             =   1200
            Width           =   3765
            Begin VB.TextBox txtPartTotal 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1860
               TabIndex        =   90
               Text            =   "0.00"
               Top             =   1890
               Width           =   1755
            End
            Begin VB.TextBox txtPartAccessories 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1860
               TabIndex        =   89
               Text            =   "0.00"
               Top             =   1500
               Width           =   1755
            End
            Begin VB.TextBox txtPartMaterials 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1860
               TabIndex        =   88
               Text            =   "0.00"
               Top             =   1080
               Width           =   1755
            End
            Begin VB.TextBox txtPartParts 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1860
               TabIndex        =   87
               Text            =   "0.00"
               Top             =   660
               Width           =   1755
            End
            Begin VB.TextBox txtPartLabor 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   1860
               TabIndex        =   86
               Text            =   "0.00"
               Top             =   240
               Width           =   1755
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   53
               Left            =   690
               TabIndex        =   95
               Top             =   1920
               Width           =   405
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Accessories"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   52
               Left            =   690
               TabIndex        =   94
               Top             =   1530
               Width           =   1050
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Materials"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   49
               Left            =   690
               TabIndex        =   93
               Top             =   1110
               Width           =   765
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Parts"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   51
               Left            =   690
               TabIndex        =   92
               Top             =   690
               Width           =   435
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Labor"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   50
               Left            =   690
               TabIndex        =   91
               Top             =   270
               Width           =   480
            End
         End
         Begin VB.CommandButton cmdPartSave 
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
            Height          =   585
            Left            =   1140
            TabIndex        =   99
            Top             =   3720
            Width           =   915
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   345
            Left            =   0
            TabIndex        =   259
            Top             =   0
            Width           =   4065
            _Version        =   655364
            _ExtentX        =   7170
            _ExtentY        =   609
            _StockProps     =   14
            Caption         =   "Input Insurance Participation"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.01
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            VisualTheme     =   3
            Alignment       =   1
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOA Amount :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   48
            Left            =   150
            TabIndex        =   100
            Top             =   510
            Width           =   1425
         End
      End
      Begin VB.PictureBox cmdDiscount 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
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
         Height          =   1785
         Left            =   5070
         ScaleHeight     =   1755
         ScaleWidth      =   3105
         TabIndex        =   82
         Top             =   2160
         Width           =   3135
         Begin VB.TextBox txtDiscAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Left            =   1140
            MaxLength       =   3
            TabIndex        =   305
            Text            =   "Tex"
            Top             =   480
            Width           =   735
         End
         Begin VB.Frame fraDiscount 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Height          =   1425
            Left            =   180
            TabIndex        =   83
            Top             =   1980
            Width           =   2955
         End
         Begin wizButton.cmd cmdOkDisc 
            Height          =   405
            Left            =   570
            TabIndex        =   306
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   714
            TX              =   "&Ok"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "DataEntry.frx":1299C
         End
         Begin wizButton.cmd cmdCancelDisk 
            Height          =   405
            Left            =   1560
            TabIndex        =   307
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   714
            TX              =   "&Cancel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "DataEntry.frx":129B8
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption5 
            Height          =   375
            Left            =   -30
            TabIndex        =   309
            Top             =   -30
            Width           =   3195
            _Version        =   655364
            _ExtentX        =   5636
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Enter Discount Percentage"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label Label23 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1920
            TabIndex        =   308
            Top             =   510
            Width           =   345
         End
      End
      Begin VB.PictureBox picCustLimit 
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
         Height          =   2955
         Left            =   4695
         ScaleHeight     =   2925
         ScaleWidth      =   3765
         TabIndex        =   75
         Top             =   1950
         Width           =   3795
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Customer Credit Limit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   46
            Left            =   0
            TabIndex        =   81
            Top             =   0
            Width           =   3795
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Terms in Days"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   43
            Left            =   60
            TabIndex        =   80
            Top             =   480
            Width           =   1740
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Limit Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   44
            Left            =   60
            TabIndex        =   79
            Top             =   1260
            Width           =   1680
         End
         Begin VB.Label labCreditDays 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "30 Days"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1380
            TabIndex        =   78
            Top             =   840
            Width           =   2235
         End
         Begin VB.Label labCreditLimit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "50,000.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1380
            TabIndex        =   77
            Top             =   1620
            Width           =   2235
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   3750
            Y1              =   2250
            Y2              =   2250
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Terms and Limit can be adjusted in Customer Master File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   465
            Index           =   45
            Left            =   60
            TabIndex        =   76
            Top             =   2340
            Width           =   3555
         End
      End
      Begin VB.PictureBox cmdAddAccessories 
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
         Height          =   5595
         Left            =   4515
         ScaleHeight     =   5565
         ScaleWidth      =   4125
         TabIndex        =   138
         Top             =   225
         Width           =   4155
         Begin VB.Frame fraAddAccessories 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Caption         =   "Add/Edit Accessories"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   5175
            Left            =   0
            TabIndex        =   139
            Top             =   300
            Width           =   4155
            Begin VB.TextBox txtAccLineNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1230
               TabIndex        =   157
               Text            =   "01"
               Top             =   90
               Width           =   555
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1230
               MaxLength       =   2
               TabIndex        =   154
               Text            =   "Text1"
               Top             =   5370
               Width           =   375
            End
            Begin VB.ComboBox cboAccCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1230
               Sorted          =   -1  'True
               TabIndex        =   153
               Text            =   "cboMatCode"
               Top             =   450
               Width           =   2655
            End
            Begin VB.ComboBox cboAccessories 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   120
               Sorted          =   -1  'True
               TabIndex        =   152
               Text            =   "cboMaterial"
               Top             =   1080
               Width           =   3765
            End
            Begin VB.CommandButton cmdAccDelete 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Delete"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Left            =   120
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":129D4
               Style           =   1  'Graphical
               TabIndex        =   151
               ToolTipText     =   "Delete Entry"
               Top             =   4290
               Width           =   825
            End
            Begin VB.ComboBox cboAccChargeTo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1230
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   150
               Top             =   2610
               Width           =   585
            End
            Begin VB.TextBox txtAccQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1230
               MaxLength       =   5
               TabIndex        =   149
               Text            =   "0.0"
               Top             =   1500
               Width           =   555
            End
            Begin VB.TextBox txtAccUnitPrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1230
               MaxLength       =   10
               TabIndex        =   148
               Text            =   "0.00"
               Top             =   1860
               Width           =   1545
            End
            Begin VB.TextBox txtAccAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1230
               MaxLength       =   10
               TabIndex        =   147
               Text            =   "0.00"
               Top             =   2220
               Width           =   1545
            End
            Begin VB.Frame Frame7 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
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
               Height          =   675
               Left            =   1200
               TabIndex        =   141
               Top             =   3450
               Width           =   3015
               Begin VB.OptionButton optAccByAmt 
                  Caption         =   "By Amount"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   0
                  TabIndex        =   145
                  Top             =   360
                  Width           =   1395
               End
               Begin VB.OptionButton optAccByPerc 
                  Caption         =   "By Percentage"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   0
                  TabIndex        =   144
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1515
               End
               Begin VB.TextBox txtAccDiscount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   3
                  TabIndex        =   143
                  Text            =   "0"
                  Top             =   0
                  Width           =   465
               End
               Begin VB.TextBox txtAccDiscountAmt 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1530
                  MaxLength       =   10
                  TabIndex        =   142
                  Text            =   "0"
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.Label Label18 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Index           =   57
                  Left            =   2040
                  TabIndex        =   146
                  Top             =   30
                  Width           =   225
               End
            End
            Begin VB.ComboBox cboAcctCodeAccessories 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   120
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   140
               Top             =   3000
               Width           =   3765
            End
            Begin VB.CommandButton cmdAccCancel 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancel"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Left            =   3210
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":13A56
               Style           =   1  'Graphical
               TabIndex        =   155
               ToolTipText     =   "Cancel Entry"
               Top             =   4290
               Width           =   825
            End
            Begin VB.CommandButton cmdAccSave 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Save"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Left            =   2430
               MaskColor       =   &H0000FFFF&
               Picture         =   "DataEntry.frx":14AD8
               Style           =   1  'Graphical
               TabIndex        =   156
               ToolTipText     =   "Save Materials"
               Top             =   4290
               Width           =   795
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Amount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   32
               Left            =   450
               TabIndex        =   165
               Top             =   2220
               Width           =   660
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Charge To"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   33
               Left            =   255
               TabIndex        =   164
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Discount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   34
               Left            =   390
               TabIndex        =   163
               Top             =   3600
               Width           =   720
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Description "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   29
               Left            =   120
               TabIndex        =   162
               Top             =   810
               Width           =   990
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Line No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   27
               Left            =   450
               TabIndex        =   161
               Top             =   120
               Width           =   660
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   30
               Left            =   435
               TabIndex        =   160
               Top             =   1530
               Width           =   675
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Price"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   31
               Left            =   330
               TabIndex        =   159
               Top             =   1890
               Width           =   780
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Acc. Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   28
               Left            =   285
               TabIndex        =   158
               Top             =   480
               Width           =   825
            End
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Issued Accessories in Repair Order"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   56
            Left            =   0
            MouseIcon       =   "DataEntry.frx":16BAA
            MousePointer    =   99  'Custom
            TabIndex        =   166
            Top             =   0
            Width           =   4125
         End
      End
      Begin VB.PictureBox cmdAddMaterials 
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
         Height          =   5775
         Left            =   3480
         Negotiate       =   -1  'True
         ScaleHeight     =   5745
         ScaleWidth      =   4065
         TabIndex        =   196
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Frame fraAddMaterials 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Caption         =   "Add/Edit Materials"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   6075
            Left            =   0
            TabIndex        =   197
            Top             =   300
            Width           =   4005
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1740
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   321
               Top             =   90
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.PictureBox piccontol 
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   240
               Negotiate       =   -1  'True
               ScaleHeight     =   855
               ScaleWidth      =   3615
               TabIndex        =   314
               Top             =   5160
               Width           =   3615
               Begin VB.CommandButton cmdMatSave 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "&Save"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   855
                  Left            =   960
                  MaskColor       =   &H0000FFFF&
                  Picture         =   "DataEntry.frx":16EB4
                  Style           =   1  'Graphical
                  TabIndex        =   317
                  ToolTipText     =   "Save Materials"
                  Top             =   0
                  Width           =   855
               End
               Begin VB.CommandButton cmdMatCancel 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "&Cancel"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   855
                  Left            =   2730
                  MaskColor       =   &H0000FFFF&
                  Picture         =   "DataEntry.frx":18F86
                  Style           =   1  'Graphical
                  TabIndex        =   316
                  ToolTipText     =   "Cancel Entry"
                  Top             =   0
                  Width           =   885
               End
               Begin VB.CommandButton cmdMatDelete 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "&Delete"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   855
                  Left            =   0
                  MaskColor       =   &H0000FFFF&
                  Picture         =   "DataEntry.frx":1A008
                  Style           =   1  'Graphical
                  TabIndex        =   315
                  ToolTipText     =   "Delete Entry"
                  Top             =   0
                  Width           =   885
               End
            End
            Begin VB.TextBox txtdetail 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   825
               Left            =   240
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   0
               Top             =   4200
               Width           =   3645
            End
            Begin VB.TextBox txtMatAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1110
               MaxLength       =   10
               TabIndex        =   212
               Text            =   "0.00"
               Top             =   2250
               Width           =   1545
            End
            Begin VB.TextBox txtMatUnitPrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1110
               MaxLength       =   10
               TabIndex        =   211
               Text            =   "0.00"
               Top             =   1890
               Width           =   1545
            End
            Begin VB.TextBox txtMatQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1110
               MaxLength       =   5
               TabIndex        =   210
               Text            =   "0.0"
               Top             =   1530
               Width           =   555
            End
            Begin VB.ComboBox cboMatChargeTo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1110
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   209
               Top             =   2610
               Width           =   585
            End
            Begin VB.ComboBox cboMaterial 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   120
               Sorted          =   -1  'True
               TabIndex        =   208
               Text            =   "cboMaterial"
               Top             =   1080
               Width           =   3735
            End
            Begin VB.ComboBox cboMatCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1110
               Sorted          =   -1  'True
               TabIndex        =   207
               Text            =   "cboMatCode"
               Top             =   480
               Width           =   2745
            End
            Begin VB.TextBox txtMatPOCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               MaxLength       =   2
               TabIndex        =   206
               Text            =   "Text1"
               Top             =   5640
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtMatLineNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1110
               TabIndex        =   205
               Text            =   "Text1"
               Top             =   120
               Width           =   555
            End
            Begin VB.Frame Frame6 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
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
               Height          =   675
               Left            =   1080
               TabIndex        =   199
               Top             =   3420
               Width           =   3015
               Begin VB.TextBox txtMatDiscountAmt 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1590
                  MaxLength       =   10
                  TabIndex        =   203
                  Text            =   "0"
                  Top             =   360
                  Width           =   1155
               End
               Begin VB.TextBox txtMatDiscount 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   1590
                  MaxLength       =   3
                  TabIndex        =   202
                  Text            =   "0"
                  Top             =   0
                  Width           =   465
               End
               Begin VB.OptionButton optMatByPerc 
                  Caption         =   "By Percentage"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   0
                  TabIndex        =   201
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1545
               End
               Begin VB.OptionButton optMatByAmt 
                  Caption         =   "By Amount"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   0
                  TabIndex        =   200
                  Top             =   360
                  Width           =   1395
               End
               Begin VB.Label Label2 
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Left            =   2160
                  TabIndex        =   204
                  Top             =   30
                  Width           =   225
               End
            End
            Begin VB.ComboBox cboAcctCodeMaterials 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   120
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   198
               Top             =   3000
               Width           =   3735
            End
            Begin VB.Label lbldetail 
               Caption         =   "Detail"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   313
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Mat. Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   25
               Left            =   165
               TabIndex        =   220
               Top             =   480
               Width           =   825
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Price"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   22
               Left            =   210
               TabIndex        =   219
               Top             =   1920
               Width           =   780
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   23
               Left            =   315
               TabIndex        =   218
               Top             =   1560
               Width           =   675
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Line No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   26
               Left            =   330
               TabIndex        =   217
               Top             =   150
               Width           =   660
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Material"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Index           =   24
               Left            =   330
               TabIndex        =   216
               Top             =   840
               Width           =   660
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Discount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   19
               Left            =   270
               TabIndex        =   215
               Top             =   3600
               Width           =   720
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Charge To"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   20
               Left            =   135
               TabIndex        =   214
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Amount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   21
               Left            =   330
               TabIndex        =   213
               Top             =   2250
               Width           =   660
            End
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Issued Materials in Repair Order"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   68
            Left            =   0
            MouseIcon       =   "DataEntry.frx":1B08A
            MousePointer    =   99  'Custom
            TabIndex        =   221
            Top             =   0
            Width           =   4125
         End
      End
      Begin VB.PictureBox cmdRelBut 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   5010
         ScaleHeight     =   1575
         ScaleWidth      =   3135
         TabIndex        =   231
         Top             =   2610
         Width           =   3165
         Begin VB.Frame fraRelBut 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   30
            TabIndex        =   232
            Top             =   330
            Width           =   3075
            Begin VB.TextBox txtReleaseDate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   465
               Left            =   390
               MaxLength       =   10
               TabIndex        =   233
               Text            =   "Text1"
               Top             =   120
               Width           =   2295
            End
            Begin wizButton.cmd cmdCancelRel 
               Height          =   375
               Left            =   1560
               TabIndex        =   234
               Top             =   690
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   661
               TX              =   "&Cancel"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FOCUSR          =   -1  'True
               MPTR            =   0
               MICON           =   "DataEntry.frx":1B394
            End
            Begin wizButton.cmd cmdOkRel 
               Height          =   375
               Left            =   570
               TabIndex        =   235
               Top             =   690
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   661
               TX              =   "&Ok"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FOCUSR          =   -1  'True
               MPTR            =   0
               MICON           =   "DataEntry.frx":1B3B0
            End
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   315
            Index           =   10
            Left            =   0
            TabIndex        =   236
            Top             =   0
            Width           =   3135
            _Version        =   655364
            _ExtentX        =   5530
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Enter RO Release Date"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.01
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientColorLight=   16711680
            GradientColorDark=   8388608
         End
      End
      Begin VB.PictureBox cmdBillBut 
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
         Height          =   2895
         Left            =   4320
         ScaleHeight     =   2865
         ScaleWidth      =   4515
         TabIndex        =   237
         Top             =   1650
         Width           =   4545
         Begin VB.Frame fraBillBut 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Height          =   2445
            Left            =   -30
            TabIndex        =   238
            Top             =   450
            Width           =   4455
            Begin VB.OptionButton Option2 
               Caption         =   "CHARGE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1470
               TabIndex        =   243
               Top             =   120
               Width           =   1485
            End
            Begin VB.OptionButton Option1 
               Caption         =   "CASH"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   180
               TabIndex        =   242
               Top             =   120
               Value           =   -1  'True
               Width           =   1245
            End
            Begin VB.TextBox txtDateReleased 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   465
               Left            =   2280
               MaxLength       =   10
               TabIndex        =   241
               Text            =   "Text1"
               Top             =   2880
               Width           =   2115
            End
            Begin VB.TextBox txtInvoiceNumber 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   465
               Left            =   150
               MaxLength       =   6
               TabIndex        =   240
               Text            =   "Text1"
               Top             =   1560
               Width           =   1875
            End
            Begin VB.TextBox txtInvoiceDate 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   465
               Left            =   2280
               MaxLength       =   10
               TabIndex        =   239
               Top             =   480
               Width           =   2115
            End
            Begin wizButton.cmd cmdCancelBill 
               Height          =   465
               Left            =   3360
               TabIndex        =   244
               ToolTipText     =   "Cancel"
               Top             =   1560
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   820
               TX              =   "&Cancel"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FOCUSR          =   -1  'True
               MPTR            =   0
               MICON           =   "DataEntry.frx":1B3CC
            End
            Begin wizButton.cmd cmdOkBill 
               Height          =   465
               Left            =   2400
               TabIndex        =   245
               ToolTipText     =   "Ok"
               Top             =   1560
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   820
               TX              =   "&Ok"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FOCUSR          =   -1  'True
               MPTR            =   0
               MICON           =   "DataEntry.frx":1B3E8
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00DEDFDE&
               Caption         =   "*** PLEASE MAKE SURE ALL ENTRIES ARE CORRECT ***"
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
               Height          =   255
               Left            =   120
               TabIndex        =   249
               Top             =   2130
               Width           =   4305
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Enter Invoice Number"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   435
               Index           =   70
               Left            =   90
               TabIndex        =   248
               Top             =   1020
               Width           =   4275
            End
            Begin VB.Label Label57 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Released"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   247
               Top             =   2940
               Width           =   2115
            End
            Begin VB.Label Label18 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Invoice Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   69
               Left            =   150
               TabIndex        =   246
               Top             =   570
               Width           =   1905
            End
         End
         Begin VB.PictureBox picoverride 
            Height          =   1095
            Left            =   360
            ScaleHeight     =   1035
            ScaleWidth      =   3555
            TabIndex        =   322
            Top             =   1080
            Visible         =   0   'False
            Width           =   3615
            Begin VB.TextBox txtoverride 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   435
               IMEMode         =   3  'DISABLE
               Left            =   360
               PasswordChar    =   "*"
               TabIndex        =   323
               Top             =   480
               Width           =   2805
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   345
               Index           =   0
               Left            =   -360
               TabIndex        =   324
               Top             =   0
               Width           =   4515
               _Version        =   655364
               _ExtentX        =   7964
               _ExtentY        =   609
               _StockProps     =   14
               Caption         =   "Enter Code to Override"
               ForeColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   1
               GradientColorLight=   255
               GradientColorDark=   4210752
               ForeColor       =   16777215
            End
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   345
            Index           =   4
            Left            =   0
            TabIndex        =   250
            Top             =   0
            Width           =   4515
            _Version        =   655364
            _ExtentX        =   7964
            _ExtentY        =   609
            _StockProps     =   14
            Caption         =   "GENERATE SERVICE INVOICE "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.01
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            GradientColorLight=   8421504
            GradientColorDark=   4210752
         End
      End
      Begin VB.PictureBox cmdFollowUp 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   3518
         ScaleHeight     =   3345
         ScaleWidth      =   6105
         TabIndex        =   287
         Top             =   2033
         Visible         =   0   'False
         Width           =   6135
         Begin VB.CommandButton Command6 
            Caption         =   "Cancel"
            Height          =   405
            Left            =   4470
            TabIndex        =   293
            Top             =   2850
            Width           =   1485
         End
         Begin VB.CommandButton cmdSaveFollowUp 
            Caption         =   "Save Result"
            Height          =   405
            Left            =   3000
            TabIndex        =   292
            Top             =   2850
            Width           =   1485
         End
         Begin VB.CheckBox chkCALLED_FOLLOWUP 
            Caption         =   "CUSTOMER CALLED FOR FOLLOW UP"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   150
            TabIndex        =   288
            Top             =   390
            Value           =   1  'Checked
            Width           =   4575
         End
         Begin RichTextLib.RichTextBox txtCALLED_RESULT 
            Height          =   2115
            Left            =   150
            TabIndex        =   289
            Top             =   690
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   3731
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"DataEntry.frx":1B404
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   291
            Top             =   0
            Width           =   6195
            _Version        =   655364
            _ExtentX        =   10927
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "FOLLOW UP RESULT/REMARKS"
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
         End
         Begin VB.Label Label18 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   39
            Left            =   150
            TabIndex        =   290
            Top             =   660
            Width           =   2955
         End
      End
      Begin VB.Label labDetId 
         BackColor       =   &H000000C0&
         Caption         =   "Label48"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1740
         TabIndex        =   284
         Top             =   6450
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label labID 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "0"
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
         Left            =   900
         TabIndex        =   283
         Top             =   6450
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label labPrevID 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "Label59"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   282
         Top             =   6450
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SJ # :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7950
         TabIndex        =   261
         Top             =   6810
         Width           =   840
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OR # :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5790
         TabIndex        =   260
         Top             =   6810
         Width           =   840
      End
      Begin VB.Label labF1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F1 - Input Notes after follow up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   30
         TabIndex        =   257
         Top             =   6810
         Width           =   2865
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F2 - Input Insurance Participation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   2910
         TabIndex        =   256
         Top             =   6810
         Width           =   2865
      End
      Begin VB.Label labORNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6630
         TabIndex        =   255
         Top             =   6810
         Width           =   1305
      End
      Begin VB.Label labSJNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8790
         TabIndex        =   254
         Top             =   6810
         Width           =   1305
      End
      Begin VB.Label labAddOrEdit 
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -510
         TabIndex        =   253
         Top             =   7590
         Width           =   255
      End
      Begin VB.Label lblMSG2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   225
         Left            =   1530
         TabIndex        =   251
         Top             =   6210
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblMSG 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   225
         Left            =   1530
         TabIndex        =   252
         Top             =   6000
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   258
      Top             =   2820
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7435
      Left            =   30
      ScaleHeight     =   7410
      ScaleWidth      =   13065
      TabIndex        =   21
      Top             =   30
      Width           =   13095
      Begin XtremeReportControl.ReportControl rptRO 
         Height          =   5805
         Left            =   30
         TabIndex        =   286
         Top             =   1110
         Width           =   12945
         _Version        =   655364
         _ExtentX        =   22834
         _ExtentY        =   10239
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnReorder=   0   'False
         MultipleSelection=   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00800000&
         Caption         =   "Model &Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   10860
         Style           =   1  'Graphical
         TabIndex        =   304
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00808000&
         Caption         =   "&Vin no."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   303
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00008000&
         Caption         =   "&Plate No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   302
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00008080&
         Caption         =   "I&nvoice No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   301
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00004080&
         Caption         =   "Repair &Order no"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   300
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H00000080&
         Caption         =   "&Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   299
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   885
         Left            =   12180
         MouseIcon       =   "DataEntry.frx":1B482
         MousePointer    =   99  'Custom
         Picture         =   "DataEntry.frx":1B5D4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Select this Customer"
         Top             =   6420
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   825
         Left            =   90
         MouseIcon       =   "DataEntry.frx":1B910
         MousePointer    =   99  'Custom
         Picture         =   "DataEntry.frx":1BA62
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cancel"
         Top             =   7800
         Visible         =   0   'False
         Width           =   825
      End
      Begin MSComctlLib.ListView lsvSearch 
         Height          =   795
         Left            =   5430
         TabIndex        =   23
         Top             =   5310
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1402
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Repair Order #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Invoice no"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Plate no"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Vin no"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vehicle Model"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1770
         TabIndex        =   22
         Top             =   750
         Width           =   6825
      End
      Begin VB.Shape shpColor 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   3
         Left            =   4770
         Top             =   7110
         Width           =   135
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "- FINISHED JOB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4920
         MouseIcon       =   "DataEntry.frx":1BDA0
         MousePointer    =   99  'Custom
         TabIndex        =   312
         Top             =   7080
         Width           =   1260
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "- RELEASED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3180
         MouseIcon       =   "DataEntry.frx":1BEF2
         MousePointer    =   99  'Custom
         TabIndex        =   298
         Top             =   7080
         Width           =   945
      End
      Begin VB.Shape shpColor 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   2
         Left            =   3030
         Top             =   7110
         Width           =   135
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "- INVOICED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1560
         MouseIcon       =   "DataEntry.frx":1C044
         MousePointer    =   99  'Custom
         TabIndex        =   297
         Top             =   7080
         Width           =   915
      End
      Begin VB.Shape shpColor 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   1
         Left            =   1380
         Top             =   7110
         Width           =   135
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "- PARK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   270
         MouseIcon       =   "DataEntry.frx":1C196
         MousePointer    =   99  'Custom
         TabIndex        =   296
         Top             =   7080
         Width           =   570
      End
      Begin VB.Shape shpColor 
         BackStyle       =   1  'Opaque
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   0
         Left            =   90
         Top             =   7110
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Type Keyword Here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   295
         Top             =   870
         Width           =   1620
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
         Height          =   330
         Index           =   1
         Left            =   -30
         TabIndex        =   294
         Top             =   0
         Width           =   13095
         _Version        =   655364
         _ExtentX        =   23098
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "  Search By"
         ForeColor       =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   4194304
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Press F3 to Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   8730
         TabIndex        =   262
         Top             =   870
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmCSMSDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPOR                                            As ADODB.Recordset
Dim rsRO_DET                                           As ADODB.Recordset
Dim rsCustomer                                         As ADODB.Recordset
Dim rsEmpNo                                            As ADODB.Recordset
Dim rsS_Model                                          As ADODB.Recordset
Dim rsROJOBS                                           As ADODB.Recordset
Dim rsROJOBS2                                          As ADODB.Recordset
Dim rsRemarks                                          As ADODB.Recordset

Dim rsORD_HIST                                         As ADODB.Recordset
Dim rsDAYTRAN                                          As ADODB.Recordset
Dim rsOrd_Hd                                           As ADODB.Recordset
Dim rsTdayTran                                         As ADODB.Recordset
Dim rsPartMas                                          As ADODB.Recordset
Dim rsMatMas                                           As ADODB.Recordset
Dim rsMATISS                                           As ADODB.Recordset
Dim rsACCISS                                           As ADODB.Recordset

Dim rsCSMS_INVOICE                                     As ADODB.Recordset

Public TOTJOBAMT                                       As Double
Public TOTJOBDISC                                      As Double
Public TOTJOBDISCVAL                                   As Double
Public TOTJOBTAX                                       As Double

Public TOTPARTSAMT                                     As Double
Public TOTPARTSDISC                                    As Double
Public TOTPARTSDISCVAL                                 As Double
Public TOTPARTSTAX                                     As Double

Public TOTMATAMT                                       As Double
Public TOTMATDISC                                      As Double
Public TOTMATDISCVAL                                   As Double
Public TOTMATTAX                                       As Double

'For Accessories - FML - 08282007
Public TOTACCAMT                                       As Double
Public TOTACCDISC                                      As Double
Public TOTACCDISCVAL                                   As Double
Public TOTACCTAX                                       As Double

Dim JobTotal                                           As Double
Dim JobComTotal                                        As Double
Dim JobSalesTotal                                      As Double
Dim JobWarTotal                                        As Double
Dim JobInsTotal                                        As Double
Dim JobDiscTotal                                       As Double
Dim JobVatTotal                                        As Double

Dim PartsTotal                                         As Double
Dim PartsComTotal                                      As Double
Dim PartsSalesTotal                                    As Double
Dim PartsWarTotal                                      As Double
Dim PartsInsTotal                                      As Double
Dim PartsDiscTotal                                     As Double
Dim PartsVatTotal                                      As Double

Dim MatTotal                                           As Double
Dim MatComTotal                                        As Double
Dim MatSalesTotal                                      As Double
Dim MatWarTotal                                        As Double
Dim MatInsTotal                                        As Double
Dim MatDiscTotal                                       As Double
Dim MatVatTotal                                        As Double

'For Accessories - FML - 08282007
Dim ACCTotal                                           As Double
Dim ACCComTotal                                        As Double
Dim ACCSalesTotal                                      As Double
Dim ACCWarTotal                                        As Double
Dim AccInsTotal                                        As Double
Dim ACCDiscTotal                                       As Double
Dim ACCVatTotal                                        As Double

Dim COMTotal                                           As Double
Dim SALESTotal                                         As Double
Dim WARTotal                                           As Double
Dim INSTotal                                           As Double
Dim VATTotal                                           As Double
Dim ROTotal                                            As Double

Dim AddorEdit                                          As String
Dim kcnt                                               As Integer
Dim Pcnt                                               As Integer
Dim Mcnt                                               As Integer
Dim Acnt                                               As Integer
Dim DiscTotal                                          As Double

Dim PrevRoNumber                                        As String
Dim RO_RIV_Tranno(100)                                  As Integer
Dim RO_RIV_Tranno_Counter                               As Integer
Dim RO_MRIS_Tranno(100)                                 As Integer
Dim RO_MRIS_Tranno_Counter                              As Integer
Dim flag                                                As Boolean    'BTT - 05222007
Dim Vusercode, VLastUpdate, VLastUpdateTime             As String
Dim vPIS_NO_CHARGE_TO                                   As String

Dim REPRINT                                             As String

Dim PREV_LABOR_CHARGE_TO                                As String
Dim PREV_PARTS_CHARGE_TO                                As String
Dim PREV_ACCESSORIES_CHARGE_TO                          As String
Dim PREV_MATERIALS_CHARGE_TO                            As String
Dim ictr                                                As Integer
Dim Ichg                                                As Boolean
Dim WithEvents frm                                      As frmCSMSROCusveh
Attribute frm.VB_VarHelpID = -1
Dim WithEvents FRMx                                     As frmCSMS_MasterSearchCustomer
Attribute FRMx.VB_VarHelpID = -1

Function GetTaym()
    Dim rstmp                                          As New ADODB.Recordset
    Dim X                                              As Integer
    Dim cnt                                            As Integer
    cnt = 0
    Set rstmp = gconDMIS.Execute("Select PromiseDate From CSMS_RepairOrder Where RO_no = '" & txtRep_Or.Text & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        For X = 1 To Len(rstmp!PromiseDate)
            If Mid(rstmp!PromiseDate, X, 1) = "/" Then cnt = cnt + 1
            If cnt = 2 Then
                GetTaym = Mid(rstmp!PromiseDate, X + 6, Len(rstmp!PromiseDate) - X)
                Exit For
            End If
        Next
    End If

    Set rstmp = Nothing
End Function

Function CheckIfPlateNoAlreadyExist() As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("Select Plate_NO From CSMS_CusVeh Where Plate_NO = '" & txtPlate_No.Text & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        CheckIfPlateNoAlreadyExist = True
    Else
        CheckIfPlateNoAlreadyExist = False
    End If

    Set rstmp = Nothing
End Function

Function ReturnVehicleID(vplate_no As String) As Integer
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("select ID from csms_Cusveh where plate_no = '" & vplate_no & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        ReturnVehicleID = rstmp!ID
    End If
    Set rstmp = Nothing
End Function

Function SetPartDesc(ppp As String)
    Set rsPartMas = New ADODB.Recordset
    Set rsPartMas = gconDMIS.Execute("Select partno,partdesc from PMIS_PartMas where partno = '" & ppp & "'")
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartDesc = Null2String(rsPartMas!PartDesc)
    Else
        SetPartDesc = cboDescription.Text
    End If
    Set rsPartMas = Nothing
End Function

Function setTechnicianCode(jjj As String)
    If jjj <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select Technician from CSMS_vw_Technician where Tech_Name = '" & jjj & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setTechnicianCode = Null2String(rsROJOBS!Technician)
        Else
            setTechnicianCode = ""
        End If
        Set rsROJOBS = Nothing
    End If
End Function

Function setTechnicianName(jjj As String)
    If jjj <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select Tech_Name from CSMS_vw_Technician where ltrim(rtrim(Technician)) = '" & jjj & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setTechnicianName = Null2String(rsROJOBS!TECH_NAME)
        Else
            setTechnicianName = ""
        End If
        Set rsROJOBS = Nothing
    End If
End Function

Function SetAcctCode(VVV As Variant) As String
    Dim rsChartAccount2                                As New ADODB.Recordset
    rsChartAccount2.Open "Select * from CMIS_SBOOK where BOOK = 'S' AND DESCNAME = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        SetAcctCode = Null2String(rsChartAccount2!Code)
    Else
        SetAcctCode = ""
    End If
End Function

Function SetAcctName(VVV As Variant) As String
    Dim rsChartAccount2                                As New ADODB.Recordset
    rsChartAccount2.Open "Select * from CMIS_SBOOK where BOOK = 'S' AND CODE = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        SetAcctName = Null2String(rsChartAccount2!DESCNAME)
    Else
        SetAcctName = ""
    End If
End Function

Function setJobCode(jjj As String)
    If jjj <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select jcode,desc1,std_mhrs from CSMS_Jobs where desc1 = '" & jjj & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setJobCode = Null2String(rsROJOBS!JCode)
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobCode = ""
            labJobDet_Vol.Caption = 0
        End If
        Set rsROJOBS = Nothing
    End If
End Function

Function setJobDesc(jjj As String)
    If jjj <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select jcode,desc1,std_mhrs from CSMS_Jobs where jcode = '" & jjj & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setJobDesc = Null2String(rsROJOBS!desc1)
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobDesc = ""
            labJobDet_Vol.Caption = 0
        End If
        Set rsROJOBS = Nothing
    End If
End Function

Function setJobPOcode(ppp As String)
    If ppp <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select desc1,std_mhrs from CSMS_Jobs where desc1 = '" & Repleys(ppp) & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobPOcode = ""
            labJobDet_Vol.Caption = 0
        End If
        Set rsROJOBS = Nothing
    End If
End Function

Function setJobDetail(ppp As String)
    If ppp <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select desc1,detail,std_mhrs from CSMS_Jobs where desc1 = '" & ppp & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setJobDetail = Null2String(rsROJOBS!Detail)
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobDetail = ""
            labJobDet_Vol.Caption = 0
        End If
        Set rsROJOBS = Nothing
    End If
End Function

Function setJobRate(ppp As String)
    If ppp <> "" Then
        Set rsROJOBS = New ADODB.Recordset
        Set rsROJOBS = gconDMIS.Execute("Select desc1,flatrate,std_mhrs from CSMS_Jobs where desc1 = '" & Repleys(ppp) & "'")
        If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
            setJobRate = Null2String(rsROJOBS!FLATRATE)
            labJobDet_Vol.Caption = N2Str2IntZero(rsROJOBS!std_mhrs)
        Else
            setJobRate = 0
            labJobDet_Vol.Caption = 0
        End If
        Set rsROJOBS = Nothing
    End If
End Function

Function CheckHasCreditLimit(XXX As String) As Boolean
    Dim rsALL_Customer_Credit                          As New ADODB.Recordset
    Set rsALL_Customer_Credit = gconDMIS.Execute("Select CREDITLIMIT from ALL_Customer Where CUSCDE = '" & XXX & "'")
    If Not rsALL_Customer_Credit.EOF And Not rsALL_Customer_Credit.BOF Then
        If N2Str2Zero(rsALL_Customer_Credit!CreditLimit) > 0 Then
            CheckHasCreditLimit = True
        Else
            CheckHasCreditLimit = False
        End If
    Else
        CheckHasCreditLimit = False
    End If
End Function

Function CheckORNum(yyy As String) As String
    Dim rsCMIS_OFF_DT                                  As New ADODB.Recordset
    Set rsCMIS_OFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE TRANTYPE = 'SI' AND INVOICENO = '" & yyy & "' AND (CANCEL = 0 OR CANCEL IS NULL)")
    If Not rsCMIS_OFF_DT.EOF And Not rsCMIS_OFF_DT.BOF Then
        CheckORNum = Null2String(rsCMIS_OFF_DT!OR_NUM)
    End If
    Set rsCMIS_OFF_DT = Nothing
End Function

Function CheckSJNum(yyy As String) As String
    Dim rsAMIS_JournalSJ                               As New ADODB.Recordset
    Set rsAMIS_JournalSJ = gconDMIS.Execute("Select * from AMIS_JOURNAL_HD WHERE INVOICETYPE = 'SI' AND INVOICENO = '" & yyy & "' AND STATUS <> 'C' AND JTYPE = 'SJ'")
    If Not rsAMIS_JournalSJ.EOF And Not rsAMIS_JournalSJ.BOF Then
        CheckSJNum = Null2String(rsAMIS_JournalSJ!VOUCHERNO)
    End If
    Set rsAMIS_JournalSJ = Nothing
End Function

Function CheckSJINTRONum(yyy As String) As String
    Dim rsAMIS_JournalSJ                               As New ADODB.Recordset
    Set rsAMIS_JournalSJ = gconDMIS.Execute("Select * from AMIS_JOURNAL_HD WHERE INVOICETYPE = 'SI' AND REFNO = '" & yyy & "' AND STATUS <> 'C'")
    If Not rsAMIS_JournalSJ.EOF And Not rsAMIS_JournalSJ.BOF Then
        CheckSJINTRONum = Null2String(rsAMIS_JournalSJ!VOUCHERNO)
    End If
    Set rsAMIS_JournalSJ = Nothing
End Function

Function SetMake(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select Description from CSMS_CusVeh where Plate_No = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetMake = Null2String(rsS_Model!Description) Else SetMake = ""
    Set rsS_Model = Nothing
End Function

Function SetSA(emp As String)
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym from CSMS_vw_EmpNo where code = '" & emp & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetSA = Null2String(rsEmpNo!NAYM)
    Set rsEmpNo = Nothing
End Function

Function SetCodeSA(nam As String)
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym from CSMS_vw_EmpNo where naym = '" & nam & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetCodeSA = Null2String(rsEmpNo!Code)
    Set rsEmpNo = Nothing
End Function

Function CheckIfHasSellingDealer(XXX As String) As Boolean
    Dim rsCusVeh                                       As New ADODB.Recordset
    Set rsCusVeh = gconDMIS.Execute("Select SELLING_DEALER from CSMS_CUSVEH Where PLATE_NO = '" & XXX & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        If Null2String(rsCusVeh!Selling_Dealer) = "" Then
            CheckIfHasSellingDealer = False
        Else
            If CheckIfValidSellingDealer(Null2String(rsCusVeh!Selling_Dealer)) = True Then
                CheckIfHasSellingDealer = True
            Else
                CheckIfHasSellingDealer = False
            End If
        End If
    End If
    Set rsCusVeh = Nothing
End Function

Function CheckIfValidSellingDealer(XXX As String) As Boolean
    Dim rsSELLING_DEALER                               As New ADODB.Recordset
    Set rsSELLING_DEALER = gconDMIS.Execute("Select * from CSMS_SellingDealer Where DealerCode = '" & XXX & "'")
    If Not rsSELLING_DEALER.EOF And Not rsSELLING_DEALER.BOF Then
        CheckIfValidSellingDealer = True
    Else
        CheckIfValidSellingDealer = False
    End If
    Set rsSELLING_DEALER = Nothing
End Function

Function StoreJobsEntry(ByVal ID As Variant)
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select QUICK_SERVICE,JOBTYPE,TECHCODE,id,LINE_NO,detcde,detdsc,technician,HRSWRK,pocode,wcode,det_amt,discrate,rep_or,livil,detail,discount_2,code from CSMS_RO_Det where id = " & ID)
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        labDetID.Caption = rsRO_DET!ID
        txtJobLineNo.Text = Null2String(rsRO_DET!LINE_NO)
        cboJcode.Text = Null2String(rsRO_DET!DETCDE)
        
        cboJobCode.Text = Null2String(rsRO_DET!DETDSC)
        If cboJobCode.Text = "" Then cboJobCode.Text = Null2String(rsRO_DET!Detail)
        txtJobPostCode.Text = Null2String(rsRO_DET!pocode)
        PREV_LABOR_CHARGE_TO = Null2String(rsRO_DET!wCode)
        If Null2String(rsRO_DET!wCode) <> "" Then
            cboJobChargeTo.Text = Null2String(rsRO_DET!wCode)
            If Null2String(rsRO_DET!wCode) = "C" Or Null2String(rsRO_DET!wCode) = "S" Then
                If SetAcctName(Null2String(rsRO_DET!Code)) = "" Then
                    cboAcctCodeLabor.ListIndex = -1
                    cboAcctCodeLabor.Enabled = False
                Else
                    cboAcctCodeLabor.Text = SetAcctName(Null2String(rsRO_DET!Code))
                    cboAcctCodeLabor.Enabled = True
                End If
            Else
                cboAcctCodeLabor.ListIndex = -1
                cboAcctCodeLabor.Enabled = False
            End If
        Else
            cboJobChargeTo.ListIndex = -1
            cboAcctCodeLabor.ListIndex = -1
            cboAcctCodeLabor.Enabled = False
        End If
        lblTECHCODE_X.Caption = LTrim(RTrim(Null2String(rsRO_DET!TechCode)))

        cboTechnician.Text = setTechnicianName(LTrim(RTrim(Null2String(rsRO_DET!TechCode))))
        txtDET_HRS.Text = N2Str2Zero(rsRO_DET!HRSWRK)
        txtJobRate.Text = N2Str2Zero(rsRO_DET!DET_AMT)
        txtJobDiscount.Text = N2Str2Zero(rsRO_DET!discrate)
        txtJobDiscountAmt.Text = N2Str2Zero(rsRO_DET!Discount_2)
        txtJobDetail.Text = Null2String(rsRO_DET!Detail)

        If Null2String(rsRO_DET!JOBTYPE) = "PMS" Then
            optQUICK.Visible = True
            If Null2String(rsRO_DET!QUICK_SERVICE) = "Y" Then
                optQUICK.Value = 1
            Else
                optQUICK.Value = 0
            End If
        Else
            optQUICK.Visible = False
        End If
    End If
    Set rsRO_DET = Nothing
End Function

Function StorePartsEntry(ByVal ID As Variant)
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select id,LINE_NO,detcde,detdsc,pocode,detvol,detprc,det_amt,wcode,discrate,discount_2,code from CSMS_RO_Det where id = " & ID)
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        labDetID.Caption = rsRO_DET!ID
        txtPartsLineNo.Text = Null2String(rsRO_DET!LINE_NO)
        cboPartNo.Text = Null2String(rsRO_DET!DETCDE)
        cboDescription.Text = SetPartDesc(Null2String(rsRO_DET!DETCDE))
        txtPartCode.Text = Null2String(rsRO_DET!pocode)
        txtQty.Text = N2Str2Zero(rsRO_DET!detvol)
        txtUnitPrice.Text = Null2String(rsRO_DET!DetPrc)
        txtPartAmount.Text = N2Str2Zero(rsRO_DET!DET_AMT)
        PREV_PARTS_CHARGE_TO = Null2String(rsRO_DET!wCode)
        If Null2String(rsRO_DET!wCode) <> "" Then
            cboChargeTo.Text = Null2String(rsRO_DET!wCode)
            If Null2String(rsRO_DET!wCode) = "C" Or Null2String(rsRO_DET!wCode) = "S" Then
                If SetAcctName(Null2String(rsRO_DET!Code)) = "" Then
                    cboAcctCodeParts.ListIndex = -1
                    cboAcctCodeParts.Enabled = False
                Else
                    cboAcctCodeParts.Text = SetAcctName(Null2String(rsRO_DET!Code))
                    cboAcctCodeParts.Enabled = True
                End If
            Else
                cboAcctCodeParts.ListIndex = -1
                cboAcctCodeParts.Enabled = False
            End If
        Else
            cboAcctCodeParts.ListIndex = -1
            cboAcctCodeParts.Enabled = False
            cboChargeTo.ListIndex = -1
        End If
        txtPartDiscount.Text = N2Str2IntZero(rsRO_DET!discrate)
        txtPartDiscountAmt.Text = N2Str2IntZero(rsRO_DET!Discount_2)
        cmdPartsDelete.Enabled = False
    End If
    Set rsRO_DET = Nothing
End Function

Function StoreMatEntry(ByVal ID As Variant)
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select id,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discrate,pocode,discount_2,code,detail from CSMS_RO_Det where id = " & ID)
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        labDetID.Caption = rsRO_DET!ID
        txtMatLineNo.Text = Null2String(rsRO_DET!LINE_NO)
        cboMatCode.Text = Null2String(rsRO_DET!DETCDE)
        cboMaterial.Text = SetMatDisc(Null2String(rsRO_DET!DETCDE))
        txtMatQty.Text = N2Str2Zero(rsRO_DET!detvol)
        txtMatUnitPrice.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!DetPrc))
        txtMatAmount.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!DET_AMT))
        PREV_MATERIALS_CHARGE_TO = Null2String(rsRO_DET!wCode)
        If LOGCODE = "ADM" Or Trim(cboMatCode.Text) = "MISC" Then
            txtMatQty.Enabled = True
            txtMatUnitPrice.Enabled = True
            txtMatAmount.Enabled = False
        Else
            txtMatQty.Enabled = False
            txtMatUnitPrice.Enabled = False
            txtMatAmount.Enabled = False
        End If
        If Null2String(rsRO_DET!Detail) <> "" Then
            Me.txtdetail.Text = Null2String(rsRO_DET!Detail)
        Else
             Me.txtdetail.Text = ""
        End If
        If Null2String(rsRO_DET!wCode) <> "" Then
            cboMatChargeTo.Text = rsRO_DET!wCode
            If Null2String(rsRO_DET!wCode) = "C" Or Null2String(rsRO_DET!wCode) = "S" Then
                If SetAcctName(Null2String(rsRO_DET!Code)) = "" Then
                    cboAcctCodeMaterials.ListIndex = -1
                    cboAcctCodeMaterials.Enabled = False
                Else
                    cboAcctCodeMaterials.Text = SetAcctName(Null2String(rsRO_DET!Code))
                    cboAcctCodeMaterials.Enabled = True
                End If
            Else
                cboAcctCodeMaterials.ListIndex = -1
                cboAcctCodeMaterials.Enabled = False
            End If
        Else
            cboAcctCodeMaterials.ListIndex = -1
            cboAcctCodeMaterials.Enabled = False
            cboMatChargeTo.ListIndex = -1
        End If
        txtMatDiscount.Text = N2Str2Zero(rsRO_DET!discrate)
        txtMatDiscountAmt.Text = N2Str2Zero(rsRO_DET!Discount_2)
        txtMatPOCode.Text = Null2String(rsRO_DET!pocode)
        
        If LTrim(RTrim(cboMatCode.Text)) = "MISC" Then
            cmdMatDelete.Enabled = True
        Else
            cmdMatDelete.Enabled = False
        End If
    End If
    Set rsRO_DET = Nothing
End Function

Function StoreAccEntry(ByVal ID As Variant)
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select id,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discrate,pocode,discount_2,code from CSMS_RO_Det where id = " & ID)
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        labDetID.Caption = rsRO_DET!ID
        txtAccLineNo.Text = Null2String(rsRO_DET!LINE_NO)
        cboAccCode.Text = Null2String(rsRO_DET!DETCDE)
        cboAccessories.Text = SetAccDisc(Null2String(rsRO_DET!DETCDE))
        txtAccQty.Text = N2Str2Zero(rsRO_DET!detvol)
        txtAccUnitPrice.Text = Null2String(rsRO_DET!DetPrc)
        txtAccAmount.Text = N2Str2Zero(rsRO_DET!DET_AMT)
        PREV_ACCESSORIES_CHARGE_TO = Null2String(rsRO_DET!wCode)
        If LOGCODE = "ADM" Then
            txtAccQty.Enabled = True
            txtAccUnitPrice.Enabled = True
            txtAccAmount.Enabled = True
        Else
            txtAccQty.Enabled = False
            txtAccUnitPrice.Enabled = False
            txtAccAmount.Enabled = False
        End If
        If Null2String(rsRO_DET!wCode) <> "" Then
            cboAccChargeTo.Text = Null2String(rsRO_DET!wCode)
            If Null2String(rsRO_DET!wCode) = "C" Or Null2String(rsRO_DET!wCode) = "S" Then
                If SetAcctName(Null2String(rsRO_DET!Code)) = "" Then
                    cboAcctCodeAccessories.ListIndex = -1
                    cboAcctCodeAccessories.Enabled = False
                Else
                    cboAcctCodeAccessories.Text = SetAcctName(Null2String(rsRO_DET!Code))
                    cboAcctCodeAccessories.Enabled = True
                End If
            Else
                cboAcctCodeAccessories.ListIndex = -1
                cboAcctCodeAccessories.Enabled = False
            End If
        Else
            cboAcctCodeAccessories.ListIndex = -1
            cboAcctCodeAccessories.Enabled = False
            cboAccChargeTo.ListIndex = -1
        End If
        txtAccDiscount.Text = N2Str2Zero(rsRO_DET!discrate)
        txtAccDiscountAmt.Text = N2Str2Zero(rsRO_DET!Discount_2)
        cmdAccDelete.Enabled = False
    End If
    Set rsRO_DET = Nothing
End Function

Function StoreParticipationEntry(ByVal RO_NO As String)
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select * from CSMS_Repor where Rep_or = '" & RO_NO & "'")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        chkAllowManDist.Value = 0
        fraParticipation.Enabled = False
        txtLOAAmount.Enabled = True
        txtPartLabor.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!PARTLABOR))
        txtPartParts.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!PARTPARTS))
        txtPartMaterials.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!PARTMATERIALS))
        txtPartAccessories.Text = ToDoubleNumber(N2Str2Zero(rsRO_DET!PARTACCESSORIES))
        SetTotalParticipation
        txtLOAAmount.Text = txtPartTotal.Text
    End If
    Set rsRO_DET = Nothing
End Function

Function SetMatCode(mmm As String)
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        Set rsMatMas = gconDMIS.Execute("select matcde,matdsc from CSMS_MatMas where matdsc = '" & mmm & "'")
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then SetMatCode = Null2String(rsMatMas!MATCDE) Else SetMatCode = cboMatCode.Text
        Set rsMatMas = Nothing
    End If
End Function

Function SetAccCode(mmm As String)
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        Set rsMatMas = gconDMIS.Execute("select STOCKNO,STOCKDESC from PMIS_STOCKMAS where TYPE = 'A' AND STOCKDESC = '" & mmm & "'")
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then SetAccCode = Null2String(rsMatMas!STOCKNO) Else SetAccCode = cboAccCode.Text
        Set rsMatMas = Nothing
    End If
End Function

Function SetMatDisc(mmm As String)
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        Set rsMatMas = gconDMIS.Execute("select matcde,matdsc from CSMS_MatMas where stockno = '" & mmm & "'")
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then SetMatDisc = Null2String(rsMatMas!MatDsc) Else SetMatDisc = cboMaterial.Text
        Set rsMatMas = Nothing
    End If
End Function

Function SetAccDisc(mmm As String)
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        Set rsMatMas = gconDMIS.Execute("select STOCKNO,STOCKDESC from PMIS_STOCKMAS where TYPE = 'A' AND STOCKNO = '" & mmm & "'")
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then SetAccDisc = Null2String(rsMatMas!STOCKDESC) Else SetAccDisc = cboMaterial.Text
        Set rsMatMas = Nothing
    End If
End Function

Function SetAccPrice(mmm As String)
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        Set rsMatMas = gconDMIS.Execute("select STOCKNO,SRP from PMIS_STOCKMAS where TYPE = 'A' AND STOCKNO = '" & mmm & "'")
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then SetAccPrice = Null2String(rsMatMas!SRP) Else SetAccPrice = ""
        Set rsMatMas = Nothing
    End If
End Function

Function SetMatPrice(mmm As String)
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        Set rsMatMas = gconDMIS.Execute("select matcde,s_price from CSMS_MatMas where matcde = '" & mmm & "'")
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then SetMatPrice = Null2String(rsMatMas!s_price) Else SetMatPrice = ""
        Set rsMatMas = Nothing
    End If
End Function

Function SetMatPOCode(mmm As String)
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        Set rsMatMas = gconDMIS.Execute("select matcde,pocode from CSMS_MatMas where matcde = '" & mmm & "'")
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then SetMatPOCode = Null2String(rsMatMas!pocode) Else SetMatPOCode = ""
        Set rsMatMas = Nothing
    End If
End Function

Function SetAccPOCode(mmm As String)
    If mmm <> "" Then
        Set rsMatMas = New ADODB.Recordset
        Set rsMatMas = gconDMIS.Execute("select stockno from PMIS_STOCKMAS where TYPE = 'A' AND STOCKNO = '" & mmm & "'")
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then SetAccPOCode = Null2String(rsMatMas!STOCKNO) Else SetAccPOCode = ""
        Set rsMatMas = Nothing
    End If
End Function

Function SetPartDisc(xx As String)
    If xx <> "" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select partno,partdesc from PMIS_PartMas where partno = '" & xx & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then SetPartDisc = Null2String(rsPartMas!PartDesc) Else SetPartDisc = cboDescription.Text
        Set rsPartMas = Nothing
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set rsPartMas = New ADODB.Recordset
        Set rsPartMas = gconDMIS.Execute("Select partno,srp from PMIS_PartMas where partno = '" & ppp & "'")
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then SetPartPrice = N2Str2Zero(rsPartMas!SRP) Else SetPartPrice = ""
        Set rsPartMas = Nothing
    End If
End Function

Function SetParticipatname(XXX As String)
    Dim rsCustPart                                     As New ADODB.Recordset
    rsCustPart.Open "select cuscde,cusnam from All_Cusmas where cuscde = '" & XXX & "'", gconDMIS
    If Not rsCustPart.EOF And Not rsCustPart.BOF Then
        SetParticipatname = Null2String(rsCustPart!CUSNAM)
    End If
    Set rsCustPart = Nothing
End Function

Function CheckIfAllJobIsFinish() As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    '------------------------------------------------------------
    Dim RSRODET                                        As New ADODB.Recordset
    Dim lng                                            As Long
    Set RSRODET = gconDMIS.Execute("SELECT  DONE FROM CSMS_RO_DET WHERE REP_OR = '" & txtRep_Or & "' AND LIVIL = '1' AND ISNULL(DONE,'') <> 'Y'")
    If (RSRODET.BOF And RSRODET.EOF) Then
        gconDMIS.Execute ("UPDATE CSMS_REPAIRORDER SET JSTATUS = 'F', STATUS = 'Finish Job' where RO_NO = '" & txtRep_Or & "'")
        Call gconDMIS.Execute("UPDATE HRMS_EMPINFO SET JSTATUS='A' , ASSIGNEDRO=NULL  WHERE ASSIGNEDRO=" & N2Str2Null(txtRep_Or.Text), lng)
        If lng = 0 Then
            gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET JSTATUS='A' , ASSIGNEDRO=NULL  WHERE ASSIGNEDRO=" & N2Str2Null(txtRep_Or.Text))
        End If
    End If
    Set RSRODET = Nothing
    '------------------------------------------------------------
    Set rstmp = gconDMIS.Execute("Select STATUS FROM CSMS_RepairOrder Where RO_NO = '" & txtRep_Or.Text & "' and STATUS = 'Finish Job'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        CheckIfAllJobIsFinish = True
    Else
        CheckIfAllJobIsFinish = False
    End If
    Set rstmp = Nothing
End Function

Sub SearchRepairOrder(XXX As String)
    Call rsRefresh
    On Error GoTo ErrorCode
    rsREPOR.Bookmark = rsFind(rsREPOR.Clone, "rep_or", XXX).Bookmark
    Call SendToBack
    labPrevID.Caption = labid.Caption
    Call StoreMemVars
    Exit Sub

ErrorCode:
    ShowCantFind XXX
    Resume Next
End Sub

Sub UpdateParticipation()
    Screen.MousePointer = 11
    SQL_STATEMENT = "Update CSMS_Repor Set " & _
        "PartLabor = " & NumericVal(txtPartLabor) & "," & _
        "PartParts = " & NumericVal(txtPartParts) & "," & _
        "PartMaterials = " & NumericVal(txtPartMaterials) & "," & _
        "Partaccessories = " & NumericVal(txtPartAccessories) & "," & _
        "INSAMT = " & NumericVal(txtPartTotal) & _
        " Where REP_OR = '" & txtRep_Or.Text & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or & " - INSURANCE", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call FillJobs
    Call FillParts
    Call FillMaterials
    Call FillAccessories
    
'   ***** JBF 05/18/2010 *****  change the code for HBI
'    If COMPANY_CODE = "HII" Or COMPANY_CODE = "HBI" Then
'        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
'    Else
'        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(txtPartTotal))
'    End If
'   ***** JBF 05/18/2010 *****
    
     If COMPANY_CODE = "HII" Then
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
     Else
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(txtPartTotal))
     End If
    
    
    
    
''
    If NumericVal(txtPartTotal) = 0 Then
        SQL_STATEMENT = "update CSMS_RepOr set" & _
            " labor = " & TOTJOBAMT - TOTJOBTAX - (NumericVal(txtPartLabor)) & "," & _
            " l_amtvalue = " & Round(TOTJOBAMT, 2) - (NumericVal(txtPartLabor)) & "," & _
            " l_disc = " & Round(TOTJOBDISCVAL, 2) & ", l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
            " l_taxval = " & Round(TOTJOBTAX, 2) & ", l_discount = " & Round(TOTJOBDISC, 2) & "," & _
            " wl_amt = " & 0 & "," & _
            " parts = " & TOTPARTSAMT - TOTPARTSTAX - (NumericVal(txtPartParts)) & "," & _
            " p_amtvalue = " & TOTPARTSAMT - NumericVal(txtPartParts) & "," & _
            " p_disc = " & TOTPARTSDISCVAL & ", p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
            " p_taxval = " & TOTPARTSTAX & ", p_discount = " & TOTPARTSDISC & "," & _
            " wp_amt = " & 0 & "," & _
            " material = " & TOTMATAMT - TOTMATTAX - NumericVal(txtPartMaterials) & "," & _
            " m_amtvalue = " & TOTMATAMT - NumericVal(txtPartMaterials) & "," & _
            " m_disc = " & TOTMATDISCVAL & ", m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
            " m_taxval = " & TOTMATTAX & ", m_discount = " & TOTMATDISC & "," & _
            " Accessories = " & TOTACCAMT - TOTACCTAX - NumericVal(txtPartAccessories) & "," & _
            " A_amtvalue = " & TOTACCAMT - NumericVal(txtPartAccessories) & "," & _
            " A_disc = " & TOTACCDISCVAL & ", A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
            " A_taxval = " & TOTACCTAX & ", A_discount = " & TOTACCDISC & "," & _
            " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
            " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
            " WA_amt = " & 0 & "," & _
            " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
            " where id = " & labid.Caption
    Else
        SQL_STATEMENT = "update CSMS_RepOr set" & _
            " labor = " & TOTJOBAMT - (NumericVal(txtPartLabor)) & "," & _
            " l_amtvalue = " & Round(TOTJOBAMT, 2) - (NumericVal(txtPartLabor)) & "," & _
            " l_disc = " & Round(TOTJOBDISCVAL, 2) & ", l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
            " l_taxval = " & Round(TOTJOBTAX, 2) & ", l_discount = " & Round(TOTJOBDISC, 2) & "," & _
            " wl_amt = " & 0 & "," & _
            " parts = " & TOTPARTSAMT - (NumericVal(txtPartParts)) & "," & _
            " p_amtvalue = " & TOTPARTSAMT - NumericVal(txtPartParts) & "," & _
            " p_disc = " & TOTPARTSDISCVAL & ", p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
            " p_taxval = " & TOTPARTSTAX & ", p_discount = " & TOTPARTSDISC & "," & _
            " wp_amt = " & 0 & "," & _
            " material = " & TOTMATAMT - NumericVal(txtPartMaterials) & "," & _
            " m_amtvalue = " & TOTMATAMT - NumericVal(txtPartMaterials) & "," & _
            " m_disc = " & TOTMATDISCVAL & ", m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
            " m_taxval = " & TOTMATTAX & ", m_discount = " & TOTMATDISC & "," & _
            " Accessories = " & TOTACCAMT - NumericVal(txtPartAccessories) & "," & _
            " A_amtvalue = " & TOTACCAMT - NumericVal(txtPartAccessories) & "," & _
            " A_disc = " & TOTACCDISCVAL & ", A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
            " A_taxval = " & TOTACCTAX & ", A_discount = " & TOTACCDISC & "," & _
            " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
            " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
            " WA_amt = " & 0 & "," & _
            " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
            " where id = " & labid.Caption
    End If
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or & " - INSURANCE PARTICIPATION", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call rsRefresh
    rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    Call StoreMemVars
    Screen.MousePointer = 0
    
    'MessagePop InfoFriend, "Repair order Information", "Insurance amount successfully Set", 1000
    
''
'    SQL_STATEMENT = "update CSMS_RepOr set" & _
'        " labor = " & TOTJOBAMT - TOTJOBTAX - (NumericVal(txtPartLabor)) & "," & _
'        " l_amtvalue = " & Round(TOTJOBAMT, 2) - (NumericVal(txtPartLabor)) & "," & _
'        " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
'        " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
'        " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
'        " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
'        " amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & "," & _
'        " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & "," & _
'        " wl_amt = " & 0 & "," & _
'        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
'        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
'        " wm_amt = " & 0 & "," & _
'        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
'        " where id = " & labID.Caption
'    gconDMIS.Execute SQL_STATEMENT
'    'NEW LOG AUDIT-----------------------------------------------------
'        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - LABOR INSURANCE", "", "")
'    'NEW LOG AUDIT-----------------------------------------------------
'
'    Call FillParts
'    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(txtPartTotal))
'    SQL_STATEMENT = "update CSMS_RepOr set" & _
'        " parts = " & TOTPARTSAMT - TOTPARTSTAX - (NumericVal(txtPartParts)) & "," & _
'        " p_amtvalue = " & TOTPARTSAMT - NumericVal(txtPartParts) & "," & _
'        " p_disc = " & TOTPARTSDISCVAL & "," & _
'        " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
'        " p_taxval = " & TOTPARTSTAX & "," & _
'        " p_discount = " & TOTPARTSDISC & "," & _
'        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
'        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
'        " wp_amt = " & 0 & "," & _
'        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
'        " where id = " & labID.Caption
'    gconDMIS.Execute SQL_STATEMENT
'    'NEW LOG AUDIT-----------------------------------------------------
'        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - PART INSURANCE", "", "")
'    'NEW LOG AUDIT-----------------------------------------------------
'
'    Call FillMaterials
'    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
'    SQL_STATEMENT = "update CSMS_RepOr set" & _
'        " material = " & TOTMATAMT - TOTMATTAX - NumericVal(txtPartMaterials) & "," & _
'        " m_amtvalue = " & TOTMATAMT - NumericVal(txtPartMaterials) & "," & _
'        " m_disc = " & TOTMATDISCVAL & "," & _
'        " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
'        " m_taxval = " & TOTMATTAX & "," & _
'        " m_discount = " & TOTMATDISC & "," & _
'        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
'        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
'        " wm_amt = " & 0 & "," & _
'        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
'        " where id = " & labID.Caption
'    gconDMIS.Execute SQL_STATEMENT
'    'NEW LOG AUDIT-----------------------------------------------------
'        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - MAT INSURANCE", "", "")
'    'NEW LOG AUDIT-----------------------------------------------------
'
'    Call FillAccessories
'    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
'    SQL_STATEMENT = "update CSMS_RepOr set" & _
'        " Accessories = " & TOTACCAMT - TOTACCTAX - NumericVal(txtPartAccessories) & "," & _
'        " A_amtvalue = " & TOTACCAMT - NumericVal(txtPartAccessories) & "," & _
'        " A_disc = " & TOTACCDISCVAL & "," & _
'        " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
'        " A_taxval = " & TOTACCTAX & "," & _
'        " A_discount = " & TOTACCDISC & "," & _
'        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
'        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
'        " WA_amt = " & 0 & "," & _
'        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
'        " where id = " & labID.Caption
'    gconDMIS.Execute SQL_STATEMENT
'    'NEW LOG AUDIT-----------------------------------------------------
'        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labID, "", "RO NO: " & txtRep_Or & " - ACC INSURANCE", "", "")
'    'NEW LOG AUDIT-----------------------------------------------------


End Sub

Sub UpdateROAmount()
    Screen.MousePointer = 11
    Dim TotalCompanyAmount                             As Double
    Dim TotalSalesAmount                               As Double
    Dim TotalWarrantyAmount                            As Double

    Call FillJobs
    TotalCompanyAmount = JobComTotal + PartsComTotal + MatComTotal + ACCComTotal
    TotalSalesAmount = JobSalesTotal + PartsSalesTotal + MatSalesTotal + ACCSalesTotal
    TotalWarrantyAmount = JobWarTotal + PartsWarTotal + MatWarTotal + ACCWarTotal
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
                  " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) - (NumericVal(txtPartLabor)) & "," & _
                  " l_amtvalue = " & Round(TOTJOBAMT, 2) & "," & _
                  " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
                  " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
                  " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
                  " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
                  " amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC + (TotalCompanyAmount + TotalSalesAmount + TotalWarrantyAmount), 2) & "," & _
                  " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & "," & _
                  " wl_amt = " & 0 & "," & _
                  " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                  " where REP_OR = '" & txtRep_Or.Text & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-------------------------------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "RO NO: " & txtRep_Or & " - LABOR", "", "")
    'NEW LOG AUDIT-------------------------------------------------------------------------------------------

    Call FillParts
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
                  " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                  " p_amtvalue = " & TOTPARTSAMT & "," & _
                  " p_disc = " & TOTPARTSDISCVAL & "," & _
                  " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
                  " p_taxval = " & TOTPARTSTAX & "," & _
                  " p_discount = " & TOTPARTSDISC & "," & _
                  " amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC + (TotalCompanyAmount + TotalSalesAmount + TotalWarrantyAmount), 2) & "," & _
                  " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                  " wp_amt = " & 0 & "," & _
                  " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-------------------------------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "RO NO: " & txtRep_Or & " - PARTS", "", "")
    'NEW LOG AUDIT-------------------------------------------------------------------------------------------

    Call FillMaterials
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
                  " material = " & TOTMATAMT - TOTMATTAX & "," & _
                  " m_amtvalue = " & TOTMATAMT & "," & _
                  " m_disc = " & TOTMATDISCVAL & "," & _
                  " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
                  " m_taxval = " & TOTMATTAX & "," & _
                  " m_discount = " & TOTMATDISC & "," & _
                  " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC + (TotalCompanyAmount + TotalSalesAmount + TotalWarrantyAmount) & "," & _
                  " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                  " wm_amt = " & 0 & "," & _
                  " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-------------------------------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "RO NO: " & txtRep_Or & " - MATERIALS", "", "")
    'NEW LOG AUDIT-------------------------------------------------------------------------------------------

    Call FillAccessories
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
                  " Accessories = " & TOTACCAMT - TOTACCTAX & "," & _
                  " A_amtvalue = " & TOTACCAMT & "," & _
                  " A_disc = " & TOTACCDISCVAL & "," & _
                  " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
                  " A_taxval = " & TOTACCTAX & "," & _
                  " A_discount = " & TOTACCDISC & "," & _
                  " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC + (TotalCompanyAmount + TotalSalesAmount + TotalWarrantyAmount) & "," & _
                  " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                  " WA_amt = " & 0 & "," & _
                  " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-------------------------------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "RO NO: " & txtRep_Or & " - ACCESSORIES", "", "")
    'NEW LOG AUDIT-------------------------------------------------------------------------------------------

    Call rsRefresh
    rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    Call StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub PRINTRODISCOUNT()
    Screen.MousePointer = 11

    rptRepairOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptRepairOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If COMPANY_CODE = "HBK" Then: rptRepairOrder.Formulas(3) = "taym= '" & GetTaym & "'"
    If COMPANY_CODE = "HGC" Then
        If REPRINT_CAPTION = "YES" Then
            rptRepairOrder.Formulas(3) = "REPRINT = 'RE-PRINT'"
        Else
            rptRepairOrder.Formulas(3) = "REPRINT = ''"
        End If
        REPRINT_CAPTION = ""
    End If
    
    PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorderdisc.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
     If COMPANY_CODE = "HSR" Then
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorderdisc_summary.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
     End If
     If COMPANY_CODE = "HLI" Then
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "Service_Slip.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
     End If

End Sub

Sub PRINTROWCODE()
    Screen.MousePointer = 11
    rptRepairOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptRepairOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If COMPANY_CODE = "HBK" Then: rptRepairOrder.Formulas(3) = "taym= '" & GetTaym & "'"
    If COMPANY_CODE = "HGC" Then
        If REPRINT_CAPTION = "YES" Then
            rptRepairOrder.Formulas(3) = "REPRINT = 'RE-PRINT'"
        Else
            rptRepairOrder.Formulas(3) = "REPRINT = ''"
        End If
        REPRINT_CAPTION = ""
    End If
    
    
    PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorderdisc2.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
     If COMPANY_CODE = "HSR" Then
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorderdisc2_summary.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
     End If
     If COMPANY_CODE = "HLI" Then
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "Service_Slip.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
     End If
End Sub

Sub PRINTRO()
    Screen.MousePointer = 11
    rptRepairOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptRepairOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If COMPANY_CODE = "HBK" Then: rptRepairOrder.Formulas(3) = "taym= '" & GetTaym & "'"
    If COMPANY_CODE = "HGC" Then
        If REPRINT_CAPTION = "YES" Then
            rptRepairOrder.Formulas(3) = "REPRINT = 'RE-PRINT'"
        Else
            rptRepairOrder.Formulas(3) = "REPRINT = ''"
        End If
        REPRINT_CAPTION = ""
    End If

    If COMPANY_CODE = "HPC" Then
        If NumericVal(rsREPOR!INSAMT) > 0 Then
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorder_INS.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorder.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
        End If
    Else
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorder.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    End If
    If COMPANY_CODE = "HSR" Then
       PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorder_summary.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    End If
    If COMPANY_CODE = "HLI" Then
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "Service_Slip.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
     End If
    Screen.MousePointer = 0
End Sub

Sub PRINTROWSC()
    Screen.MousePointer = 11
    If txtPlate_No.Text = "000000" Then
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorder3.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    
         If COMPANY_CODE = "HLI" Then
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "Service_Slip.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
         End If
    Else
        rptRepairOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptRepairOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        If COMPANY_CODE = "HBK" Then: rptRepairOrder.Formulas(3) = "taym= '" & GetTaym & "'"
        If COMPANY_CODE = "HGC" Then
            If REPRINT_CAPTION = "YES" Then
                rptRepairOrder.Formulas(3) = "REPRINT = 'RE-PRINT'"
            Else
                rptRepairOrder.Formulas(3) = "REPRINT = ''"
            End If
            REPRINT_CAPTION = ""
        End If
                
        If COMPANY_CODE = "HQA" Then
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorder3_PMS.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorder3.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
        End If
        If COMPANY_CODE = "HSR" Then
           PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "repairorder3_summary.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
        End If
        If COMPANY_CODE = "HLI" Then
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "Service_Slip.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
        End If
    End If
    Screen.MousePointer = 0
End Sub

Function CheckIfThereAPMS(vRO As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT JOBTYPE FROM CSMS_RO_DET WHERE REP_OR = '" & vRO & "' AND LIVIL = '1' AND JOBTYPE = 'PMS'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        CheckIfThereAPMS = True
    Else
        CheckIfThereAPMS = False
    End If

    Set rstmp = Nothing
End Function

Sub PrintWARRANTY()
'    If SSTab1.TabEnabled(2) = True Then
'        If MsgQuestionBox("Print this Repair Order on Warranty Format?", "Print Repair Order for Warranty") = True Then
'            If txtInvoiceNo.Text = "" Then
'                MsgBox "Transaction must be Billed First.", vbInformation, "Cannot Print this Transaction..."
'                Exit Sub
'            Else
'                If DiscTotal > 0 And (WARTotal > 0 Or SALESTotal > 0 Or COMTotal > 0) Then
'                    WARRANTY_PRINTROWCODE
'                ElseIf DiscTotal > 0 And (WARTotal = 0 And SALESTotal = 0 And COMTotal = 0) Then
'                    WARRANTY_PRINTRODISCOUNT
'                ElseIf DiscTotal = 0 And (WARTotal > 0 Or SALESTotal > 0 Or COMTotal > 0) Then
'                    WARRANTY_PRINTROWSC
'                Else
'                    WARRANTY_PRINTRO
'                End If
'            End If
'        End If
'    Else
'        MsgSpeechBox "Print Repair Order Not yet Implemented!"
'    End If
End Sub

Sub WARRANTY_PRINTRODISCOUNT()
    Screen.MousePointer = 11
    PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "WARRANTYrepairorderdisc.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub WARRANTY_PRINTROWCODE()
    Screen.MousePointer = 11
    PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "WARRANTYrepairorderdisc2.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub WARRANTY_PRINTRO()
    Screen.MousePointer = 11
    PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "WARRANTYrepairorder.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub WARRANTY_PRINTROWSC()
    Screen.MousePointer = 11
    If txtPlate_No.Text = "000000" Then
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "WARRANTYrepairorder3.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    Else
        PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "WARRANTYrepairorder2.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
    End If
    Screen.MousePointer = 0
End Sub

Sub SetROTransToZeroRatedVat(XXX As String)
    Dim rsRO_DET                                       As New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_Ro_Det Where Rep_or = '" & XXX & "' Order by Livil,Line_No asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If COMPANY_CODE = "HPI" Then
                SQL_STATEMENT = "Update CSMS_RO_DET Set " & _
                    " TAXRATE = 0, " & _
                    " TAXVAL = 0 " & _
                    " Where ID = " & rsRO_DET!ID
            Else
                SQL_STATEMENT = "Update CSMS_RO_DET Set " & _
                    " DET_AMT = (DET_AMT  / " & ConvertToBIRDecimalFormat(VAT_RATE) & ") " & _
                    ", TAXRATE = 0 " & _
                    ", TAXVAL = 0 " & _
                    " Where ID = " & rsRO_DET!ID
            End If
            gconDMIS.Execute SQL_STATEMENT

            'NEW LOG AUDIT---------------------------------------------------------------------------------------
                Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or & " - SET TO ZERO RATED", "", Null2String(rsRO_DET!ID))
            'NEW LOG AUDIT---------------------------------------------------------------------------------------
            rsRO_DET.MoveNext
        Loop
    End If
End Sub

Sub SetROTransToNonZeroRatedVat(XXX As String)
    Dim vDiscKRate                                      As Double
    Dim VDiscKVal                                       As Double
    Dim rsRO_DET                                        As New ADODB.Recordset
    
    Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_Ro_Det Where Rep_or = '" & XXX & "' Order by Livil,Line_No asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If COMPANY_CODE = "HPI" Then
                'UPDATED BY: JUN
                'DATE UPDATED: 03-20-2009
                    SQL_STATEMENT = "Update CSMS_RO_DET Set " & _
                        " TAXRATE = " & (VAT_RATE / 100) & _
                        ", TAXVAL = (((DET_AMT - DISCOUNT_2) / 1.12) * " & (VAT_RATE / 100) & ") " & _
                        " Where ID = " & rsRO_DET!ID
                'UPDATED BY: JUN
                
            Else
                SQL_STATEMENT = "Update CSMS_RO_DET Set " & _
                    " DET_AMT = DET_AMT + (DET_AMT * 0.12) " & _
                    ", TAXRATE = " & (VAT_RATE / 100) & _
                    ", TAXVAL = ((DETPRC - DISCOUNT_2) * " & (VAT_RATE / 100) & ") " & _
                    " Where ID = " & rsRO_DET!ID
            End If
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT---------------------------------------------------------------------------------------
                Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or & " - SET TO NON ZERO RATED", "", Null2String(rsRO_DET!ID))
            'NEW LOG AUDIT---------------------------------------------------------------------------------------
            rsRO_DET.MoveNext
        Loop
    End If
    
    Call FillJobs
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
        " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) & "," & _
        " l_amtvalue = " & Round(TOTJOBAMT, 2) & "," & _
        " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
        " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
        " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
        " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
        " amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & "," & _
        " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & "," & _
        " wl_amt = " & 0 & "," & _
        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
        " where REP_OR = '" & txtRep_Or.Text & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT--------------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or & " - SET TO NON TO ZERO", "", "")
    'NEW LOG AUDIT--------------------------------------------------------------------------
    Exit Sub

ErrorCode:
    MsgBox Err.Description
End Sub

Sub rsRefresh()
    Set rsREPOR = New ADODB.Recordset
    Set rsREPOR = gconDMIS.Execute("select * from CSMS_RepOr WHERE TRANSTYPE = 'R' AND ID = " & labid.Caption & "")
End Sub

Sub initMemvars()
    txtEstimateno.Text = "": txtInvoiceNo.Text = ""
    txtRep_Or.Text = "": txtROType.Text = "0"
    txtSvc_No.Text = "Q": txtAcct_No.Text = ""
    txtNiym.Text = "": txtPlate_No.Text = ""
    txtMake.Text = "": txtTerm.Text = "CSH"
    txtSektion.Text = "": txtKm_rdg.Text = ""
    txtDte_recd.Value = LOGDATE
    txtCertific8.Text = "": txtDte_comp.Text = ""
    txtDte_Rel.Text = "": txtVIN.Text = ""
    txtParticipat.Text = ""
    JobTotal = 0: JobWarTotal = 0: JobDiscTotal = 0: JobVatTotal = 0
    PartsTotal = 0: PartsWarTotal = 0: PartsDiscTotal = 0: PartsVatTotal = 0
    MatTotal = 0: MatWarTotal = 0: MatDiscTotal = 0: MatVatTotal = 0
    ROTotal = 0: Pcnt = 0: kcnt = 0: Mcnt = 0
    Call clearDetailsgrd
    Call InitGrid
    lstJObs.Sorted = True: lstJObs.ListItems.Clear: lstJObs.Refresh
    lstParts.Sorted = True: lstParts.ListItems.Clear: lstParts.Refresh
    lstMaterials.Sorted = True: lstMaterials.ListItems.Clear: lstMaterials.Refresh
End Sub

Sub InitMaterials()
    cboMatCode.Text = "MISC"
    cboMatCode.Enabled = False
    cboMaterial.Text = "MISCELLANEOUS CHARGES"
    cboMaterial.Enabled = False
    txtMatLineNo.Text = Format(Mcnt + 1, "00")
    txtMatQty.Text = 1
    txtMatUnitPrice.Text = ZERO
    txtMatAmount.Text = ZERO
    cboMatChargeTo.Clear
    cboMatChargeTo.AddItem ""
    cboMatChargeTo.AddItem "W"
    cboMatChargeTo.AddItem "S"
    cboMatChargeTo.AddItem "C"
    cboMatChargeTo.ListIndex = -1
    txtMatDiscount.Text = ZERO
    txtMatPOCode.Text = "01"
    PREV_MATERIALS_CHARGE_TO = ""
    txtMatQty.Enabled = True
    txtMatUnitPrice.Enabled = True
End Sub

Sub InitAccessories()
    cboAccCode.Text = ""
    cboAccessories.Text = ""
    txtAccLineNo.Text = Format(Mcnt + 1, "00")
    txtAccQty.Text = 1
    txtAccUnitPrice.Text = ZERO
    txtAccAmount.Text = ZERO
    cboAccChargeTo.Clear
    cboAccChargeTo.AddItem ""
    cboAccChargeTo.AddItem "W"
    cboAccChargeTo.AddItem "S"
    cboAccChargeTo.AddItem "C"
    cboAccChargeTo.ListIndex = -1
    txtAccDiscount.Text = ZERO
    txtMatPOCode.Text = "01"
    PREV_ACCESSORIES_CHARGE_TO = ""
End Sub

Sub InitParts()
    txtPartsLineNo.Text = Format(Pcnt + 1, "00")
    cboPartNo.ListIndex = -1
    cboDescription.ListIndex = -1
    txtPartCode.Text = "01"
    txtQty.Text = 1
    txtUnitPrice.Text = ZERO
    txtPartAmount.Text = ZERO
    cboChargeTo.Clear
    cboChargeTo.AddItem ""
    cboChargeTo.AddItem "W"
    cboChargeTo.AddItem "S"
    cboChargeTo.AddItem "C"
    cboChargeTo.ListIndex = -1
    cboAcctCodeParts.Enabled = False
    cboAcctCodeParts.ListIndex = -1
    txtPartDiscount.Text = ZERO
    PREV_PARTS_CHARGE_TO = ""
End Sub

Sub ImportParts()
    Dim RONOformat                                     As String
    Dim yza                                            As Integer
    Dim tisoy, keikei                                  As String
    RONOformat = ""

    keikei = "": tisoy = "": yza = 0
    For yza = 1 To Len(rsREPOR!REP_OR)
        tisoy = Mid(rsREPOR!REP_OR, yza, 1)
        keikei = keikei + tisoy
    Next
    RONOformat = keikei
    
    Dim VarPartsLINE_NO                                 As String
    Dim VarPartNo                                       As String
    Dim VarDescription                                  As String
    Dim VarPartCode                                     As String
    Dim VarQty                                          As Double
    Dim VarUnitCost                                     As Double
    Dim VarUnitPrice                                    As Double
    Dim VarPartAmount                                   As String
    Dim VarChargeTo                                     As String
    Dim VarPartDiscount                                 As String
    Dim PARTSREP_OR                                     As String
    Dim PARTSLEVEL                                      As String
    Dim PARTSLINE_NO                                    As String
    Dim PARTSDETCDE                                     As String
    Dim PARTSDETDSC                                     As String
    Dim PARTSDETUNT                                     As String
    Dim PARTSDETVOL                                     As Double
    Dim PARTSDETPRC                                     As Double
    Dim PARTSDETAMT                                     As Double
    Dim PARTSCODE                                       As String
    Dim PARTSWCODE                                      As String
    Dim PARTSTAXRATE                                    As Double
    Dim PARTSDISCRATE                                   As Double
    Dim PARTSTAXVAL                                     As Double
    Dim PARTSDISVAL                                     As Double
    Dim PARTSPOCODE                                     As String
    Dim PARTSRep_Or2                                    As String
    Dim PARTSDETAIL                                     As String
    Dim PARTSDET_AMT                                    As Double
    Dim PARTSDETCOST                                    As Double
    Dim PARTSDIS_VAL                                    As Double
    Dim PARTSDISCOUNT_2                                 As Double
    Dim PARTSREMARKS                                    As String
    Dim REF_RIV_ADB                                     As String
    Dim rsRR_HDCheck                                    As ADODB.Recordset
    Dim rsRR_HDTdaytranCheck                            As ADODB.Recordset
    vPIS_NO_CHARGE_TO = "NULL"
    Dim vGJorBP                                         As String
    vGJorBP = "NULL"

    'gconDMIS.Execute "delete from CSMS_RO_Det where livil = '2' AND ISNULL(jobtype,'') <> 'SR' AND rep_or = " & N2Str2Null(txtRep_Or.Text)

    SQL_STATEMENT = "delete from CSMS_RO_Det where livil = '2' AND rep_or = " & N2Str2Null(txtRep_Or.Text)
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT---------------------------------------------------------------------
    Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, labid, "P", "RO NO: " & txtRep_Or & " - IMPORT PARTS", "", "")
    'NEW LOG AUDIT---------------------------------------------------------------------

    'UPDATED BY: JUN--------IMPORT PARTS FROM SUBLET------------------------------------
    'DATE: 06/05/2008
    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HAS" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HCO" Or COMPANY_CODE = "HPC" Then
        Call ImportPartsFromSublet(txtRep_Or.Text)
    End If
    '-----------------------------------------------------------------------------------
    Pcnt = 0
    RO_RIV_Tranno_Counter = 0
    Set rsORD_HIST = New ADODB.Recordset
    Set rsORD_HIST = gconDMIS.Execute("select rono,trandate,trantype,tranno,REFPISNO from PMIS_ord_hist where TRANTYPE = 'RIV' AND [TYPE] = 'P' AND isnull(status,'') <> 'C' and isnull(status,'')  <> 'N' and rono = '" & RONOformat & "'")
    If Not rsORD_HIST.EOF And Not rsORD_HIST.BOF Then
        rsORD_HIST.MoveFirst
        Do While Not rsORD_HIST.EOF
            If Mid(Null2String(rsORD_HIST!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(rsORD_HIST!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(rsORD_HIST!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(rsORD_HIST!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
            If Mid(Null2String(rsORD_HIST!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

            Set rsDAYTRAN = New ADODB.Recordset
            Set rsDAYTRAN = gconDMIS.Execute("select itemno,trantype,tranno,stock_ord,stock_sup,tranqty,tranucost,tranuprice from PMIS_DayTran where [TYPE] = 'P' AND trantype = 'RIV' and tranno = " & N2Str2Null(rsORD_HIST!TRANNO) & " order by itemno asc")
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                Screen.MousePointer = 11
                rsDAYTRAN.MoveFirst
                RO_RIV_Tranno_Counter = RO_RIV_Tranno_Counter + 1
                RO_RIV_Tranno(RO_RIV_Tranno_Counter) = Null2String(rsORD_HIST!TRANNO)
                Do While Not rsDAYTRAN.EOF
                    Pcnt = Pcnt + 1
                    VarPartsLINE_NO = "": VarPartNo = "": VarDescription = ""
                    VarPartCode = "": VarQty = 0: VarUnitPrice = 0
                    VarPartAmount = "": VarChargeTo = " ": VarPartDiscount = ZERO

                    VarPartsLINE_NO = Format(Pcnt, "00")
                    VarPartNo = Null2String(rsDAYTRAN!STOCK_ORD)
                    VarDescription = SetPartDesc(Null2String(rsDAYTRAN!STOCK_SUP))
                    VarPartCode = "01"
                    VarQty = Format(N2Str2IntZero(rsDAYTRAN!tranqty), "####0")
                    VarUnitCost = N2Str2Zero(rsDAYTRAN!TRANUCOST)
                    VarUnitPrice = N2Str2Zero(rsDAYTRAN!TRANUPRICE)
                    VarPartAmount = N2Str2Zero(rsDAYTRAN!tranqty) * N2Str2Zero(rsDAYTRAN!TRANUPRICE)
                    VarChargeTo = " "
                    VarPartDiscount = ZERO

                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'P' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsORD_HIST!TRANNO), "000000") & "' AND STATUS = 'P'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Do While Not rsRR_HDCheck.EOF
                            Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                            Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'P' AND STATUS = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                            If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                                rsRR_HDTdaytranCheck.MoveFirst
                                Do While Not rsRR_HDTdaytranCheck.EOF
                                    If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarPartNo) Then
                                        MsgBox "Warning: Part Number : " & VarPartNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                        Pcnt = Pcnt - 1
                                        GoTo 10000
                                    End If
                                    rsRR_HDTdaytranCheck.MoveNext
                                Loop
                            End If
                            rsRR_HDCheck.MoveNext
                        Loop
                    End If
                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'P' AND STATUS = 'P' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsORD_HIST!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Do While Not rsRR_HDCheck.EOF
                            Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                            Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'P' AND STATUS = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                            If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                                rsRR_HDTdaytranCheck.MoveFirst
                                Do While Not rsRR_HDTdaytranCheck.EOF
                                    If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarPartNo) Then
                                        MsgBox "Warning: Part Number : " & VarPartNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                        Pcnt = Pcnt - 1
                                        GoTo 10000
                                    End If
                                    rsRR_HDTdaytranCheck.MoveNext
                                Loop
                            End If
                            rsRR_HDCheck.MoveNext
                        Loop
                    End If
                    REF_RIV_ADB = "'RIV" & Format(Null2String(rsDAYTRAN!TRANNO), "000000") & Format(Null2String(rsDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                    PARTSREP_OR = N2Str2Null(txtRep_Or.Text)
                    PARTSLEVEL = "'2'"
                    PARTSLINE_NO = N2Str2Null(Format(VarPartsLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(VarPartNo)
                    PARTSDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VarQty)
                    PARTSDETCOST = NumericVal(VarUnitCost)
                    PARTSDETPRC = NumericVal(VarUnitPrice)
                    PARTSDETAMT = Round(NumericVal(VarPartAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = vPIS_NO_CHARGE_TO
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = NumericVal(VarPartDiscount) / 100
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VarPartCode)
                    PARTSRep_Or2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VarPartAmount)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    gconDMIS.Execute "insert into CSMS_RO_Det " & _
                                     "(TRANSTYPE,rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                   " values ('R'," & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                   " " & vGJorBP & "," & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                   " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                   " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                     ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                     ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                     ", " & PARTSRep_Or2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                     ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    Screen.MousePointer = 0
10000               rsDAYTRAN.MoveNext
                Loop
            End If
            rsORD_HIST.MoveNext
        Loop
    End If

    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hd where [TYPE] = 'P' AND status <> 'C' and status <> 'N' and trantype = 'RIV' and rono = '" & RONOformat & "'")
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        Do While Not rsOrd_Hd.EOF
            If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(rsOrd_Hd!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
            If Mid(Null2String(rsOrd_Hd!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

            Set rsTdayTran = New ADODB.Recordset
            Set rsTdayTran = gconDMIS.Execute("select itemno,trantype,tranno,STOCK_ord,STOCK_sup,tranqty,tranucost,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND trantype = 'RIV' and tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " order by itemno asc")
            If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                Screen.MousePointer = 11
                rsTdayTran.MoveFirst
                RO_RIV_Tranno_Counter = RO_RIV_Tranno_Counter + 1
                RO_RIV_Tranno(RO_RIV_Tranno_Counter) = Null2String(rsOrd_Hd!TRANNO)
                Do While Not rsTdayTran.EOF
                    Pcnt = Pcnt + 1
                    VarPartsLINE_NO = "": VarPartNo = "": VarDescription = ""
                    VarPartCode = "": VarQty = 0: VarUnitPrice = 0
                    VarPartAmount = "": VarChargeTo = " ": VarPartDiscount = ZERO

                    VarPartsLINE_NO = Format(Pcnt, "00")
                    VarPartNo = Null2String(rsTdayTran!STOCK_ORD)
                    VarDescription = SetPartDesc(Null2String(rsTdayTran!STOCK_SUP))
                    VarPartCode = "01"
                    VarQty = Format(N2Str2IntZero(rsTdayTran!tranqty), "####0")
                    VarUnitCost = N2Str2Zero(rsTdayTran!TRANUCOST)
                    VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                    VarPartAmount = N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE)
                    VarChargeTo = " "
                    VarPartDiscount = ZERO
                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'P' AND STATUS = 'P' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsOrd_Hd!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Do While Not rsRR_HDCheck.EOF
                            Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                            Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'P' AND STATUS = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                            If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                                rsRR_HDTdaytranCheck.MoveFirst
                                Do While Not rsRR_HDTdaytranCheck.EOF
                                    If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarPartNo) Then
                                        MsgBox "Warning: Part Number : " & VarPartNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                        Pcnt = Pcnt - 1
                                        GoTo 20000
                                    End If
                                    rsRR_HDTdaytranCheck.MoveNext
                                Loop
                            End If
                            rsRR_HDCheck.MoveNext
                        Loop
                    End If
                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'P' AND STATUS = 'P' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsOrd_Hd!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Do While Not rsRR_HDCheck.EOF
                            Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                            Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'P' AND STATUS = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                            If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                                rsRR_HDTdaytranCheck.MoveFirst
                                Do While Not rsRR_HDTdaytranCheck.EOF
                                    If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarPartNo) Then
                                        MsgBox "Warning: Part Number : " & VarPartNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                        Pcnt = Pcnt - 1
                                        GoTo 20000
                                    End If
                                    rsRR_HDTdaytranCheck.MoveNext
                                Loop
                            End If
                            rsRR_HDCheck.MoveNext
                        Loop
                    End If
                    REF_RIV_ADB = "'RIV" & Format(Null2String(rsTdayTran!TRANNO), "000000") & Format(Null2String(rsTdayTran!itemno), "000") & "'"
                    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                    PARTSREP_OR = N2Str2Null(txtRep_Or.Text)
                    PARTSLEVEL = "'2'"
                    PARTSLINE_NO = N2Str2Null(Format(VarPartsLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(VarPartNo)
                    PARTSDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VarQty)
                    PARTSDETCOST = NumericVal(VarUnitCost)
                    PARTSDETPRC = NumericVal(VarUnitPrice)
                    PARTSDETAMT = Round(NumericVal(VarPartAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = vPIS_NO_CHARGE_TO
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = Round(NumericVal(VarPartDiscount) / 100, 2)
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VarPartCode)
                    PARTSRep_Or2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VarPartAmount)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    gconDMIS.Execute "insert into CSMS_RO_Det " & _
                                     "(TRANSTYPE,rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                   " values ('R'," & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                   " " & vGJorBP & "," & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                   " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                   " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                     ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                     ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                     ", " & PARTSRep_Or2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                     ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    Screen.MousePointer = 0
20000               rsTdayTran.MoveNext
                Loop
            End If
            rsOrd_Hd.MoveNext
        Loop
    End If

    'PART ADVANCE BILL
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hd where [TYPE] = 'P' AND status <> 'C' AND STATUS <> 'N' and trantype = 'ADB' and rono = '" & RONOformat & "'")
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        If MsgQuestionBox("Advance Bill for Repair Order: " & txtRep_Or.Text & " is Available " & vbCrLf & _
                          "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then
            Do While Not rsOrd_Hd.EOF
                If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
                If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
                If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

                If Mid(Null2String(rsOrd_Hd!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
                If Mid(Null2String(rsOrd_Hd!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

                Set rsTdayTran = New ADODB.Recordset
                Set rsTdayTran = gconDMIS.Execute("select itemno,trantype,tranno,stock_ord,stock_sup,tranqty,tranucost,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND trantype = 'ADB' and tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " order by itemno asc")
                If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                    Screen.MousePointer = 11
                    rsTdayTran.MoveFirst
                    Do While Not rsTdayTran.EOF
                        Pcnt = Pcnt + 1
                        VarPartsLINE_NO = "": VarPartNo = "": VarDescription = ""
                        VarPartCode = "": VarQty = 0: VarUnitPrice = 0
                        VarPartAmount = "": VarChargeTo = " ": VarPartDiscount = ZERO

                        VarPartsLINE_NO = Format(Pcnt, "00")
                        VarPartNo = Null2String(rsTdayTran!STOCK_ORD)
                        VarDescription = Null2String(rsTdayTran!STOCK_SUP)
                        VarPartCode = "01"
                        VarQty = Format(N2Str2IntZero(rsTdayTran!tranqty), "####0")
                        VarUnitCost = N2Str2Zero(rsTdayTran!TRANUCOST)
                        VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                        VarPartAmount = N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE)
                        VarChargeTo = " "
                        VarPartDiscount = ZERO

                        PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                        PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                        REF_RIV_ADB = "'ADB" & Format(Null2String(rsTdayTran!TRANNO), "000000") & Format(Null2String(rsTdayTran!itemno), "000") & "'"
                        PARTSREP_OR = N2Str2Null(txtRep_Or.Text)
                        PARTSLEVEL = "'2'"
                        PARTSLINE_NO = N2Str2Null(Format(VarPartsLINE_NO, "00"))
                        PARTSDETCDE = N2Str2Null(VarPartNo)
                        PARTSDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                        PARTSDETUNT = "NULL"
                        PARTSDETVOL = N2Str2Zero(VarQty)
                        PARTSDETPRC = NumericVal(VarUnitPrice)
                        PARTSDETAMT = Round(NumericVal(VarPartAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                        PARTSCODE = "NULL"
                        PARTSWCODE = vPIS_NO_CHARGE_TO
                        PARTSTAXRATE = (VAT_RATE / 100)
                        PARTSDISCRATE = NumericVal(VarPartDiscount) / 100
                        PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                        PARTSPOCODE = N2Str2Null(VarPartCode)
                        PARTSRep_Or2 = "NULL"
                        PARTSDETAIL = "NULL"
                        PARTSDET_AMT = NumericVal(VarPartAmount)
                        PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                        PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                        PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                        gconDMIS.Execute "insert into CSMS_RO_Det " & _
                                         "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                       " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                       " " & vGJorBP & "," & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                       " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                       " " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                         ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                         ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                         ", " & PARTSRep_Or2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                         ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                        Screen.MousePointer = 0
                        rsTdayTran.MoveNext
                    Loop
                End If
                rsOrd_Hd.MoveNext
            Loop
        End If
    End If
    
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hist where [TYPE] = 'P' AND status <> 'C' AND STATUS <> 'N' and trantype = 'ADB' and rono = '" & RONOformat & "'")
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        If MsgQuestionBox("Advance Bill for Repair Order: " & txtRep_Or.Text & " is Available " & vbCrLf & _
                          "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then
            Do While Not rsOrd_Hd.EOF
                If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
                If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
                If Mid(Null2String(rsOrd_Hd!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

                If Mid(Null2String(rsOrd_Hd!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
                If Mid(Null2String(rsOrd_Hd!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

                Set rsTdayTran = New ADODB.Recordset
                Set rsTdayTran = gconDMIS.Execute("select itemno,trantype,tranno,stock_ord,stock_sup,tranqty,tranuprice from PMIS_DayTran where [TYPE] = 'P' AND trantype = 'ADB' and tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " order by itemno asc")
                If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                    Screen.MousePointer = 11
                    rsTdayTran.MoveFirst
                    Do While Not rsTdayTran.EOF
                        Pcnt = Pcnt + 1
                        VarPartsLINE_NO = "": VarPartNo = "": VarDescription = ""
                        VarPartCode = "": VarQty = 0: VarUnitPrice = 0
                        VarPartAmount = "": VarChargeTo = " ": VarPartDiscount = ZERO

                        VarPartsLINE_NO = Format(Pcnt, "00")
                        VarPartNo = Null2String(rsTdayTran!STOCK_ORD)
                        VarDescription = Null2String(rsTdayTran!STOCK_SUP)
                        VarPartCode = "01"
                        VarQty = Format(N2Str2IntZero(rsTdayTran!tranqty), "####0")
                        VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                        VarPartAmount = N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE)
                        VarChargeTo = " "
                        VarPartDiscount = ZERO

                        PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                        PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                        REF_RIV_ADB = "'ADB" & Format(Null2String(rsTdayTran!TRANNO), "000000") & Format(Null2String(rsTdayTran!itemno), "000") & "'"
                        PARTSREP_OR = N2Str2Null(txtRep_Or.Text)
                        PARTSLEVEL = "'2'"
                        PARTSLINE_NO = N2Str2Null(Format(VarPartsLINE_NO, "00"))
                        PARTSDETCDE = N2Str2Null(VarPartNo)
                        PARTSDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                        PARTSDETUNT = "NULL"
                        PARTSDETVOL = N2Str2Zero(VarQty)
                        PARTSDETPRC = NumericVal(VarUnitPrice)
                        PARTSDETAMT = Round(NumericVal(VarPartAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                        PARTSCODE = "NULL"
                        PARTSWCODE = vPIS_NO_CHARGE_TO
                        PARTSTAXRATE = (VAT_RATE / 100)
                        PARTSDISCRATE = NumericVal(VarPartDiscount) / 100
                        PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                        PARTSPOCODE = N2Str2Null(VarPartCode)
                        PARTSRep_Or2 = "NULL"
                        PARTSDETAIL = "NULL"
                        PARTSDET_AMT = NumericVal(VarPartAmount)
                        PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                        PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                        PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                        gconDMIS.Execute "insert into CSMS_RO_Det " & _
                                         "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                       " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                       " " & vGJorBP & "," & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                       " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                       " " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                         ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                         ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                         ", " & PARTSRep_Or2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                         ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                        Screen.MousePointer = 0
                        rsTdayTran.MoveNext
                    Loop
                End If
                rsOrd_Hd.MoveNext
            Loop
        End If
    End If

    If chkParticipat.Value = 0 Then
        FillParts
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        gconDMIS.Execute "update CSMS_RepOr set" & _
                       " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                       " p_amtvalue = " & TOTPARTSAMT & "," & _
                       " p_disc = " & TOTPARTSDISCVAL & "," & _
                       " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
                       " p_taxval = " & TOTPARTSTAX & "," & _
                       " p_discount = " & TOTPARTSDISC & "," & _
                       " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                       " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                       " wp_amt = " & 0 & "," & _
                       " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                       " where id = " & labid.Caption
        rsRefresh
        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
        cmdPartsCancel.Value = True
    End If
    Screen.MousePointer = 0
    Exit Sub


ErrorCode:
    Screen.MousePointer = 0: ShowVBError
End Sub

Sub ImportMaterials()
    Dim RONOformat                                     As String
    Dim yza                                            As Integer
    Dim tisoy, keikei                                  As String
    RONOformat = ""

    keikei = "": tisoy = "": yza = 0
    For yza = 1 To Len(rsREPOR!REP_OR)
        tisoy = Mid(rsREPOR!REP_OR, yza, 1)

        keikei = keikei + tisoy
    Next
    RONOformat = keikei
    Dim VarMatLINE_NO                                   As String
    Dim VarMatNo                                        As String
    Dim VarDescription                                  As String
    Dim VarMatCode                                      As String
    Dim VarQty                                          As String
    Dim VarUnitPrice                                    As String
    Dim VarUnitCost                                     As String
    Dim VarMatAmount                                    As String
    Dim VarChargeTo                                     As String
    Dim VarMatDiscount                                  As String

    Dim REF_RIV_ADB                                     As String
    Dim MATREP_OR                                       As String
    Dim MATLEVEL                                        As String
    Dim MATLINE_NO                                      As String
    Dim MATDETCDE                                       As String
    Dim MATDETDSC                                       As String
    Dim MATDETUNT                                       As String
    Dim MATDETVOL                                       As Double
    Dim MATDETPRC                                       As Double
    Dim MATDETAMT                                       As Double
    Dim MATDETCOST                                      As Double
    Dim MatCode                                         As String
    Dim MATWCODE                                        As String
    Dim MATTAXRATE                                      As Double
    Dim MATDISCRATE                                     As Double
    Dim MATTAXVAL                                       As Double
    Dim MATDISVAL                                       As Double
    Dim MATPOCODE                                       As String
    Dim MATRep_Or2                                      As String
    Dim MATDETAIL                                       As String
    Dim MATDET_AMT                                      As Double
    Dim MATDIS_VAL                                      As Double
    Dim MATDISCOUNT_2                                   As Double
    Dim MATREMARKS                                      As String
    Dim Tdaycnt                                         As Integer
    Dim rsRR_HDCheck                                    As ADODB.Recordset
    Dim rsRR_HDTdaytranCheck                            As ADODB.Recordset
    Dim vGJorBP                                         As String
    
    Mcnt = 0
    vPIS_NO_CHARGE_TO = "NULL"

    'UPDATED BY: JUN----------------------------------------------------------------------------------------------------------------------------------------
    'DATE UPDATED: 11-28-2008
    'DESCRIPTION: TCN 12592 - TO CHECK IF THE MISC IS COMMING FROM THE PARTS OR IN THE F5
    '                         AND ONLY MISC FROM PARTS ONLY SHOULD BE DELETED
    If VALID_COMPANY_CODE_FORHAI = True Then
        gconDMIS.Execute "delete from CSMS_RO_Det where  DETCOST IS NOT NULL AND livil = '3' and rep_or = " & N2Str2Null(txtRep_Or.Text)
    Else
        'ORIGINAL CODE----------11-28-2008-----------------
        gconDMIS.Execute "delete from CSMS_RO_Det where DETCDE <> 'MISC' AND livil = '3' and rep_or = " & N2Str2Null(txtRep_Or.Text)
        'ORIGINAL CODE----------11-28-2008-----------------
    End If

    'UPDATE BY JUN 06/05/2008--------IMPORTING OF MATERIALS-------------
    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HAS" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HCO" Then
        Call importMaterialsFromSublet(txtRep_Or.Text)
    End If
    '-------------------------------------------------------------------
    Set rsMATISS = New ADODB.Recordset
    Set rsMATISS = gconDMIS.Execute("select rono,tranno,REFPISNO from PMIS_ORD_HD where [TYPE] = 'M' and TRANTYPE = 'RIV' and rono = '" & RONOformat & "' and (status <> 'C' AND STATUS <> 'N')")
    If Not rsMATISS.EOF And Not rsMATISS.BOF Then
        rsMATISS.MoveFirst
        Do While Not rsMATISS.EOF
            If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(rsMATISS!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
            If Mid(Null2String(rsMATISS!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

            Set rsTdayTran = New ADODB.Recordset
            Set rsTdayTran = gconDMIS.Execute("select trantype,tranno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,tranucost from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'RIV' and tranno = " & N2Str2Null(rsMATISS!TRANNO) & " and (status <> 'C' AND STATUS <> 'N')")
            If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                Screen.MousePointer = 11
                rsTdayTran.MoveFirst
                Do While Not rsTdayTran.EOF
                    Mcnt = Mcnt + 1
                    VarMatLINE_NO = "": VarMatNo = "": VarDescription = ""
                    VarMatCode = "": VarQty = 0: VarUnitPrice = 0
                    VarMatAmount = "": VarChargeTo = " ": VarMatDiscount = ZERO

                    VarMatLINE_NO = Format(Mcnt, "00")
                    VarMatNo = Null2String(rsTdayTran!STOCK_ORD)
                    VarDescription = SetMatDisc(Null2String(rsTdayTran!STOCK_SUP))
                    VarMatCode = "01"
                    VarQty = N2Str2IntZero(rsTdayTran!tranqty)
                    VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                    VarMatAmount = VarQty * VarUnitPrice
                    VarChargeTo = " "
                    VarMatDiscount = ZERO

                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'M' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsMATISS!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Do While Not rsRR_HDCheck.EOF
                            Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                            Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                            If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                                rsRR_HDTdaytranCheck.MoveFirst
                                Do While Not rsRR_HDTdaytranCheck.EOF
                                    If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarMatNo) Then
                                        MsgBox "Warning: Material Number : " & VarMatNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                        Mcnt = Mcnt - 1
                                        GoTo 40000
                                    End If
                                    rsRR_HDTdaytranCheck.MoveNext
                                Loop
                            End If
                            rsRR_HDCheck.MoveNext
                        Loop
                    End If


                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'M' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsMATISS!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Do While Not rsRR_HDCheck.EOF
                            Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                            Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                            If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                                rsRR_HDTdaytranCheck.MoveFirst
                                Do While Not rsRR_HDTdaytranCheck.EOF
                                    If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarMatNo) Then
                                        MsgBox "Warning: Material Number : " & VarMatNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                        Mcnt = Mcnt - 1
                                        GoTo 40000
                                    End If
                                    rsRR_HDTdaytranCheck.MoveNext
                                Loop
                            End If
                            rsRR_HDCheck.MoveNext
                        Loop
                    End If

                    MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
                    MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

                    MATREP_OR = N2Str2Null(txtRep_Or.Text)
                    MATLEVEL = "'3'"
                    MATLINE_NO = N2Str2Null(Format(VarMatLINE_NO, "00"))
                    MATDETCDE = N2Str2Null(VarMatNo)
                    MATDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                    MATDETUNT = "NULL"
                    MATDETVOL = N2Str2Zero(VarQty)
                    MATDETPRC = NumericVal(VarUnitPrice)
                    MATDETAMT = Round(NumericVal(VarMatAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    MATDETCOST = Round(NumericVal(N2Str2Zero(rsTdayTran!TRANUCOST)), 2)
                    MatCode = "NULL"
                    MATWCODE = vPIS_NO_CHARGE_TO
                    MATTAXRATE = (VAT_RATE / 100)
                    MATDISCRATE = NumericVal(VarMatDiscount) / 100
                    MATDISVAL = Round((MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE), 2)
                    MATPOCODE = N2Str2Null(VarMatCode)
                    MATRep_Or2 = "NULL"
                    MATDETAIL = "NULL"
                    MATDET_AMT = NumericVal(VarMatAmount)
                    MATDIS_VAL = Round(MATDISVAL * MATTAXRATE, 2)
                    MATDISCOUNT_2 = Round(MATDET_AMT * MATDISCRATE, 2)
                    MATTAXVAL = Round((MATDETAMT - MATDISCOUNT_2) * MATTAXRATE, 2)

                    gconDMIS.Execute "insert into CSMS_RO_Det " & _
                        "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detprc,detcost,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2)" & _
                        " values (" & MATREP_OR & ", " & MATLEVEL & ", " & MATLINE_NO & "," & _
                        " " & vGJorBP & "," & MATDETCDE & "," & MATDETDSC & "," & _
                        " " & MATDETUNT & ", " & MATDETVOL & "," & _
                        " " & MATDETPRC & "," & MATDETCOST & ", " & MATDETAMT & ", " & MatCode & _
                        ", " & MATWCODE & ", " & MATTAXRATE * 100 & ", " & MATDISCRATE * 100 & _
                        ", " & MATTAXVAL & ", " & MATDISVAL & ", " & MATPOCODE & _
                        ", " & MATRep_Or2 & ", " & MATDETAIL & ", " & MATDET_AMT & _
                        ", " & MATDIS_VAL & ", " & MATDISCOUNT_2 & ")"
                    Screen.MousePointer = 0
40000               rsTdayTran.MoveNext
                Loop
            End If
            rsMATISS.MoveNext
        Loop
    End If
    
    Set rsMATISS = New ADODB.Recordset
    Set rsMATISS = gconDMIS.Execute("select rono,tranno,REFPISNO from PMIS_ORD_Hist where [TYPE] = 'M' and TRANTYPE = 'RIV' and rono = '" & RONOformat & "'")
    If Not rsMATISS.EOF And Not rsMATISS.BOF Then
        rsMATISS.MoveFirst
        Do While Not rsMATISS.EOF
            If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(rsMATISS!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
            If Mid(Null2String(rsMATISS!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

            Set rsTdayTran = New ADODB.Recordset
            Set rsTdayTran = gconDMIS.Execute("select trantype,tranno,STOCK_ORD,STOCK_ORD,tranqty,tranuprice,tranucost from PMIS_DayTran where [TYPE] = 'M' AND trantype = 'RIV' and tranno = " & N2Str2Null(rsMATISS!TRANNO) & " and (status <> 'C' AND STATUS <> 'N')")
            If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                Screen.MousePointer = 11
                rsTdayTran.MoveFirst
                Do While Not rsTdayTran.EOF
                    Mcnt = Mcnt + 1
                    VarMatLINE_NO = "": VarMatNo = "": VarDescription = ""
                    VarMatCode = "": VarQty = 0: VarUnitPrice = 0
                    VarMatAmount = "": VarChargeTo = " ": VarMatDiscount = ZERO
                    VarMatLINE_NO = Format(Mcnt, "00")
                    VarMatNo = Null2String(rsTdayTran!STOCK_ORD)
                    VarDescription = SetMatDisc(Null2String(rsTdayTran!STOCK_ORD))
                    VarMatCode = "01"
                    VarQty = N2Str2IntZero(rsTdayTran!tranqty)
                    VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                    VarMatAmount = VarQty * VarUnitPrice
                    VarChargeTo = " "
                    VarMatDiscount = ZERO

                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'M' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsMATISS!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Do While Not rsRR_HDCheck.EOF
                            Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                            Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                            If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                                rsRR_HDTdaytranCheck.MoveFirst
                                Do While Not rsRR_HDTdaytranCheck.EOF
                                    If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarMatNo) Then
                                        MsgBox "Warning: Material Number : " & VarMatNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                        Mcnt = Mcnt - 1
                                        GoTo 50000
                                    End If
                                    rsRR_HDTdaytranCheck.MoveNext
                                Loop
                            End If
                            rsRR_HDCheck.MoveNext
                        Loop
                    End If

                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'M' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsMATISS!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Do While Not rsRR_HDCheck.EOF
                            Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                            Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                            If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                                rsRR_HDTdaytranCheck.MoveFirst
                                Do While Not rsRR_HDTdaytranCheck.EOF
                                    If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarMatNo) Then
                                        MsgBox "Warning: Material Number : " & VarMatNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                        Mcnt = Mcnt - 1
                                        GoTo 50000
                                    End If
                                    rsRR_HDTdaytranCheck.MoveNext
                                Loop
                            End If
                            rsRR_HDCheck.MoveNext
                        Loop
                    End If

                    MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
                    MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

                    MATREP_OR = N2Str2Null(txtRep_Or.Text)
                    MATLEVEL = "'3'"
                    MATLINE_NO = N2Str2Null(Format(VarMatLINE_NO, "00"))
                    MATDETCDE = N2Str2Null(VarMatNo)
                    MATDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                    MATDETUNT = "NULL"
                    MATDETVOL = N2Str2Zero(VarQty)
                    MATDETPRC = NumericVal(VarUnitPrice)
                    MATDETAMT = Round(NumericVal(VarMatAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    MATDETCOST = Round(NumericVal(N2Str2Zero(rsTdayTran!TRANUCOST)), 2)
                    MatCode = "NULL"
                    MATWCODE = vPIS_NO_CHARGE_TO
                    MATTAXRATE = (VAT_RATE / 100)
                    MATDISCRATE = NumericVal(VarMatDiscount) / 100
                    MATDISVAL = Round((MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE), 2)
                    MATPOCODE = N2Str2Null(VarMatCode)
                    MATRep_Or2 = "NULL"
                    MATDETAIL = "NULL"
                    MATDET_AMT = NumericVal(VarMatAmount)
                    MATDIS_VAL = Round(MATDISVAL * MATTAXRATE, 2)
                    MATDISCOUNT_2 = Round(MATDET_AMT * MATDISCRATE, 2)
                    MATTAXVAL = Round((MATDETAMT - MATDISCOUNT_2) * MATTAXRATE, 2)

                    gconDMIS.Execute "insert into CSMS_RO_Det " & _
                        "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detprc,detcost,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2)" & _
                        " values (" & MATREP_OR & "," & MATLEVEL & "," & MATLINE_NO & "," & _
                        " " & vGJorBP & "," & MATDETCDE & "," & MATDETDSC & "," & _
                        " " & MATDETUNT & ", " & MATDETVOL & "," & _
                        " " & MATDETPRC & "," & MATDETCOST & ", " & MATDETAMT & ", " & MatCode & _
                        ", " & MATWCODE & ", " & MATTAXRATE * 100 & ", " & MATDISCRATE * 100 & _
                        ", " & MATTAXVAL & ", " & MATDISVAL & ", " & MATPOCODE & _
                        ", " & MATRep_Or2 & ", " & MATDETAIL & ", " & MATDET_AMT & _
                        ", " & MATDIS_VAL & ", " & MATDISCOUNT_2 & ")"
                    Screen.MousePointer = 0
50000               rsTdayTran.MoveNext
                Loop
            End If
            rsMATISS.MoveNext
        Loop
    End If

    'MATERIALS ADVANCE BILL
    Set rsMATISS = New ADODB.Recordset
    Set rsMATISS = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hd where [TYPE] = 'M' AND status <> 'C' and trantype = 'ADB' and rono = '" & RONOformat & "'")
    If Not rsMATISS.EOF And Not rsMATISS.BOF Then
        If MsgQuestionBox("Advance Bill for Repair Order: " & txtRep_Or.Text & " is Available " & vbCrLf & _
                          "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then
            Do While Not rsMATISS.EOF
                If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
                If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
                If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

                If Mid(Null2String(rsMATISS!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
                If Mid(Null2String(rsMATISS!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

                Set rsTdayTran = New ADODB.Recordset
                Set rsTdayTran = gconDMIS.Execute("select itemno,trantype,tranno,stock_ord,stock_sup,tranqty,tranucost,tranuprice from PMIS_TdayTran where [TYPE] = 'M' AND trantype = 'ADB' and tranno = " & N2Str2Null(rsMATISS!TRANNO) & " order by itemno asc")
                If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                    Screen.MousePointer = 11
                    rsTdayTran.MoveFirst
                    Do While Not rsTdayTran.EOF
                        Mcnt = Mcnt + 1
                        VarMatLINE_NO = "": VarMatNo = "": VarDescription = ""
                        VarMatCode = "": VarQty = 0: VarUnitPrice = 0
                        VarMatAmount = "": VarChargeTo = " ": VarMatDiscount = ZERO

                        VarMatLINE_NO = Format(Pcnt, "00")
                        VarMatNo = Null2String(rsTdayTran!STOCK_ORD)
                        VarDescription = Null2String(rsTdayTran!STOCK_SUP)
                        VarMatCode = "01"
                        VarQty = Format(N2Str2IntZero(rsTdayTran!tranqty), "####0")
                        VarUnitCost = N2Str2Zero(rsTdayTran!TRANUCOST)
                        VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                        VarMatAmount = N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE)
                        VarChargeTo = " "
                        VarMatDiscount = ZERO

                        MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
                        MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

                        REF_RIV_ADB = "'ADB" & Format(Null2String(rsTdayTran!TRANNO), "000000") & Format(Null2String(rsTdayTran!itemno), "000") & "'"
                        MATREP_OR = N2Str2Null(txtRep_Or.Text)
                        MATLEVEL = "'3'"
                        MATLINE_NO = N2Str2Null(Format(MATLINE_NO, "00"))
                        MATDETCDE = N2Str2Null(VarMatNo)
                        MATDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                        MATDETUNT = "NULL"
                        MATDETVOL = N2Str2Zero(VarQty)
                        MATDETPRC = NumericVal(VarUnitPrice)
                        MATDETAMT = Round(NumericVal(VarMatAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                        MatCode = "NULL"
                        MATWCODE = vPIS_NO_CHARGE_TO
                        MATTAXRATE = (VAT_RATE / 100)
                        MATDISCRATE = NumericVal(VarMatDiscount) / 100
                        MATDISVAL = Round((MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE), 2)
                        MATPOCODE = N2Str2Null(VarMatCode)
                        MATRep_Or2 = "NULL"
                        MATDETAIL = "NULL"
                        MATDET_AMT = NumericVal(VarMatAmount)
                        MATDIS_VAL = Round(MATDISVAL * MATTAXRATE, 2)
                        MATDISCOUNT_2 = Round(MATDET_AMT * MATDISCRATE, 2)
                        MATTAXVAL = Round((MATDETAMT - MATDISCOUNT_2) * MATTAXRATE, 2)

                        gconDMIS.Execute "insert into CSMS_RO_Det " & _
                            "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                            " values (" & MATREP_OR & ", " & MATLEVEL & ", " & VarMatLINE_NO & "," & _
                            " " & vGJorBP & "," & MATDETCDE & "," & MATDETDSC & "," & _
                            " " & MATDETUNT & ", " & MATDETVOL & "," & _
                            " " & MATDETPRC & ", " & MATDETAMT & ", " & MatCode & _
                            ", " & MATWCODE & ", " & MATTAXRATE * 100 & ", " & MATDISCRATE * 100 & _
                            ", " & MATTAXVAL & ", " & MATDISVAL & ", " & MATPOCODE & _
                            ", " & MATRep_Or2 & ", " & MATDETAIL & ", " & MATDET_AMT & _
                            ", " & MATDIS_VAL & ", " & MATDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                        Screen.MousePointer = 0
                        rsTdayTran.MoveNext
                    Loop
                End If
                rsMATISS.MoveNext
            Loop
        End If
    End If

    Set rsMATISS = New ADODB.Recordset
    Set rsMATISS = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hist where [TYPE] = 'M' AND status <> 'C' and trantype = 'ADB' and rono = '" & RONOformat & "'")
    If Not rsMATISS.EOF And Not rsMATISS.BOF Then
        If MsgQuestionBox("Advance Bill for Repair Order: " & txtRep_Or.Text & " is Available " & vbCrLf & _
                          "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then
            Do While Not rsMATISS.EOF
                If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
                If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
                If Mid(Null2String(rsMATISS!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

                If Mid(Null2String(rsMATISS!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
                If Mid(Null2String(rsMATISS!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

                Set rsTdayTran = New ADODB.Recordset
                Set rsTdayTran = gconDMIS.Execute("select itemno,trantype,tranno,stock_ord,stock_sup,tranqty,tranuprice from PMIS_DayTran where [TYPE] = 'M' AND trantype = 'ADB' and tranno = " & N2Str2Null(rsMATISS!TRANNO) & " order by itemno asc")
                If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                    Screen.MousePointer = 11
                    rsTdayTran.MoveFirst
                    Do While Not rsTdayTran.EOF
                        Mcnt = Mcnt + 1
                        VarMatLINE_NO = "": VarMatNo = "": VarDescription = ""
                        VarMatCode = "": VarQty = 0: VarUnitPrice = 0
                        VarMatAmount = "": VarChargeTo = " ": VarMatDiscount = ZERO

                        VarMatLINE_NO = Format(Mcnt, "00")
                        VarMatNo = Null2String(rsTdayTran!STOCK_ORD)
                        VarDescription = Null2String(rsTdayTran!STOCK_SUP)
                        VarMatCode = "01"
                        VarQty = Format(N2Str2IntZero(rsTdayTran!tranqty), "####0")
                        VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                        VarMatAmount = N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE)
                        VarChargeTo = " "
                        VarMatDiscount = ZERO

                        MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
                        MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

                        REF_RIV_ADB = "'ADB" & Format(Null2String(rsTdayTran!TRANNO), "000000") & Format(Null2String(rsTdayTran!itemno), "000") & "'"
                        MATREP_OR = N2Str2Null(txtRep_Or.Text)
                        MATLEVEL = "'3'"
                        MATLINE_NO = N2Str2Null(Format(VarMatLINE_NO, "00"))
                        MATDETCDE = N2Str2Null(VarMatNo)
                        MATDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                        MATDETUNT = "NULL"
                        MATDETVOL = N2Str2Zero(VarQty)
                        MATDETPRC = NumericVal(VarUnitPrice)
                        MATDETAMT = Round(NumericVal(VarMatAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                        MatCode = "NULL"
                        MATWCODE = vPIS_NO_CHARGE_TO
                        MATTAXRATE = (VAT_RATE / 100)
                        MATDISCRATE = NumericVal(VarMatDiscount) / 100
                        MATDISVAL = Round((MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE), 2)
                        MATPOCODE = N2Str2Null(VarMatCode)
                        MATRep_Or2 = "NULL"
                        MATDETAIL = "NULL"
                        MATDET_AMT = NumericVal(VarMatAmount)
                        MATDIS_VAL = Round(MATDISVAL * MATTAXRATE, 2)
                        MATDISCOUNT_2 = Round(MATDET_AMT * MATDISCRATE, 2)
                        MATTAXVAL = Round((MATDETAMT - MATDISCOUNT_2) * MATTAXRATE, 2)

                        gconDMIS.Execute "insert into CSMS_RO_Det " & _
                            "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                            " values (" & MATREP_OR & ", " & MATLEVEL & ", " & VarMatLINE_NO & "," & _
                            " " & vGJorBP & "," & MATDETCDE & "," & MATDETDSC & "," & _
                            " " & MATDETUNT & ", " & MATDETVOL & "," & _
                            " " & MATDETPRC & ", " & MATDETAMT & ", " & MatCode & _
                            ", " & MATWCODE & ", " & MATTAXRATE * 100 & ", " & MATDISCRATE * 100 & _
                            ", " & MATTAXVAL & ", " & MATDISVAL & ", " & MATPOCODE & _
                            ", " & MATRep_Or2 & ", " & MATDETAIL & ", " & MATDET_AMT & _
                            ", " & MATDIS_VAL & ", " & MATDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                        Screen.MousePointer = 0
                        rsTdayTran.MoveNext
                    Loop
                End If
                rsMATISS.MoveNext
            Loop
        End If
    End If
    'ADVANCE BILL FOR MATERIAL

    If chkParticipat.Value = 0 Then
        FillMaterials
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        gconDMIS.Execute "update CSMS_RepOr set" & _
                       " material = " & TOTMATAMT - TOTMATTAX & "," & _
                       " m_amtvalue = " & TOTMATAMT & "," & _
                       " m_disc = " & TOTMATDISCVAL & "," & _
                       " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
                       " m_taxval = " & TOTMATTAX & "," & _
                       " m_discount = " & TOTMATDISC & "," & _
                       " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                       " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                       " wm_amt = " & 0 & "," & _
                       " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                       " where id = " & labid.Caption
        rsRefresh
        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
        cmdMatCancel.Value = True
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0: ShowVBError
End Sub

Sub ImportAccessories()
    Dim RONOformat                                     As String
    Dim yza                                            As Integer
    Dim tisoy, keikei                                  As String
    RONOformat = ""

    keikei = "": tisoy = "": yza = 0
    For yza = 1 To Len(rsREPOR!REP_OR)
        tisoy = Mid(rsREPOR!REP_OR, yza, 1)
        keikei = keikei + tisoy
    Next
    RONOformat = keikei
    
    Dim VarAccLINE_NO                                   As String
    Dim VarAccNo                                        As String
    Dim VarDescription                                  As String
    Dim VarAccCode                                      As String
    Dim VarQty                                          As String
    Dim VarUnitPrice                                    As String
    Dim VarAccAmount                                    As String
    Dim VarChargeTo                                     As String
    Dim VarAccDiscount                                  As String

    Dim ACCREP_OR                                       As String
    Dim ACCLEVEL                                        As String
    Dim ACCLINE_NO                                      As String
    Dim ACCDETCDE                                       As String
    Dim ACCDETDSC                                       As String
    Dim ACCDETUNT                                       As String
    Dim ACCDETVOL                                       As Double
    Dim ACCDETPRC                                       As Double
    Dim ACCDETAMT                                       As Double
    Dim ACCDETCOST                                      As Double
    Dim MatCode                                         As String
    Dim ACCWCODE                                        As String
    Dim ACCTAXRATE                                      As Double
    Dim ACCDISCRATE                                     As Double
    Dim ACCTAXVAL                                       As Double
    Dim ACCDISVAL                                       As Double
    Dim ACCPOCODE                                       As String
    Dim ACCRep_Or2                                      As String
    Dim ACCDETAIL                                       As String
    Dim ACCDET_AMT                                      As Double
    Dim ACCDIS_VAL                                      As Double
    Dim ACCDISCOUNT_2                                   As Double
    Dim ACCREMARKS                                      As String
    Dim Tdaycnt                                         As Integer
    Dim rsRR_HDCheck                                    As ADODB.Recordset
    Dim rsRR_HDTdaytranCheck                            As ADODB.Recordset

    Acnt = 0
    vPIS_NO_CHARGE_TO = "NULL"
    gconDMIS.Execute "delete from CSMS_RO_Det where livil = '4' and rep_or = " & N2Str2Null(txtRep_Or.Text)
    Dim vGJorBP                                        As String
    vGJorBP = "NULL"
    Set rsACCISS = New ADODB.Recordset
    Set rsACCISS = gconDMIS.Execute("select rono,tranno,REFPISNO from PMIS_ORD_HD where [TYPE] = 'A' and TRANTYPE = 'RIV' and rono = '" & RONOformat & "' and status <> 'C' AND STATUS <> 'N'")
    If Not rsACCISS.EOF And Not rsACCISS.BOF Then
        rsACCISS.MoveFirst
        Do While Not rsACCISS.EOF
            If Mid(Null2String(rsACCISS!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(rsACCISS!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(rsACCISS!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(rsACCISS!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
            If Mid(Null2String(rsACCISS!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

            Set rsTdayTran = New ADODB.Recordset
            Set rsTdayTran = gconDMIS.Execute("select trantype,tranno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,tranucost from PMIS_Tdaytran where [TYPE] = 'A' AND trantype = 'RIV' and tranno = " & N2Str2Null(rsACCISS!TRANNO) & " and status <> 'C' AND STATUS <> 'N'")
            If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                Screen.MousePointer = 11
                rsTdayTran.MoveFirst
                Do While Not rsTdayTran.EOF
                    Mcnt = Mcnt + 1
                    VarAccLINE_NO = "": VarAccNo = "": VarDescription = ""
                    VarAccCode = "": VarQty = 0: VarUnitPrice = 0
                    VarAccAmount = "": VarChargeTo = " ": VarAccDiscount = ZERO

                    VarAccLINE_NO = Format(Mcnt, "00")
                    VarAccNo = Null2String(rsTdayTran!STOCK_ORD)
                    VarDescription = SetAccDisc(Null2String(rsTdayTran!STOCK_SUP))
                    VarAccCode = "01"
                    VarQty = N2Str2IntZero(rsTdayTran!tranqty)
                    VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                    VarAccAmount = VarQty * VarUnitPrice
                    VarChargeTo = " "
                    VarAccDiscount = ZERO

                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'A' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsACCISS!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                        Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'A' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                        If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                            rsRR_HDTdaytranCheck.MoveFirst
                            Do While Not rsRR_HDTdaytranCheck.EOF
                                If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarAccNo) Then
                                    MsgBox "Warning: Accessories Number : " & VarAccNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                    Mcnt = Mcnt - 1
                                    GoTo 80000
                                End If
                                rsRR_HDTdaytranCheck.MoveNext
                            Loop
                        End If
                    End If
                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'A' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsACCISS!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                        Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'A' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                        If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                            rsRR_HDTdaytranCheck.MoveFirst
                            Do While Not rsRR_HDTdaytranCheck.EOF
                                If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarAccNo) Then
                                    MsgBox "Warning: Accessories Number : " & VarAccNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                    Mcnt = Mcnt - 1
                                    GoTo 80000
                                End If
                                rsRR_HDTdaytranCheck.MoveNext
                            Loop
                        End If
                    End If

                    ACCDISVAL = 0: ACCTAXVAL = 0: ACCDETAMT = 0
                    ACCDIS_VAL = 0: ACCDISCOUNT_2 = 0: ACCDISCRATE = 0

                    ACCREP_OR = N2Str2Null(txtRep_Or.Text)
                    ACCLEVEL = "'4'"
                    ACCLINE_NO = N2Str2Null(Format(VarAccLINE_NO, "00"))
                    ACCDETCDE = N2Str2Null(VarAccNo)
                    ACCDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                    ACCDETUNT = "NULL"
                    ACCDETVOL = N2Str2Zero(VarQty)
                    ACCDETPRC = NumericVal(VarUnitPrice)
                    ACCDETAMT = Round(NumericVal(VarAccAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    ACCDETCOST = Round(NumericVal(N2Str2Zero(rsTdayTran!TRANUCOST)), 2)
                    MatCode = "NULL"
                    ACCWCODE = vPIS_NO_CHARGE_TO
                    ACCTAXRATE = (VAT_RATE / 100)
                    ACCDISCRATE = NumericVal(VarAccDiscount) / 100
                    ACCDISVAL = Round((ACCDETPRC * ACCDISCRATE) - ((ACCDETPRC * ACCDISCRATE) * ACCTAXRATE), 2)
                    ACCPOCODE = N2Str2Null(VarAccCode)
                    ACCRep_Or2 = "NULL"
                    ACCDETAIL = "NULL"
                    ACCDET_AMT = NumericVal(VarAccAmount)
                    ACCDIS_VAL = Round(ACCDISVAL * ACCTAXRATE, 2)
                    ACCDISCOUNT_2 = Round(ACCDET_AMT * ACCDISCRATE, 2)
                    ACCTAXVAL = Round((ACCDETAMT - ACCDISCOUNT_2) * ACCTAXRATE, 2)

                    gconDMIS.Execute "insert into CSMS_RO_Det " & _
                        "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detprc,detcost,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2)" & _
                        " values (" & ACCREP_OR & ", " & ACCLEVEL & ", " & ACCLINE_NO & "," & _
                        " " & vGJorBP & "," & ACCDETCDE & "," & ACCDETDSC & "," & _
                        " " & ACCDETUNT & ", " & ACCDETVOL & "," & _
                        " " & ACCDETPRC & "," & ACCDETCOST & ", " & ACCDETAMT & ", " & MatCode & _
                        ", " & ACCWCODE & ", " & ACCTAXRATE * 100 & ", " & ACCDISCRATE * 100 & _
                        ", " & ACCTAXVAL & ", " & ACCDISVAL & ", " & ACCPOCODE & _
                        ", " & ACCRep_Or2 & ", " & ACCDETAIL & ", " & ACCDET_AMT & _
                        ", " & ACCDIS_VAL & ", " & ACCDISCOUNT_2 & ")"
                    Screen.MousePointer = 0
80000               rsTdayTran.MoveNext
                Loop
            End If
            rsACCISS.MoveNext
        Loop
    End If
    
    Set rsACCISS = New ADODB.Recordset
    Set rsACCISS = gconDMIS.Execute("select rono,tranno,REFPISNO from PMIS_ORD_Hist where [TYPE] = 'A' and TRANTYPE = 'RIV' and rono = '" & RONOformat & "'")
    If Not rsACCISS.EOF And Not rsACCISS.BOF Then
        rsACCISS.MoveFirst
        Do While Not rsACCISS.EOF
            If Mid(Null2String(rsACCISS!refpisno), 5, 1) = "C" Then vPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(rsACCISS!refpisno), 5, 1) = "I" Then vPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(rsACCISS!refpisno), 5, 1) = "W" Then vPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(rsACCISS!refpisno), 4, 1) = "B" Then vGJorBP = "'BP'"
            If Mid(Null2String(rsACCISS!refpisno), 4, 1) = "G" Then vGJorBP = "'GJ'"

            Set rsTdayTran = New ADODB.Recordset
            Set rsTdayTran = gconDMIS.Execute("select trantype,tranno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,tranucost from PMIS_DayTran where [TYPE] = 'A' AND trantype = 'RIV' and tranno = " & N2Str2Null(rsACCISS!TRANNO) & " and status <> 'C' AND STATUS <> 'N'")
            If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                Screen.MousePointer = 11
                rsTdayTran.MoveFirst
                Do While Not rsTdayTran.EOF
                    Mcnt = Mcnt + 1
                    VarAccLINE_NO = "": VarAccNo = "": VarDescription = ""
                    VarAccCode = "": VarQty = 0: VarUnitPrice = 0
                    VarAccAmount = "": VarChargeTo = " ": VarAccDiscount = ZERO

                    VarAccLINE_NO = Format(Mcnt, "00")
                    VarAccNo = Null2String(rsTdayTran!STOCK_ORD)
                    VarDescription = Null2String(rsTdayTran!STOCK_SUP)
                    VarAccCode = "01"
                    VarQty = N2Str2IntZero(rsTdayTran!tranqty)
                    VarUnitPrice = N2Str2Zero(rsTdayTran!TRANUPRICE)
                    VarAccAmount = VarQty * VarUnitPrice
                    VarChargeTo = " "
                    VarAccDiscount = ZERO

                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'A' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsACCISS!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                        Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'A' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                        If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                            rsRR_HDTdaytranCheck.MoveFirst
                            Do While Not rsRR_HDTdaytranCheck.EOF
                                If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarAccNo) Then
                                    MsgBox "Warning: Accessories Number : " & VarAccNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                    Mcnt = Mcnt - 1
                                    GoTo 90000
                                End If
                                rsRR_HDTdaytranCheck.MoveNext
                            Loop
                        End If
                    End If
                    Set rsRR_HDCheck = New ADODB.Recordset
                    Set rsRR_HDCheck = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'A' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(rsACCISS!TRANNO), "000000") & "'")
                    If Not rsRR_HDCheck.EOF And Not rsRR_HDCheck.BOF Then
                        Set rsRR_HDTdaytranCheck = New ADODB.Recordset
                        Set rsRR_HDTdaytranCheck = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'A' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HDCheck!rrno) & " order by Itemno asc")
                        If Not rsRR_HDTdaytranCheck.EOF And Not rsRR_HDTdaytranCheck.BOF Then
                            rsRR_HDTdaytranCheck.MoveFirst
                            Do While Not rsRR_HDTdaytranCheck.EOF
                                If UCase(Null2String(rsRR_HDTdaytranCheck!STOCK_ORD)) = UCase(VarAccNo) Then
                                    MsgBox "Warning: Accessories Number : " & VarAccNo & " was returned in warehouse." & vbCrLf & "         this will not be imported in billing"
                                    Mcnt = Mcnt - 1
                                    GoTo 90000
                                End If
                                rsRR_HDTdaytranCheck.MoveNext
                            Loop
                        End If
                    End If

                    ACCDISVAL = 0: ACCTAXVAL = 0: ACCDETAMT = 0
                    ACCDIS_VAL = 0: ACCDISCOUNT_2 = 0: ACCDISCRATE = 0

                    ACCREP_OR = N2Str2Null(txtRep_Or.Text)
                    ACCLEVEL = "'4'"
                    ACCLINE_NO = N2Str2Null(Format(VarAccLINE_NO, "00"))
                    ACCDETCDE = N2Str2Null(VarAccNo)
                    ACCDETDSC = N2Str2Null(Mid(VarDescription, 1, 100))
                    ACCDETUNT = "NULL"
                    ACCDETVOL = N2Str2Zero(VarQty)
                    ACCDETPRC = NumericVal(VarUnitPrice)
                    ACCDETAMT = Round(NumericVal(VarAccAmount) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    ACCDETCOST = Round(NumericVal(N2Str2Zero(rsTdayTran!TRANUCOST)), 2)
                    MatCode = "NULL"
                    ACCWCODE = vPIS_NO_CHARGE_TO
                    ACCTAXRATE = (VAT_RATE / 100)
                    ACCDISCRATE = NumericVal(VarAccDiscount) / 100
                    ACCDISVAL = Round((ACCDETPRC * ACCDISCRATE) - ((ACCDETPRC * ACCDISCRATE) * ACCTAXRATE), 2)
                    ACCPOCODE = N2Str2Null(VarAccCode)
                    ACCRep_Or2 = "NULL"
                    ACCDETAIL = "NULL"
                    ACCDET_AMT = NumericVal(VarAccAmount)
                    ACCDIS_VAL = Round(ACCDISVAL * ACCTAXRATE, 2)
                    ACCDISCOUNT_2 = Round(ACCDET_AMT * ACCDISCRATE, 2)
                    ACCTAXVAL = Round((ACCDETAMT - ACCDISCOUNT_2) * ACCTAXRATE, 2)

                    gconDMIS.Execute "insert into CSMS_RO_Det " & _
                        "(rep_or,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detcost,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2)" & _
                        " values (" & ACCREP_OR & ", " & ACCLEVEL & ", " & ACCLINE_NO & "," & _
                        " " & ACCDETCDE & "," & ACCDETDSC & "," & _
                        " " & ACCDETUNT & ", " & ACCDETVOL & "," & _
                        " " & ACCDETPRC & "," & ACCDETCOST & ", " & ACCDETAMT & ", " & MatCode & _
                        ", " & ACCWCODE & ", " & ACCTAXRATE * 100 & ", " & ACCDISCRATE * 100 & _
                        ", " & ACCTAXVAL & ", " & ACCDISVAL & ", " & ACCPOCODE & _
                        ", " & ACCRep_Or2 & ", " & ACCDETAIL & ", " & ACCDET_AMT & _
                        ", " & ACCDIS_VAL & ", " & ACCDISCOUNT_2 & ")"
                    Screen.MousePointer = 0
90000               rsTdayTran.MoveNext
                Loop
            End If
            rsACCISS.MoveNext
        Loop
    End If
    
    If chkParticipat.Value = 0 Then
        Call FillAccessories
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        
        gconDMIS.Execute "update CSMS_RepOr set" & _
            " Accessories = " & TOTACCAMT - TOTACCTAX & "," & _
            " A_amtvalue = " & TOTACCAMT & "," & _
            " A_disc = " & TOTACCDISCVAL & "," & _
            " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
            " A_taxval = " & TOTACCTAX & "," & _
            " A_discount = " & TOTACCDISC & "," & _
            " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
            " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
            " WA_amt = " & 0 & "," & _
            " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
            " where id = " & labid.Caption
        
        Call rsRefresh
        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
        cmdAccCancel.Value = True
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    Call ShowVBError
End Sub

Sub AddJobs()
    SSTab1.SelectedItem = 1
    Call SendToBack
    cmdAddJobs.ZOrder 0: fraAddJobs.ZOrder 0
    fraAddJobs.Enabled = True: AddorEdit = "ADD": Call InitJobs
End Sub

Sub AddParts()
    SSTab1.SelectedItem = 2
    Call SendToBack
    cmdAddParts.ZOrder 0: fraAddParts.ZOrder 0
    fraAddParts.Enabled = True: AddorEdit = "ADD"
    Call InitParts
    On Error Resume Next
    cboPartNo.SetFocus
End Sub

Sub AddMaterials()
    SSTab1.SelectedItem = 3
    Call SendToBack
    cmdAddMaterials.ZOrder 0: fraAddMaterials.ZOrder 0
    fraAddMaterials.Enabled = True: AddorEdit = "ADD"
    'Updated By     : IEBV 05262010 04:00 PM
    'Description    : To make the AddMAterials visible
    cmdAddMaterials.Visible = True
    'Description    : To make the AddMAterials visible
    'Updated By     :IEBV 05262010 04:00 PM
    Call InitMaterials
    
End Sub

Sub AddAccessories()
    SSTab1.SelectedItem = 4
    Call SendToBack
    cmdAddAccessories.ZOrder 0: fraAddAccessories.ZOrder 0
    fraAddAccessories.Enabled = True: AddorEdit = "ADD"
    Call InitAccessories
End Sub

Sub InitJobs()
    txtJobLineNo.Text = Format(kcnt + 1, "00")
    txtJobPostCode.Text = ""
    cboJobChargeTo.Clear
    cboJobChargeTo.AddItem "W"
    cboJobChargeTo.AddItem "S"
    cboJobChargeTo.AddItem "C"
    cboJcode.ListIndex = -1
    cboJobCode.ListIndex = -1
    cboJobChargeTo.ListIndex = -1
    cboAcctCodeLabor.Enabled = False
    cboAcctCodeLabor.ListIndex = -1
    cboTechnician.ListIndex = -1
    txtDET_HRS.Text = 0
    txtJobRate.Text = ZERO
    txtJobDiscount.Text = ZERO
    txtJobDetail.Text = ""
    PREV_LABOR_CHARGE_TO = ""
End Sub

Sub InitROJOBS()
    Set rsROJOBS = Nothing
    Set rsROJOBS = New ADODB.Recordset
    Set rsROJOBS = gconDMIS.Execute("Select Tech_Name from CSMS_vw_Technician order by Tech_Name asc")
    If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
        rsROJOBS.MoveFirst
        cboTechnician.Clear
        Do While Not rsROJOBS.EOF
            cboTechnician.AddItem Null2String(rsROJOBS!TECH_NAME)
            rsROJOBS.MoveNext
        Loop
    End If
    Set rsROJOBS = Nothing
End Sub

Sub InitCbo()
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
        rsEmpNo.MoveFirst
        cboRecd_by.Clear
        Do While Not rsEmpNo.EOF
            cboRecd_by.AddItem Null2String(rsEmpNo!NAYM)
            rsEmpNo.MoveNext
        Loop
    End If
    Set rsEmpNo = Nothing
    
    Dim rsChartAccount                                 As New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from CMIS_SBOOK WHERE BOOK = 'S' Order by CODE asc")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        rsChartAccount.MoveFirst
        Do While Not rsChartAccount.EOF
            cboAcctCodeLabor.AddItem Null2String(rsChartAccount!DESCNAME)
            cboAcctCodeParts.AddItem Null2String(rsChartAccount!DESCNAME)
            cboAcctCodeAccessories.AddItem Null2String(rsChartAccount!DESCNAME)
            cboAcctCodeMaterials.AddItem Null2String(rsChartAccount!DESCNAME)
            rsChartAccount.MoveNext
        Loop
    End If
    cboAcctCodeLabor.ListIndex = -1
    cboAcctCodeParts.ListIndex = -1
    cboAcctCodeAccessories.ListIndex = -1
    cboAcctCodeMaterials.ListIndex = -1
    Set rsChartAccount = Nothing
    InitROJOBS
    'Updated: IEBV 05100225PM
    'Description: Add ROType
    Call ADDROTYPE
    'Updated: IEBV 05100225PM
    'Description: Add ROType
End Sub

Sub StoreMemVars()
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        cboRecd_by.Enabled = False
        REPRINT = Null2String(rsREPOR!prin_dte)
        Screen.MousePointer = 11
        labid.Caption = rsREPOR!ID
        
        'UPDATE BY   : MJP09232009
        'DESCRIPTION : THIS IS TO HAVE A VALIDITY FOR THE CUSTOMER NAME IN PART MODULE AND SERVICE MODULE
            'txtRep_Or.Locked = True
        'UPDATE BY   : MJP09232009
        
        txtRep_Or.Text = Null2String(rsREPOR!REP_OR)
        txtInvoiceNo.Text = Null2String(rsREPOR!invoice)
        txtDte_Rel.Text = Null2String(rsREPOR!DTE_rel)
        'If txtInvoiceNo.Text = "" Then pic3.ZOrder 0 Else pic4.ZOrder 0
        
        txtEstimateno.Text = Null2String(rsREPOR!EstimateNo)
        txtROType.Text = Null2String(rsREPOR!ROTYPE)
        txtSvc_No.Text = Null2String(rsREPOR!svc_no)
        txtAcct_No.Text = Null2String(rsREPOR!ACCT_NO)
        If IsNull(rsREPOR!ACCT_NO) = False Then Call SetAdres(rsREPOR!ACCT_NO)
        txtNiym.Text = Null2String(rsREPOR!NIYM)
        txtPlate_No.Text = Null2String(rsREPOR!PLATE_NO)
        cboModel.Text = Null2String(rsREPOR!Model)
        txtMake.Text = SetMake(Null2String(rsREPOR!PLATE_NO))
        txtTerm.Text = Null2String(rsREPOR!TERM)
        txtSektion.Text = Null2String(rsREPOR!sektion)
        Call getcreditinfo(Null2String(rsREPOR!ACCT_NO))
        Call ShorROtype(Null2String(rsREPOR!REP_OR))
        
        If SetSA(Null2String(rsREPOR!RECD_BY)) = "" Then
            If Null2String(rsREPOR!RECD_BY) = "" Then
                cboRecd_by.ListIndex = -1
            Else
                cboRecd_by.AddItem "Missing: " & Null2String(rsREPOR!RECD_BY)
                cboRecd_by.Text = "Missing: " & Null2String(rsREPOR!RECD_BY)
            End If
        Else
            cboRecd_by.Text = SetSA(Null2String(rsREPOR!RECD_BY))
        End If
        
        txtKm_rdg.Text = Null2String(rsREPOR!km_rdg)
        txtDte_recd.Value = Null2Date(rsREPOR!DTE_RECD)
        txtCertific8.Text = Null2String(rsREPOR!certific8)
        txtDte_comp.Text = Null2String(rsREPOR!dte_comp)
        txtVIN.Text = Null2String(rsREPOR!Vin)
        txtParticipat.Text = Null2String(rsREPOR!participat)
        
        If CheckHasCreditLimit(txtAcct_No.Text) = True Then
            Label5.Enabled = True
        Else
            Label5.Enabled = False
        End If
        
        If txtParticipat.Text <> "" Then
            chkParticipat.Value = 1
            txtParticipation.Text = SetParticipatname(txtParticipat.Text)
            fraParticipation.Enabled = False
        Else
            chkParticipat.Value = 0
            txtParticipation.Text = ""
            fraParticipation.Enabled = False
        End If
        
        If txtInvoiceNo.Text <> "" Then
            cmdEdit.Enabled = False
            labF1.Enabled = True
            labF9.Enabled = True
            labF10.Enabled = True
        Else
            cmdEdit.Enabled = True
            labF1.Enabled = False
            labF9.Enabled = False
            labF10.Enabled = False
        End If
        If chkParticipat.Value = 1 Then
            Label22.Enabled = True
        Else
            Label22.Enabled = False
        End If
        
        DoEvents
        Call FillJobs
        Call FillParts
        Call FillMaterials
        Call FillAccessories
        Call FillDetails
        DoEvents
        
        If txtDte_Rel.Text = "" Then
            If txtDte_comp.Text <> "" Then
                cmdReleaseRO.Enabled = True
            Else
                cmdReleaseRO.Enabled = False
            End If
        Else
            cmdReleaseRO.Enabled = False
        End If
        
        DoEvents
        
        If Trim(Null2String(rsREPOR!CALLED_RESULT)) <> "" Then
            labCalled_Result.Caption = Null2String(rsREPOR!CALLED_RESULT)
            capFollow.VisualTheme = 3
        Else
            labCalled_Result.Caption = ""
            capFollow.VisualTheme = 2
        End If
        
        If Trim(Null2String(rsREPOR!RECOMMENDATION)) <> "" Then
            labNotes.Caption = Null2String(rsREPOR!RECOMMENDATION)
            capSUG.VisualTheme = 3
        Else
            labNotes.Caption = ""
            capSUG.VisualTheme = 2
        End If
        
        If Null2String(rsREPOR!invoice) = "" Then
            cmdNoCharge.Enabled = True
            Command1.Enabled = True
            Command2.Enabled = True
            cmdInternalRO.Enabled = True
            If N2Str2Zero(rsREPOR!VAT_EXEMPT) = 1 Then
                cmdROVatExempt.Caption = "Set to Not Zero Rated"
                cmdROVatExempt.Enabled = True
            Else
                cmdROVatExempt.Caption = "Set to Zero Rated"
                cmdROVatExempt.Enabled = True
            End If
            cmdDelete.Enabled = True
        Else
            cmdNoCharge.Enabled = False
            Command1.Enabled = False
            Command2.Enabled = False
            cmdInternalRO.Enabled = False
            If N2Str2Zero(rsREPOR!VAT_EXEMPT) = 1 Then
                cmdROVatExempt.Caption = "Set to Not Zero Rated"
            Else
                cmdROVatExempt.Caption = "Set to Zero Rated"
            End If
            cmdROVatExempt.Enabled = False
            cmdDelete.Enabled = False
        End If
        
        'UPDATE BY   : MJP 01142010 1113AM
        'DESCRIPTION : SO PROGRAM WILL NOT CHECK THE OR NUM AND SJ NUM OF THE RO IF ITS NOT INVOICE (LESS TIME QUERYING)
        If Not Null2String(rsREPOR!invoice) = "" Then
            If CheckORNum(Null2String(rsREPOR!invoice)) <> "" Or CheckSJNum(Null2String(rsREPOR!invoice)) <> "" Then
                Label5.ForeColor = &H808080
                labF9.ForeColor = &H808080
                labF10.ForeColor = &H808080
            Else
                Label5.ForeColor = &HFFFFFF
                labF9.ForeColor = &HFFFFFF
                labF10.ForeColor = &HFFFFFF
            End If
        Else
            Label5.ForeColor = &HFFFFFF
            labF9.ForeColor = &HFFFFFF
            labF10.ForeColor = &HFFFFFF
        End If
        
        'UPDATE BY   : MJP 01142010 1113AM
        'DESCRIPTION : SO PROGRAM WILL NOT CHECK THE OR NUM AND SJ NUM OF THE RO IF ITS NOT INVOICE (LESS TIME QUERYING)
        If Not Null2String(rsREPOR!invoice) = "" Then
            If CheckORNum(Null2String(rsREPOR!invoice)) <> "" Then
                labORNum.Caption = CheckORNum(Null2String(rsREPOR!invoice))
            Else
                labORNum.Caption = ""
            End If
        Else
            labORNum.Caption = ""
        End If
        
        If Null2String(rsREPOR!invoice) = "INT RO" Then
            If CheckSJINTRONum(Null2String(rsREPOR!REP_OR)) <> "" Then
                labSJNum.Caption = CheckSJINTRONum(Null2String(rsREPOR!REP_OR))
            Else
                labSJNum.Caption = ""
            End If
        Else
            'UPDATE BY   : MJP 01142010 1113AM
            'DESCRIPTION : SO PROGRAM WILL NOT CHECK THE SJ NUM OF THE RO IF ITS NOT INVOICE (LESS TIME QUERYING)
            If Not Null2String(rsREPOR!invoice) = "" Then
                If CheckSJNum(Null2String(rsREPOR!invoice)) <> "" Then
                    labSJNum.Caption = CheckSJNum(Null2String(rsREPOR!invoice))
                Else
                    labSJNum.Caption = ""
                End If
            Else
                labSJNum.Caption = ""
            End If
        End If
    
        txtLOAAmount.Text = Format(NumericVal(rsREPOR!INSAMT), MAXIMUM_DIGIT)
        txtPartLabor.Text = Format(NumericVal(rsREPOR!PARTLABOR), MAXIMUM_DIGIT)
        txtPartParts.Text = Format(NumericVal(rsREPOR!PARTPARTS), MAXIMUM_DIGIT)
        txtPartMaterials.Text = Format(NumericVal(rsREPOR!PARTMATERIALS), MAXIMUM_DIGIT)
        txtPartAccessories.Text = Format(NumericVal(rsREPOR!PARTACCESSORIES), MAXIMUM_DIGIT)
        txtPartTotal.Text = Format(NumericVal(rsREPOR!INSAMT), MAXIMUM_DIGIT)
        
        'UPDATE BY   : MJP 011409 0253PM
        'DESCRIPTION : TO SHOW THE USER THE ACTIVITIES IN THE PARTS MODULE
        lblMSG.Caption = ""
        lblMSG2.Caption = ""
        Dim XPART_MSG As Integer: Dim XPART_MSG2       As Integer
        Dim XMATE_MSG As Integer: Dim XMATE_MSG2       As Integer
        Dim XACCE_MSG As Integer: Dim XACCE_MSG2       As Integer
        Dim XXX_MSG                                    As String

        If txtInvoiceNo = "" Then
'            XPART_MSG = CheckIfRoDetIsIssued(2, "P")
'            XPART_MSG2 = CheckIfIssuedInRoDet(2, "P")
'
'            XMATE_MSG = CheckIfRoDetIsIssued(3, "M")
'            XMATE_MSG2 = CheckIfIssuedInRoDet(3, "M")
'
'            XACCE_MSG = CheckIfRoDetIsIssued(4, "A")
'            XACCE_MSG2 = CheckIfIssuedInRoDet(4, "A")
'
'            If XPART_MSG = 1 Or XMATE_MSG = 1 Or XACCE_MSG = 1 Then
'                If XPART_MSG = 1 Then XXX_MSG = "PARTS"
'                If XMATE_MSG = 1 Then XXX_MSG = XXX_MSG & ",MATERIALS"
'                If XACCE_MSG = 1 Then XXX_MSG = XXX_MSG & ",ACCS."
'                lblMSG.Caption = "SOME " & XXX_MSG & " ISSUANCE HAVE BEEN UNPOST IN PARTS MODULE, PRESS F8 TO IMPORT THE LATEST ISSUANCE FROM PARTS MODULE"
'            End If
'            XXX_MSG = ""
'            If XPART_MSG2 = 1 Or XMATE_MSG2 = 1 Or XACCE_MSG2 = 1 Then
'                If XPART_MSG2 = 1 Then XXX_MSG = "PARTS"
'                If XMATE_MSG2 = 1 Then XXX_MSG = XXX_MSG & ",MATERIALS"
'                If XACCE_MSG2 = 1 Then XXX_MSG = XXX_MSG & ",ACCS."
'                lblMSG2.Caption = "SOME " & XXX_MSG & " ISSUANCE ARE NOY YET IMPORTED TO THIS REPAIR ORDER. PRESS F8 TO IMPORT THE LATEST ISSUANCE FROM PARTS MODULE"
'            End If
        End If
        'UPDATE BY   : MJP 011409 0253PM

        Screen.MousePointer = 0
    Else
        'cmdFirst.Enabled = False: 'cmdLast.Enabled = False: 'cmdPrevious.Enabled = False:  'cmdNext.Enabled = False:
        cmdEdit.Enabled = False: cmdPrint.Enabled = False
    End If
End Sub
Function getcreditinfo(CustomerCode As String)
Dim rscredit As ADODB.Recordset
Set rscredit = New ADODB.Recordset

Set rscredit = gconDMIS.Execute("Select creditdays,creditlimit from all_customer where cuscde ='" & CustomerCode & "' ")
If Not (rscredit.EOF And rscredit.BOF) Then
    labCreditLimit.Caption = Null2String(rscredit!CreditLimit)
    labCreditDays.Caption = Null2String(rscredit!CREDITDAYS) + " Days"
Else
    labCreditLimit.Caption = "0.00"
    labCreditDays.Caption = "0"
End If

End Function

Function CheckIfRoDetIsIssued(xLIVIL As Integer, XTYPE As String) As Integer
    Dim rsCSMS                                         As New ADODB.Recordset
    Dim rsPMIS                                         As New ADODB.Recordset
    Dim rsDet                                          As New ADODB.Recordset
    Dim EXIST_HIST                                     As Boolean
    Dim EXIST_PRES                                     As Boolean
    Dim EXIST_ADVN                                     As Boolean
    Dim xMSG                                           As String
    Dim xMSG2                                          As String
    Dim XXX                                            As Integer

    'CHECK IF EXIST IN THE PARTS DEPARTMENT
    Set rsCSMS = gconDMIS.Execute("SELECT DETCDE FROM CSMS_RO_DET WHERE LIVIL = " & xLIVIL & " AND REP_OR = " & N2Str2Null(txtRep_Or) & " AND DETCDE <> 'MISC' AND ROTYPE IS NULL")
    If Not (rsCSMS.BOF And rsCSMS.EOF) Then
        Do While Not rsCSMS.EOF
            'CHECK HISTORY TRANSACTION
            Set rsPMIS = New ADODB.Recordset
            Set rsPMIS = gconDMIS.Execute("SELECT TRANNO FROM PMIS_ORD_HIST WHERE (TRANTYPE = 'RIV' OR TRANTYPE = 'ADB') AND (STATUS = 'P' OR STATUS = 'B') AND RONO = " & N2Str2Null(txtRep_Or) & " AND TYPE = " & N2Str2Null(XTYPE) & "")
            If Not (rsPMIS.BOF And rsPMIS.EOF) Then
                Do While Not rsPMIS.EOF
                    Set rsDet = gconDMIS.Execute("SELECT STOCK_ORD FROM PMIS_DAYTRAN WHERE (TRANTYPE = 'RIV' OR TRANTYPE = 'ADB') AND TYPE = " & N2Str2Null(XTYPE) & " AND TRANNO = '" & Null2String(rsPMIS!TRANNO) & "' AND (STATUS = 'P' OR STATUS = 'B') AND STOCK_ORD = '" & Null2String(rsCSMS!DETCDE) & "'")
                    If Not (rsDet.BOF And rsDet.EOF) Then
                        XXX = 0
                        GoTo NEXT_ITEM
                    Else
                        XXX = 1
                    End If
                    rsPMIS.MoveNext
                Loop
            Else
                XXX = 1
            End If
            Set rsPMIS = Nothing

            'CHECK PRESENT TRANSACTION
            Set rsPMIS = New ADODB.Recordset
            Set rsPMIS = gconDMIS.Execute("SELECT TRANNO FROM PMIS_ORD_HD WHERE (TRANTYPE = 'RIV' OR TRANTYPE = 'ADB') AND (STATUS = 'P' OR STATUS = 'B') AND RONO = " & N2Str2Null(txtRep_Or) & " AND TYPE = " & N2Str2Null(XTYPE) & "")
            If Not (rsPMIS.BOF And rsPMIS.EOF) Then
                Do While Not rsPMIS.EOF
                    Set rsDet = New ADODB.Recordset
                    Set rsDet = gconDMIS.Execute("SELECT STOCK_ORD FROM PMIS_TDAYTRAN WHERE (TRANTYPE = 'RIV' OR TRANTYPE = 'ADB') AND TYPE = " & N2Str2Null(XTYPE) & " AND TRANNO = '" & Null2String(rsPMIS!TRANNO) & "' AND (STATUS = 'P' OR STATUS = 'B') AND STOCK_ORD = '" & Null2String(rsCSMS!DETCDE) & "'")
                    If Not (rsDet.BOF And rsDet.EOF) Then
                        XXX = 0
                        GoTo NEXT_ITEM
                    Else
                        XXX = 1
                    End If
                    rsPMIS.MoveNext
                Loop
            Else
                XXX = 1
            End If
            Set rsPMIS = Nothing

NEXT_ITEM:
            rsCSMS.MoveNext
        Loop
    End If
    CheckIfRoDetIsIssued = XXX
    Set rsCSMS = Nothing
End Function

Function CheckIfIssuedInRoDet(xLIVIL As Integer, XTYPE As String) As Integer
    Dim rsCSMS                                         As New ADODB.Recordset
    Dim rsPMIS                                         As New ADODB.Recordset
    Dim rsDet                                          As New ADODB.Recordset
    Dim EXIST_HIST                                     As Boolean
    Dim EXIST_PRES                                     As Boolean
    Dim EXIST_ADVN                                     As Boolean
    Dim xMSG                                           As String
    Dim xMSG2                                          As String
    Dim XXX                                            As Integer

    Set rsPMIS = New ADODB.Recordset
    Set rsPMIS = gconDMIS.Execute("SELECT TRANNO FROM PMIS_ORD_HIST WHERE (TRANTYPE = 'RIV' OR TRANTYPE = 'ADB') AND (STATUS = 'P' OR STATUS = 'B') AND RONO = " & N2Str2Null(txtRep_Or) & " AND TYPE = " & N2Str2Null(XTYPE) & "")
    If Not (rsPMIS.BOF And rsPMIS.EOF) Then
        Do While Not rsPMIS.EOF
            Set rsDet = New ADODB.Recordset
            Set rsDet = gconDMIS.Execute("SELECT STOCK_ORD FROM PMIS_DAYTRAN WHERE (TRANTYPE = 'RIV' OR TRANTYPE = 'ADB') AND TYPE = " & N2Str2Null(XTYPE) & " AND TRANNO = '" & Null2String(rsPMIS!TRANNO) & "' AND (STATUS = 'P' OR STATUS = 'B')")
            If Not (rsDet.BOF And rsDet.EOF) Then
                Do While Not rsDet.EOF
                    Set rsCSMS = gconDMIS.Execute("SELECT DETCDE FROM CSMS_RO_DET WHERE LIVIL = " & xLIVIL & " AND REP_OR = " & N2Str2Null(txtRep_Or) & " AND ROTYPE IS NULL")
                    If Not (rsCSMS.BOF And rsCSMS.EOF) Then
                        XXX = 0
                        xMSG2 = ""
                        GoTo NEXT_ITEM1
                    Else
                        xMSG2 = "Some Issuance are not yet imported to this Repair Order. Press F8 to Import the Latest Issuance from Parts Module"
                        XXX = 1
                    End If
                    Set rsCSMS = Nothing
                    rsDet.MoveNext
                Loop
            End If

NEXT_ITEM1:
            rsPMIS.MoveNext
        Loop
    End If
    Set rsPMIS = Nothing

    Set rsPMIS = New ADODB.Recordset
    Set rsPMIS = gconDMIS.Execute("SELECT TRANNO FROM PMIS_ORD_HD WHERE (TRANTYPE = 'RIV' OR TRANTYPE = 'ADB') AND (STATUS = 'P' OR STATUS = 'B') AND RONO = " & N2Str2Null(txtRep_Or) & " AND TYPE = " & N2Str2Null(XTYPE) & "")
    If Not (rsPMIS.BOF And rsPMIS.EOF) Then
        Do While Not rsPMIS.EOF
            Set rsDet = New ADODB.Recordset
            Set rsDet = gconDMIS.Execute("SELECT STOCK_ORD FROM PMIS_TDAYTRAN WHERE (TRANTYPE = 'RIV' OR TRANTYPE = 'ADB') AND TYPE = " & N2Str2Null(XTYPE) & " AND TRANNO = '" & Null2String(rsPMIS!TRANNO) & "' AND (STATUS = 'P' OR STATUS = 'B')")
            If Not (rsDet.BOF And rsDet.EOF) Then
                Do While Not rsDet.EOF
                    Set rsCSMS = gconDMIS.Execute("SELECT DETCDE FROM CSMS_RO_DET WHERE LIVIL = " & xLIVIL & " AND REP_OR = " & N2Str2Null(txtRep_Or) & " AND ROTYPE IS NULL AND DETCDE = " & N2Str2Null(rsDet!STOCK_ORD) & "")
                    If Not (rsCSMS.BOF And rsCSMS.EOF) Then
                        XXX = 0
                        xMSG2 = ""
                        GoTo NEXT_ITEM2
                    Else
                        XXX = 1
                        xMSG2 = "Some Issuance are not yet imported to this Repair Order. Press F8 to Import the Latest Issuance from Parts Module"
                    End If
                    Set rsCSMS = Nothing
                    rsDet.MoveNext
                Loop
            End If
NEXT_ITEM2:
            rsPMIS.MoveNext
        Loop
    End If
    Set rsPMIS = Nothing
    CheckIfIssuedInRoDet = XXX
End Function

Sub SetCustLimit()
    Dim rsALL_Customer_Credit                          As New ADODB.Recordset
    Set rsALL_Customer_Credit = gconDMIS.Execute("Select CREDITLIMIT,CREDITDAYS from ALL_Customer Where CUSCDE = '" & txtAcct_No.Text & "'")
    If Not rsALL_Customer_Credit.EOF And Not rsALL_Customer_Credit.BOF Then
        labCreditDays.Caption = Null2String(rsALL_Customer_Credit!CREDITDAYS) & " DAY(S)"
        labCreditLimit.Caption = ToDoubleNumber(N2Str2Zero(rsALL_Customer_Credit!CreditLimit))
    Else
        labCreditDays.Caption = ""
        labCreditDays.Caption = ""
    End If
End Sub

Sub SetAdres(CCC As String)
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select cuscde,customeradd,telephoneno from All_Customer where cuscde = '" & CCC & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtAddress.Text = Null2String(rsCustomer!CUSTOMERADD)
    Else
        txtAddress.Text = ""
    End If
    Set rsCustomer = Nothing
End Sub

Sub clearDetailsgrd()
    Dim i, r                                           As Integer
    grdDetails.Rows = 7: grdDetails.Row = 1
    For r = 0 To grdDetails.Rows - 1
        grdDetails.Row = r
        For i = 0 To grdDetails.Cols - 1
            grdDetails.Col = i: grdDetails.Text = ""
        Next
    Next
    grdDetails.Col = 0: grdDetails.Text = "No Entry": grdDetails.Col = 2
End Sub

Sub InitGrid()
    With grdDetails
        .Rows = 8
        .ColWidth(0) = 1350
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1

        .Row = 0
        .Col = 1: .Text = "Customer"
        .Col = 2: .Text = "Company"
        .Col = 3: .Text = "Sales"
        .Col = 4: .Text = "Warranty"
        .Col = 5: .Text = "Insurance"
        .Col = 6: .Text = "Discount"
        .Col = 7: .Text = "Vat"
        .Col = 8: .Text = "ID"
        .Col = 0

        .Row = 2: .Text = "Labor"
        .Row = 3: .Text = "Parts"
        .Row = 4: .Text = "Materials"
        .Row = 5: .Text = "Accessories"
        .Row = 6: .Text = "TOTAL"
        .Row = 7: .Text = "RO Amount"
    End With
    grdDetails.RemoveItem 1
End Sub

Sub FillDetails()
    Screen.MousePointer = 11

    JobInsTotal = N2Str2Zero(rsREPOR!PARTLABOR)
    PartsInsTotal = N2Str2Zero(rsREPOR!PARTPARTS)
    MatInsTotal = N2Str2Zero(rsREPOR!PARTMATERIALS)
    AccInsTotal = N2Str2Zero(rsREPOR!PARTACCESSORIES)
    INSTotal = JobInsTotal + PartsInsTotal + MatInsTotal + AccInsTotal

    If INSTotal > 0 Then
        chkParticipat.Enabled = False
    Else
        chkParticipat.Enabled = True
    End If

    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - INSTotal
    COMTotal = JobComTotal + PartsComTotal + MatComTotal + ACCComTotal
    SALESTotal = JobSalesTotal + PartsSalesTotal + MatSalesTotal + ACCSalesTotal
    WARTotal = JobWarTotal + PartsWarTotal + MatWarTotal + ACCWarTotal

    If COMTotal + SALESTotal + WARTotal > 0 Then
        If CheckIfHasSellingDealer(Null2String(rsREPOR!PLATE_NO)) = False Then
            CapInfo.VisualTheme = 3
            labInfo.Caption = "Warning: Printing has been disabled. Please Edit Vehicle Info and Identify Source Dealer."
            cmdPrint.Enabled = False
        Else
            CapInfo.VisualTheme = 2
            labInfo.Caption = "System Info..."
            cmdPrint.Enabled = True
        End If
    Else
        CapInfo.VisualTheme = 2
        labInfo.Caption = "System Info..."
        cmdPrint.Enabled = True
    End If
    DiscTotal = N2Str2Zero(rsREPOR!l_discount) + N2Str2Zero(rsREPOR!p_discount) + N2Str2Zero(rsREPOR!m_discount) + N2Str2Zero(rsREPOR!a_discount)
    VATTotal = N2Str2Zero(rsREPOR!l_taxval) + N2Str2Zero(rsREPOR!p_taxval) + N2Str2Zero(rsREPOR!m_taxval) + N2Str2Zero(rsREPOR!A_taxval)

    TOTJOBAMT = TOTJOBAMT - JobInsTotal
    TOTPARTSAMT = TOTPARTSAMT - PartsInsTotal
    TOTMATAMT = TOTMATAMT - MatInsTotal
    TOTACCAMT = TOTACCAMT - AccInsTotal

    InitGrid
    With grdDetails
        .Rows = 7
        .Col = 1
        If N2Str2Zero(rsREPOR!VAT_EXEMPT) = 1 Then
            .Row = 1: .Text = Format(TOTJOBAMT, MAXIMUM_DIGIT)
            .Row = 2: .Text = Format(TOTPARTSAMT, MAXIMUM_DIGIT)
            .Row = 3: .Text = Format(TOTMATAMT, MAXIMUM_DIGIT)
            .Row = 4: .Text = Format(TOTACCAMT, MAXIMUM_DIGIT)
        Else
            .Row = 1: .Text = Format(TOTJOBAMT - N2Str2Zero(rsREPOR!l_discount), MAXIMUM_DIGIT)
            .Row = 2: .Text = Format(TOTPARTSAMT - N2Str2Zero(rsREPOR!p_discount), MAXIMUM_DIGIT)
            .Row = 3: .Text = Format(TOTMATAMT - N2Str2Zero(rsREPOR!m_discount), MAXIMUM_DIGIT)
            .Row = 4: .Text = Format(TOTACCAMT - N2Str2Zero(rsREPOR!a_discount), MAXIMUM_DIGIT)
        End If
        'AXP232008100
        'BEFORE'RO AMOUNT SHOULD BE SAME ONLY TOTAL AMOUNT SHOULD HAVE - DISCOUNT:HGC
        .Row = 5: .Text = Format(ROTotal - DiscTotal, MAXIMUM_DIGIT)
        .Row = 6: .Text = Format((ROTotal + INSTotal + COMTotal + SALESTotal + WARTotal) - DiscTotal, MAXIMUM_DIGIT)
        'AFTER
        '.Row = 5: .Text = Format((ROTotal + INSTotal + COMTotal + SALESTotal + WARTotal) - DiscTotal, MAXIMUM_DIGIT)
        '.Row = 6: .Text = Format(ROTotal, MAXIMUM_DIGIT) '

        .Col = 2
        .Row = 1: .Text = Format(JobComTotal, MAXIMUM_DIGIT)
        .Row = 2: .Text = Format(PartsComTotal, MAXIMUM_DIGIT)
        .Row = 3: .Text = Format(MatComTotal, MAXIMUM_DIGIT)
        .Row = 4: .Text = Format(ACCComTotal, MAXIMUM_DIGIT)
        .Row = 5: .Text = Format(COMTotal, MAXIMUM_DIGIT)
        .Col = 3
        .Row = 1: .Text = Format(JobSalesTotal, MAXIMUM_DIGIT)
        .Row = 2: .Text = Format(PartsSalesTotal, MAXIMUM_DIGIT)
        .Row = 3: .Text = Format(MatSalesTotal, MAXIMUM_DIGIT)
        .Row = 4: .Text = Format(ACCSalesTotal, MAXIMUM_DIGIT)
        .Row = 5: .Text = Format(SALESTotal, MAXIMUM_DIGIT)
        .Col = 4
        .Row = 1: .Text = Format(JobWarTotal, MAXIMUM_DIGIT)
        .Row = 2: .Text = Format(PartsWarTotal, MAXIMUM_DIGIT)
        .Row = 3: .Text = Format(MatWarTotal, MAXIMUM_DIGIT)
        .Row = 4: .Text = Format(ACCWarTotal, MAXIMUM_DIGIT)
        .Row = 5: .Text = Format(WARTotal, MAXIMUM_DIGIT)
        .Col = 5
        .Row = 1: .Text = Format(JobInsTotal, MAXIMUM_DIGIT)
        .Row = 2: .Text = Format(PartsInsTotal, MAXIMUM_DIGIT)
        .Row = 3: .Text = Format(MatInsTotal, MAXIMUM_DIGIT)
        .Row = 4: .Text = Format(AccInsTotal, MAXIMUM_DIGIT)
        .Row = 5: .Text = Format(INSTotal, MAXIMUM_DIGIT)
        .Col = 6
        .Row = 1: .Text = Format(N2Str2Zero(rsREPOR!l_discount), MAXIMUM_DIGIT)
        .Row = 2: .Text = Format(N2Str2Zero(rsREPOR!p_discount), MAXIMUM_DIGIT)
        .Row = 3: .Text = Format(N2Str2Zero(rsREPOR!m_discount), MAXIMUM_DIGIT)
        .Row = 4: .Text = Format(N2Str2Zero(rsREPOR!a_discount), MAXIMUM_DIGIT)
        .Row = 5: .Text = Format(DiscTotal, MAXIMUM_DIGIT)
        .Col = 7
        
        Dim LaborVatAmt                                As Double
        Dim PartsVatAmt                                As Double
        Dim MatsVatAmt                                 As Double
        Dim AccVatAmt                                  As Double
        If N2Str2Zero(rsREPOR!VAT_EXEMPT) = 1 Then
            LaborVatAmt = 0
            PartsVatAmt = 0
            MatsVatAmt = 0
            AccVatAmt = 0
        Else
            'LaborVatAmt = ((TOTJOBAMT + JobWarTotal + JobInsTotal) - N2Str2Zero(rsREPOR!l_discount)) / 9.3333
            'PartsVatAmt = ((TOTPARTSAMT + PartsWarTotal + PartsInsTotal) - N2Str2Zero(rsREPOR!p_discount)) / 9.3333
            'MatsVatAmt = ((TOTMATAMT + MatWarTotal + MatInsTotal) - N2Str2Zero(rsREPOR!m_discount)) / 9.3333
            'AccVatAmt = ((TOTACCAMT + AccWarTotal + AccInsTotal) - N2Str2Zero(rsREPOR!a_discount)) / 9.3333

            LaborVatAmt = (((TOTJOBAMT + JobWarTotal + JobInsTotal) - N2Str2Zero(rsREPOR!l_discount)) / 1.12) * 0.12
            PartsVatAmt = (((TOTPARTSAMT + PartsWarTotal + PartsInsTotal) - N2Str2Zero(rsREPOR!p_discount)) / 1.12) * 0.12
            MatsVatAmt = (((TOTMATAMT + MatWarTotal + MatInsTotal) - N2Str2Zero(rsREPOR!m_discount)) / 1.12) * 0.12
            AccVatAmt = (((TOTACCAMT + ACCWarTotal + AccInsTotal) - N2Str2Zero(rsREPOR!a_discount)) / 1.12) * 0.12
        End If
        .Row = 1: .Text = Format(LaborVatAmt, MAXIMUM_DIGIT)
        .Row = 2: .Text = Format(PartsVatAmt, MAXIMUM_DIGIT)
        .Row = 3: .Text = Format(MatsVatAmt, MAXIMUM_DIGIT)
        .Row = 4: .Text = Format(AccVatAmt, MAXIMUM_DIGIT)
        VATTotal = LaborVatAmt + PartsVatAmt + MatsVatAmt + AccVatAmt
        'VATTotal = Format(LaborVatAmt, MAXIMUM_DIGIT) + Format(PartsVatAmt, MAXIMUM_DIGIT) + Format(MatsVatAmt, MAXIMUM_DIGIT) + Format(AccVatAmt, MAXIMUM_DIGIT)
        .Row = 5: .Text = Format(VATTotal, MAXIMUM_DIGIT)
    End With
    Screen.MousePointer = 0
End Sub

Sub FillJobs()
    Dim Item                                           As ListItem

    Me.lstJObs.Sorted = True: Me.lstJObs.ListItems.Clear: Me.lstJObs.Enabled = False
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select LINE_NO,detcde,detdsc,det_amt,wcode,discount_2,id from CSMS_RO_Det where " & _
        " rep_or = " & N2Str2Null(rsREPOR!REP_OR) & _
        " and livil = '1' " & _
        " order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Do While Not rsRO_DET.EOF
            Set Item = lstJObs.ListItems.Add(, , rsRO_DET!LINE_NO)
            Item.SubItems(1) = Null2String(rsRO_DET!DETCDE)
            Item.SubItems(2) = Null2String(rsRO_DET!DETDSC)
            Item.SubItems(3) = Format(NumericVal(rsRO_DET!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(4) = Null2String(rsRO_DET!wCode)
            Item.SubItems(5) = Format(NumericVal(rsRO_DET!Discount_2), MAXIMUM_DIGIT)
            Item.SubItems(6) = Null2String(rsRO_DET!ID)
            rsRO_DET.MoveNext
        Loop
        Me.lstJObs.Enabled = True: Me.lstJObs.Sorted = False: Me.lstJObs.Refresh
    End If


    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    kcnt = 0: JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where " & _
        " rep_or = " & N2Str2Null(rsREPOR!REP_OR) & _
        " and livil = '1' " & _
        " order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            kcnt = kcnt + 1
            If Null2String(rsRO_DET!wCode) = "C" Then
                JobComTotal = JobComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTJOBAMT = Round(TOTJOBAMT, 2)
    TOTJOBDISC = Round(TOTJOBDISC, 2)
    TOTJOBDISCVAL = Round(TOTJOBDISCVAL, 2)
    TOTJOBTAX = Round(TOTJOBTAX, 2)
End Sub

Sub FillParts()
    Dim Item                                           As ListItem

    Me.lstParts.Sorted = True: Me.lstParts.ListItems.Clear: Me.lstParts.Enabled = False
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2,id from CSMS_RO_Det where rep_or = '" & rsREPOR!REP_OR & "' and livil = '2' order by LINE_NO asc")

    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Do While Not rsRO_DET.EOF
            Set Item = lstParts.ListItems.Add(, , rsRO_DET!LINE_NO)
            Item.SubItems(1) = Null2String(rsRO_DET!DETCDE)
            Item.SubItems(2) = Null2String(rsRO_DET!DETDSC)
            Item.SubItems(3) = NumericVal(rsRO_DET!detvol)
            Item.SubItems(4) = Format(NumericVal(rsRO_DET!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(5) = Format(NumericVal(rsRO_DET!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(6) = Null2String(rsRO_DET!wCode)
            Item.SubItems(7) = Format(NumericVal(rsRO_DET!Discount_2), MAXIMUM_DIGIT)
            Item.SubItems(8) = Null2String(rsRO_DET!ID)
            rsRO_DET.MoveNext
        Loop
        Me.lstParts.Sorted = False: Me.lstParts.Refresh: Me.lstParts.Enabled = True
    End If

    TOTPARTSAMT = 0: TOTPARTSDISC = 0: TOTPARTSDISCVAL = 0: TOTPARTSTAX = 0
    Pcnt = 0: PartsComTotal = 0: PartsSalesTotal = 0: PartsWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = '" & rsREPOR!REP_OR & "' and livil = '2' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsRO_DET.EOF
            Pcnt = Pcnt + 1
            If Null2String(rsRO_DET!wCode) = "C" Then
                PartsComTotal = PartsComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then PartsWarTotal = PartsWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTPARTSAMT = Round(TOTPARTSAMT, 2)
    TOTPARTSDISC = Round(TOTPARTSDISC, 2)
    TOTPARTSDISCVAL = Round(TOTPARTSDISCVAL, 2)
    TOTPARTSTAX = Round(TOTPARTSTAX, 2)
End Sub

Sub FillMaterials()
    '    Me.lstMaterials.Sorted = True: Me.lstMaterials.ListItems.Clear: Me.lstMaterials.Enabled = False
    '    Set rsRO_DET = New ADODB.Recordset
    '    Set rsRO_DET = gconDMIS.Execute("select LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2,id from CSMS_RO_Det where rep_or = '" & rsREPOR!REP_OR & "' and livil = '3' order by LINE_NO asc")
    '    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
    '        Listview_Loadval Me.lstMaterials.ListItems, rsRO_DET
    '        Me.lstMaterials.Sorted = False: Me.lstMaterials.Refresh: Me.lstMaterials.Enabled = True
    '    End If

    Dim Item                                           As ListItem

    Me.lstMaterials.Sorted = True: Me.lstMaterials.ListItems.Clear: Me.lstMaterials.Enabled = False
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2,id from CSMS_RO_Det where rep_or = '" & rsREPOR!REP_OR & "' and livil = '3' order by LINE_NO asc")

    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Do While Not rsRO_DET.EOF
            Set Item = lstMaterials.ListItems.Add(, , rsRO_DET!LINE_NO)
            Item.SubItems(1) = Null2String(rsRO_DET!DETCDE)
            Item.SubItems(2) = Null2String(rsRO_DET!DETDSC)
            Item.SubItems(3) = NumericVal(rsRO_DET!detvol)
            Item.SubItems(4) = Format(NumericVal(rsRO_DET!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(5) = Format(NumericVal(rsRO_DET!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(6) = Null2String(rsRO_DET!wCode)
            Item.SubItems(7) = Format(NumericVal(rsRO_DET!Discount_2), MAXIMUM_DIGIT)
            Item.SubItems(8) = Null2String(rsRO_DET!ID)
            rsRO_DET.MoveNext
        Loop
        Me.lstMaterials.Sorted = False: Me.lstMaterials.Refresh: Me.lstMaterials.Enabled = True
    End If

    TOTMATAMT = 0: TOTMATDISC = 0: TOTMATDISCVAL = 0: TOTMATTAX = 0
    Mcnt = 0: MatComTotal = 0: MatSalesTotal = 0: MatWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = '" & rsREPOR!REP_OR & "' and livil = '3' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            Mcnt = Mcnt + 1
            If Null2String(rsRO_DET!wCode) = "C" Then
                MatComTotal = MatComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then MatSalesTotal = MatSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then MatWarTotal = MatWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTMATAMT = TOTMATAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTMATDISC = TOTMATDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTMATTAX = TOTMATTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTMATAMT = Round(TOTMATAMT, 2): TOTMATDISC = Round(TOTMATDISC, 2): TOTMATDISCVAL = Round(TOTMATDISCVAL, 2): TOTMATTAX = Round(TOTMATTAX, 2)
End Sub

Sub FillAccessories()
    '    Me.lstAccessories.Sorted = True: Me.lstAccessories.ListItems.Clear: Me.lstAccessories.Enabled = False
    '    Set rsRO_DET = New ADODB.Recordset
    '    Set rsRO_DET = gconDMIS.Execute("select LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2,id from CSMS_RO_Det where rep_or = '" & rsREPOR!REP_OR & "' and livil = '4' order by LINE_NO asc")
    '    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
    '        Listview_Loadval Me.lstAccessories.ListItems, rsRO_DET
    '        Me.lstAccessories.Sorted = False: Me.lstAccessories.Refresh: Me.lstAccessories.Enabled = True
    '    End If
    Dim Item                                           As ListItem

    Me.lstAccessories.Sorted = True: Me.lstAccessories.ListItems.Clear: Me.lstAccessories.Enabled = False
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2,id from CSMS_RO_Det where rep_or = '" & rsREPOR!REP_OR & "' and livil = '4' order by LINE_NO asc")

    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Do While Not rsRO_DET.EOF
            Set Item = lstAccessories.ListItems.Add(, , rsRO_DET!LINE_NO)
            Item.SubItems(1) = Null2String(rsRO_DET!DETCDE)
            Item.SubItems(2) = Null2String(rsRO_DET!DETDSC)
            Item.SubItems(3) = NumericVal(rsRO_DET!detvol)
            Item.SubItems(4) = Format(NumericVal(rsRO_DET!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(5) = Format(NumericVal(rsRO_DET!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(6) = Null2String(rsRO_DET!wCode)
            Item.SubItems(7) = Format(NumericVal(rsRO_DET!Discount_2), MAXIMUM_DIGIT)
            Item.SubItems(8) = Null2String(rsRO_DET!ID)
            rsRO_DET.MoveNext
        Loop
        Me.lstAccessories.Sorted = False: Me.lstAccessories.Refresh: Me.lstAccessories.Enabled = True
    End If

    TOTACCAMT = 0: TOTACCDISC = 0: TOTACCDISCVAL = 0: TOTACCTAX = 0
    Acnt = 0: ACCComTotal = 0: ACCSalesTotal = 0: ACCWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = '" & rsREPOR!REP_OR & "' and livil = '4' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            Acnt = Acnt + 1
            If Null2String(rsRO_DET!wCode) = "C" Then
                ACCComTotal = ACCComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then ACCSalesTotal = ACCSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then ACCWarTotal = ACCWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTACCAMT = TOTACCAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTACCDISC = TOTACCDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTACCDISCVAL = TOTACCDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTACCTAX = TOTACCTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTACCAMT = Round(TOTACCAMT, 2): TOTACCDISC = Round(TOTACCDISC, 2): TOTACCDISCVAL = Round(TOTACCDISCVAL, 2): TOTACCTAX = Round(TOTACCTAX, 2)
End Sub

Sub SendToBack()
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
    Picture1.Enabled = True: Picture5.Enabled = True: pic3.Enabled = True: Frame2.Enabled = True
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

    cmdAddJobs.ZOrder 1: fraAddJobs.ZOrder 1: fraAddJobs.Enabled = False
    cmdAddParts.ZOrder 1: fraAddParts.ZOrder 1: fraAddParts.Enabled = False
    cmdAddMaterials.ZOrder 1: fraAddMaterials.ZOrder 1: fraAddMaterials.Enabled = False
    cmdAddAccessories.ZOrder 1: fraAddAccessories.ZOrder 1: fraAddAccessories.Enabled = False
    Picture5.ZOrder 1: Picture6.ZOrder 1: picCustLimit.ZOrder 1

End Sub

Sub SendToBackDisc()
    cmdDiscount.ZOrder 1: fraDiscount.ZOrder 1: txtDiscAmt.Text = 0
End Sub

Sub SendToFrontDisc()
    cmdDiscount.ZOrder 0
    fraDiscount.ZOrder 0
    
    txtDiscAmt.Text = 5
    
    On Error Resume Next
    txtDiscAmt.SetFocus
End Sub

Sub SendToBackRel()
    cmdRelBut.ZOrder 1: fraRelBut.ZOrder 1: txtReleaseDate.Text = ""
End Sub

Sub SendToFrontRel()
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A RELEASED DATE
    Picture1.Enabled = False: Picture5.Enabled = False: pic3.Enabled = False: Frame2.Enabled = False
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

    cmdRelBut.ZOrder 0: fraRelBut.ZOrder 0: txtReleaseDate.Text = LOGDATE
End Sub

Sub SendToBackBill()
    cmdBillBut.ZOrder 1: fraBillBut.ZOrder 1: txtInvoiceNumber.Text = ""
End Sub

Sub SendToFrontBill()
    cmdBillBut.ZOrder 0: fraBillBut.ZOrder 0: txtInvoiceDate.Text = LOGDATE
    txtDateReleased.Text = LOGDATE: txtInvoiceNumber.Text = ""
    On Error Resume Next: txtInvoiceNumber.SetFocus
End Sub

Sub SetTotalParticipation()
    txtPartTotal.Text = ToDoubleNumber(NumericVal(txtPartLabor.Text) + NumericVal(txtPartParts.Text) + NumericVal(txtPartMaterials.Text) + NumericVal(txtPartAccessories.Text))
    If chkAllowManDist.Value = 1 Then
        txtLOAAmount.Text = ToDoubleNumber(txtPartTotal.Text)
        If NumericVal(txtLOAAmount.Text) > ROTotal + INSTotal Then
            MsgBox "Warning: LOA Amount should not Exceed Repair Order Total Amount.", vbCritical, "Not Allowed!"
            txtLOAAmount.Text = NumericVal(txtPartTotal.Text)
            cmdPartSave.Enabled = False
            Exit Sub
        Else
            txtPartTotal.Text = ToDoubleNumber(NumericVal(txtPartLabor.Text) + NumericVal(txtPartParts.Text) + NumericVal(txtPartMaterials.Text) + NumericVal(txtPartAccessories.Text))
            txtLOAAmount.Text = ToDoubleNumber(txtPartTotal.Text)
            cmdPartSave.Enabled = True
        End If
    Else
        txtPartTotal.Text = ToDoubleNumber(NumericVal(txtPartLabor.Text) + NumericVal(txtPartParts.Text) + NumericVal(txtPartMaterials.Text) + NumericVal(txtPartAccessories.Text))
        cmdPartSave.Enabled = True
    End If
End Sub

Sub SetCustomer()
    Dim rsCustomer                                     As New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtAcct_No.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtNiym.Text = Null2String(rsCustomer!AcctName)
        txtAddress.Text = Null2String(rsCustomer!CUSTOMERADD)
    End If
End Sub

Sub checkIFBlalnk()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim Item                                           As ListItem
    Dim theRo                                          As String
    Dim cnt                                            As Integer
    theRo = Trim(txtRep_Or.Text)

    SQL = "SELECT TECHCODE FROM CSMS_Ro_Det WHERE LIVIL = '1' AND REP_OR = '" & theRo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListView1.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set Item = ListView1.ListItems.Add(, , cnt)
            Item.SubItems(1) = Null2String(!TechCode)

            If Item.SubItems(1) = "" Then
                MsgBox "No Assigned Technician on Job,RO Cannot Be Billed!", vbExclamation, "Information!"
                flag = True
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub ImportPartsFromSublet(XXX As String)
    'UPDATE BY: JUN
    'DATE: 06/05/2008
    'DESCRIPTION: IMPORT THE SUBLET PARTS POSTED ONLY FROM PO CREATION TO CSMS_RO_DET TO THE CORRESPANDING RO NUMBER

    Dim rsCheckIFPosted                                As New ADODB.Recordset

    Set rsCheckIFPosted = gconDMIS.Execute("Select PO_NO,STATUS from CSMS_PO_HD where RO_NO ='" & txtRep_Or.Text & "' and STATUS ='P'")
    If Not rsCheckIFPosted.EOF And Not rsCheckIFPosted.BOF Then
        Do While Not rsCheckIFPosted.EOF
            Dim getPO                                  As String
            getPO = N2Str2Null(rsCheckIFPosted!PO_NO)

            Dim rsPartsSublet                          As ADODB.Recordset

            Dim sRep_or                                As String
            Dim sROTYPE                                As String
            Dim sJOBTYPE                               As String
            Dim sLIVIL                                 As String
            Dim sLINE_NO                               As String
            Dim sDETAMT                                As Double
            Dim sDETCDE                                As String
            Dim sDETDSC                                As String
            Dim sTECHNICIAN                            As String
            Dim swCode                                 As String
            Dim sTAXRATE                               As Double
            Dim sTAXVAL                                As Double
            Dim sDETAIL                                As String
            Dim sDET_AMT                               As Double
            Dim sUSERCODE                              As String
            Dim sSAVEDATE                              As String
            Dim sTechCode                              As String
            Dim sDONE                                  As String
            Dim sPONO                                  As String

            Set rsPartsSublet = gconDMIS.Execute("Select Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETAMT,DETCDE,DETDSC,TECHNICIAN,WCODE,TAXRATE,TAXVAL,DETAIL,DET_AMT,USERCODE,SAVEDATE,TECHCODE,DONE,PO_NO from CSMS_PO_DT where Rep_or ='" & txtRep_Or.Text & "' and LIVIL = '2' and PO_NO =" & getPO & "")
            If Not rsPartsSublet.EOF And Not rsPartsSublet.BOF Then
                Do While Not rsPartsSublet.EOF
                    sRep_or = N2Str2Null(rsPartsSublet!REP_OR)
                    sROTYPE = N2Str2Null(rsPartsSublet!ROTYPE)
                    sJOBTYPE = N2Str2Null(rsPartsSublet!JOBTYPE)
                    sLIVIL = N2Str2Null(rsPartsSublet!LIVIL)
                    sLINE_NO = N2Str2Null(rsPartsSublet!LINE_NO)
                    sDETAMT = NumericVal(rsPartsSublet!DETAMT)
                    sDETCDE = N2Str2Null(rsPartsSublet!DETCDE)
                    If sLIVIL = "'2'" Or sLIVIL = "'3'" Then
                        sDETDSC = N2Str2Null(rsPartsSublet!Detail)
                    Else
                        sDETDSC = N2Str2Null(rsPartsSublet!DETDSC)
                    End If
                    sTECHNICIAN = N2Str2Null(rsPartsSublet!Technician)
                    swCode = N2Str2Null(rsPartsSublet!wCode)
                    sTAXRATE = NumericVal(rsPartsSublet!taxrate) * 100
                    sTAXVAL = NumericVal(rsPartsSublet!TAXVAL)
                    If sLIVIL = "'2'" Or sLIVIL = "'3'" Then
                        sDETAIL = "NULL"
                    Else
                        sDETAIL = N2Str2Null(rsPartsSublet!Detail)
                    End If
                    sDET_AMT = NumericVal(rsPartsSublet!DET_AMT)
                    sUSERCODE = N2Str2Null(rsPartsSublet!USERCODE)
                    sSAVEDATE = N2Str2Null(rsPartsSublet!savedate)
                    sTechCode = N2Str2Null(rsPartsSublet!TechCode)
                    sPONO = N2Str2Null(rsPartsSublet!PO_NO)

                    SQL_STATEMENT = "Insert into CSMS_ro_det" & _
                        "(Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETCDE,DETDSC,TECHNICIAN,DETAMT,WCODE,TAXRATE,TAXVAL,DETAIL,DET_AMT,USERCDE,SAVEDATE,TECHCODE,SUBPOCODE)" & _
                        "VALUES(" & sRep_or & _
                        "," & sROTYPE & _
                        "," & sJOBTYPE & _
                        "," & sLIVIL & _
                        "," & sLINE_NO & _
                        "," & sDETCDE & _
                        "," & sDETDSC & _
                        "," & sTECHNICIAN & _
                        "," & sDETAMT & _
                        "," & swCode & _
                        "," & sTAXRATE & _
                        "," & sTAXVAL & _
                        "," & sDETAIL & _
                        "," & sDET_AMT & _
                        "," & sUSERCODE & _
                        "," & sSAVEDATE & _
                        "," & sTechCode & _
                        "," & sPONO & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    'NEW LOG AUDIT----------------------------------------------------------------------
                        Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "JOB CODE: " & Null2String(rsPartsSublet!DETCDE), "", "")
                    'NEW LOG AUDIT----------------------------------------------------------------------
                    rsPartsSublet.MoveNext
                Loop
            End If

            rsCheckIFPosted.MoveNext
        Loop
    End If
End Sub

Sub importMaterialsFromSublet(XXX As String)
    'UPDATED BY: JUN
    'DATE: 06/05/2008
    'DESCRIPTION: IMPORT THE SUBLET MATERIALS POSTED ONLY FROM PO CREATION TO CSMS_RO_DET TO THE CORRESPANDING RO NUMBER
    Dim rsCheckIFPosted                                As New ADODB.Recordset

    Set rsCheckIFPosted = gconDMIS.Execute("Select PO_NO,STATUS from CSMS_PO_HD where RO_NO ='" & txtRep_Or.Text & "' and STATUS ='P'")
    If Not rsCheckIFPosted.EOF And Not rsCheckIFPosted.BOF Then
        Do While Not rsCheckIFPosted.EOF
            Dim getPONUm                               As String
            getPONUm = N2Str2Null(rsCheckIFPosted!PO_NO)
            Dim rsMaterialsFromPO                      As ADODB.Recordset

            Dim qRep_or                                As String
            Dim qROTYPE                                As String
            Dim qJOBTYPE                               As String
            Dim qLIVIL                                 As String
            Dim qLINE_NO                               As String
            Dim qDETAMT                                As Double
            Dim qDETCDE                                As String
            Dim qDETDSC                                As String
            Dim qTECHNICIAN                            As String
            Dim qwCode                                 As String
            Dim qTAXRATE                               As Double
            Dim qTAXVAL                                As Double
            Dim qDETAIL                                As String
            Dim qDET_AMT                               As Double
            Dim qUSERCODE                              As String
            Dim qSAVEDATE                              As String
            Dim qTechCode                              As String
            Dim qDONE                                  As String
            Dim qPONO                                  As String

            Set rsMaterialsFromPO = gconDMIS.Execute("Select Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETAMT,DETCDE,DETDSC,TECHNICIAN,WCODE,TAXRATE,TAXVAL,DETAIL,DET_AMT,USERCODE,SAVEDATE,TECHCODE,DONE,PO_NO from CSMS_PO_DT where Rep_or ='" & txtRep_Or.Text & "' and LIVIL = '3' and PO_NO =" & getPONUm & "")
            If Not rsMaterialsFromPO.EOF And Not rsMaterialsFromPO.BOF Then
                Do While Not rsMaterialsFromPO.EOF
                    qRep_or = N2Str2Null(rsMaterialsFromPO!REP_OR)
                    qROTYPE = N2Str2Null(rsMaterialsFromPO!ROTYPE)
                    qJOBTYPE = N2Str2Null(rsMaterialsFromPO!JOBTYPE)
                    qLIVIL = N2Str2Null(rsMaterialsFromPO!LIVIL)
                    qLINE_NO = N2Str2Null(rsMaterialsFromPO!LINE_NO)
                    qDETAMT = NumericVal(rsMaterialsFromPO!DETAMT)
                    qDETCDE = N2Str2Null(rsMaterialsFromPO!DETCDE)
                    If qLIVIL = "'2'" Or qLIVIL = "'3'" Then
                        qDETDSC = N2Str2Null(rsMaterialsFromPO!Detail)
                    Else
                        qDETDSC = N2Str2Null(rsMaterialsFromPO!DETDSC)
                    End If
                    qTECHNICIAN = N2Str2Null(rsMaterialsFromPO!Technician)
                    qwCode = N2Str2Null(rsMaterialsFromPO!wCode)
                    qTAXRATE = NumericVal(rsMaterialsFromPO!taxrate) * 100
                    qTAXVAL = NumericVal(rsMaterialsFromPO!TAXVAL)
                    If qLIVIL = "'2'" Or qLIVIL = "'3'" Then
                        qDETAIL = "NULL"
                    Else
                        qDETAIL = N2Str2Null(rsMaterialsFromPO!Detail)
                    End If
                    qDET_AMT = NumericVal(rsMaterialsFromPO!DET_AMT)
                    qUSERCODE = N2Str2Null(rsMaterialsFromPO!USERCODE)
                    qSAVEDATE = N2Str2Null(rsMaterialsFromPO!savedate)
                    qTechCode = N2Str2Null(rsMaterialsFromPO!TechCode)
                    qPONO = N2Str2Null(rsMaterialsFromPO!PO_NO)

                    gconDMIS.Execute "Insert into CSMS_ro_det" & _
                        "(Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETCDE,DETDSC,TECHNICIAN,DETAMT,WCODE,TAXRATE,TAXVAL,DETAIL,DET_AMT,USERCDE,SAVEDATE,TECHCODE,SUBPOCODE)" & _
                        "VALUES(" & qRep_or & _
                        "," & qROTYPE & _
                        "," & qJOBTYPE & _
                        "," & qLIVIL & _
                        "," & qLINE_NO & _
                        "," & qDETCDE & _
                        "," & qDETDSC & _
                        "," & qTECHNICIAN & _
                        "," & qDETAMT & _
                        "," & qwCode & _
                        "," & qTAXRATE & _
                        "," & qTAXVAL & _
                        "," & qDETAIL & _
                        "," & qDET_AMT & _
                        "," & qUSERCODE & _
                        "," & qSAVEDATE & _
                        "," & qTechCode & _
                        "," & qPONO & ")"

                    rsMaterialsFromPO.MoveNext
                Loop
            End If

            rsCheckIFPosted.MoveNext
        Loop
    End If
End Sub

Private Sub capFollow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picFollowUpResult.Top = 7110 Then
        picFollowUpResult.Top = 5250
        picFollowUpResult.ZOrder 0
        capFollow.Caption = "F1 - Enter Notes after follow up                                X"
    Else
        picFollowUpResult.Top = 7110
        capFollow.Caption = "F1 - Enter Notes after follow up"
    End If
End Sub

Private Sub capInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picInfo.Top = 7110 Then
        picInfo.Top = 5250
        picInfo.ZOrder 0
        CapInfo.Caption = "Repair Order Info                                                        X"
    Else
        picInfo.Top = 7110
        CapInfo.Caption = "Repair Order Info"
    End If
End Sub

Private Sub capSUG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Picture4.Top = 7110 Then
        Picture4.Top = 5250
        Picture4.ZOrder 0
        capSUG.Caption = "Suggestion/ Recommendation                                 X"
    Else
        Picture4.Top = 7110
        capSUG.Caption = "Suggestion/ Recommendation"
    End If
End Sub

Private Sub cboAccChargeTo_Change()
    If cboAccChargeTo.Text <> "" Then
        txtAccDiscount.Text = 0
        txtAccDiscount.Enabled = False
    Else
        txtAccDiscount.Enabled = True
    End If
    If cboAccChargeTo.Text = "C" Or cboAccChargeTo.Text = "S" Then
        cboAcctCodeAccessories.Enabled = True
        If PREV_ACCESSORIES_CHARGE_TO <> "C" And PREV_ACCESSORIES_CHARGE_TO <> "S" Then
            If NumericVal(txtAccUnitPrice.Text) > 0 Then
                txtAccUnitPrice.Text = ToDoubleNumber(NumericVal(txtAccUnitPrice) / 1.12)
            End If
        End If
    Else
        cboAcctCodeAccessories.ListIndex = -1
        cboAcctCodeAccessories.Enabled = False
        If PREV_ACCESSORIES_CHARGE_TO = "C" Or PREV_ACCESSORIES_CHARGE_TO = "S" Then
            If NumericVal(txtAccUnitPrice.Text) > 0 Then
                txtAccUnitPrice.Text = ToDoubleNumber(NumericVal(txtAccUnitPrice) * 1.12)
            End If
        End If
    End If
    PREV_ACCESSORIES_CHARGE_TO = cboAccChargeTo
End Sub

Private Sub cboAccChargeTo_Click()
    If cboAccChargeTo.Text <> "" Then
        txtAccDiscount.Text = 0
        txtAccDiscount.Enabled = False
    Else
        txtAccDiscount.Enabled = True
    End If
    If cboAccChargeTo.Text = "C" Or cboAccChargeTo.Text = "S" Then
        cboAcctCodeAccessories.Enabled = True
        If PREV_ACCESSORIES_CHARGE_TO <> "C" And PREV_ACCESSORIES_CHARGE_TO <> "S" Then
            If NumericVal(txtAccUnitPrice.Text) > 0 Then
                txtAccUnitPrice.Text = ToDoubleNumber(NumericVal(txtAccUnitPrice) / 1.12)
            End If
        End If
    Else
        cboAcctCodeAccessories.ListIndex = -1
        cboAcctCodeAccessories.Enabled = False
        If PREV_ACCESSORIES_CHARGE_TO = "C" Or PREV_ACCESSORIES_CHARGE_TO = "S" Then
            If NumericVal(txtAccUnitPrice.Text) > 0 Then
                txtAccUnitPrice.Text = ToDoubleNumber(NumericVal(txtAccUnitPrice) * 1.12)
            End If
        End If
    End If
    PREV_ACCESSORIES_CHARGE_TO = cboAccChargeTo
End Sub

Private Sub cboAccCode_Change()
    If cboAccCode.Text <> "" Then cboAccessories.Text = SetAccDisc(cboAccCode.Text)
    txtAccUnitPrice.Text = SetAccPrice(cboAccCode.Text)
    txtMatPOCode.Text = SetAccPOCode(cboAccCode.Text)
    txtAccAmount.Text = txtAccUnitPrice.Text
End Sub

Private Sub cboAccessories_Change()
    If cboAccessories.Text <> "" Then
        cboAccCode.Text = SetAccCode(cboAccessories.Text)
        txtAccUnitPrice.Text = SetAccPrice(cboAccCode.Text)
        txtMatPOCode.Text = SetAccPOCode(cboAccCode.Text)
        txtAccAmount.Text = txtAccUnitPrice.Text
    End If
End Sub

Private Sub cboAcctCodeAccessories_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboAcctCodeAccessories.ListIndex = -1
End Sub

Private Sub cboAcctCodeLabor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboAcctCodeLabor.ListIndex = -1
End Sub

Private Sub cboAcctCodeMaterials_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboAcctCodeMaterials.ListIndex = -1
End Sub

Private Sub cboAcctCodeParts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboAcctCodeParts.ListIndex = -1
End Sub

Private Sub cboChargeTo_Change()
    If cboChargeTo.Text <> "" Then
        txtPartDiscount.Text = 0
        txtPartDiscount.Enabled = False
    Else
        txtPartDiscount.Enabled = True
    End If
    If cboChargeTo.Text = "C" Or cboChargeTo.Text = "S" Then
        cboAcctCodeParts.Enabled = True
        If PREV_PARTS_CHARGE_TO <> "C" And PREV_PARTS_CHARGE_TO <> "S" Then
            If NumericVal(txtUnitPrice.Text) > 0 Then
                txtUnitPrice.Text = ToDoubleNumber(NumericVal(txtUnitPrice) / 1.12)
            End If
        End If
    Else
        cboAcctCodeParts.ListIndex = -1
        cboAcctCodeParts.Enabled = False
        If PREV_PARTS_CHARGE_TO = "C" Or PREV_PARTS_CHARGE_TO = "S" Then
            If NumericVal(txtUnitPrice.Text) > 0 Then
                txtUnitPrice.Text = ToDoubleNumber(NumericVal(txtUnitPrice) * 1.12)
            End If
        End If
    End If
    PREV_PARTS_CHARGE_TO = cboChargeTo
End Sub

Private Sub cboChargeTo_Click()
    If cboChargeTo.Text <> "" Then
        txtPartDiscount.Text = 0
        txtPartDiscount.Enabled = False
    Else
        txtPartDiscount.Enabled = True
    End If
    If cboChargeTo.Text = "C" Or cboChargeTo.Text = "S" Then
        cboAcctCodeParts.Enabled = True
        If PREV_PARTS_CHARGE_TO <> "C" And PREV_PARTS_CHARGE_TO <> "S" Then
            If NumericVal(txtUnitPrice.Text) > 0 Then
                txtUnitPrice.Text = ToDoubleNumber(NumericVal(txtUnitPrice) / 1.12)
            End If
        End If
    Else
        cboAcctCodeParts.ListIndex = -1
        cboAcctCodeParts.Enabled = False
        If PREV_PARTS_CHARGE_TO = "C" Or PREV_PARTS_CHARGE_TO = "S" Then
            If NumericVal(txtUnitPrice.Text) > 0 Then
                txtUnitPrice.Text = ToDoubleNumber(NumericVal(txtUnitPrice) * 1.12)
            End If
        End If
    End If
    PREV_PARTS_CHARGE_TO = cboChargeTo
End Sub

Private Sub cboChargeTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboChargeTo.ListIndex = -1
End Sub

Private Sub cboJcode_Change()
    If optByCode.Value = True Then
        cboJobCode.Text = setJobDesc(cboJcode.Text)
        txtJobPostCode.Text = setJobPOcode(cboJobCode.Text)
        txtJobRate.Text = setJobRate(cboJobCode.Text)
        If AddorEdit = "ADD" Then txtJobDetail.Text = setJobDetail(cboJobCode.Text)
    End If
End Sub

Private Sub cboJobChargeTo_Change()
    If cboJobChargeTo.Text <> "" Then
        txtJobDiscount.Text = 0
        txtJobDiscount.Enabled = False
    Else
        txtJobDiscount.Enabled = True
    End If
    If cboJobChargeTo.Text = "C" Or cboJobChargeTo.Text = "S" Then
        cboAcctCodeLabor.Enabled = True
        If PREV_LABOR_CHARGE_TO <> "C" And PREV_LABOR_CHARGE_TO <> "S" Then
            If NumericVal(txtJobRate.Text) > 0 Then
                txtJobRate.Text = ToDoubleNumber(NumericVal(txtJobRate) / 1.12)
            End If
        End If
    Else
        cboAcctCodeLabor.ListIndex = -1
        cboAcctCodeLabor.Enabled = False
        If PREV_LABOR_CHARGE_TO = "C" Or PREV_LABOR_CHARGE_TO = "S" Then
            If NumericVal(txtJobRate.Text) > 0 Then
                txtJobRate.Text = ToDoubleNumber(NumericVal(txtJobRate) * 1.12)
            End If
        End If
    End If
    PREV_LABOR_CHARGE_TO = cboJobChargeTo
End Sub

Private Sub cboJobChargeTo_Click()
    If cboJobChargeTo.Text <> "" Then
        txtJobDiscount.Text = 0
        txtJobDiscount.Enabled = False
    Else
        txtJobDiscount.Enabled = True
    End If
    If cboJobChargeTo.Text = "C" Or cboJobChargeTo.Text = "S" Then
        cboAcctCodeLabor.Enabled = True
        If PREV_LABOR_CHARGE_TO <> "C" And PREV_LABOR_CHARGE_TO <> "S" Then
            If NumericVal(txtJobRate.Text) > 0 Then
                txtJobRate.Text = ToDoubleNumber(NumericVal(txtJobRate) / 1.12)
            End If
        End If
    Else
        cboAcctCodeLabor.ListIndex = -1
        cboAcctCodeLabor.Enabled = False
        If PREV_LABOR_CHARGE_TO = "C" Or PREV_LABOR_CHARGE_TO = "S" Then
            If NumericVal(txtJobRate.Text) > 0 Then
                txtJobRate.Text = ToDoubleNumber(NumericVal(txtJobRate) * 1.12)
            End If
        End If
    End If
    PREV_LABOR_CHARGE_TO = cboJobChargeTo
End Sub

Private Sub cboJobChargeTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboJobChargeTo.ListIndex = -1
End Sub

Private Sub cboJobCode_Change()
    '    If optByDescription.Value = True Then
    '        cboJcode.Text = setJobCode(cboJobCode.Text)
    '        txtJobPostCode.Text = setJobPOcode(cboJobCode.Text)
    '        txtJobRate.Text = setJobRate(cboJobCode.Text)
    '        If AddorEdit = "ADD" Then txtJobDetail.Text = setJobDetail(cboJobCode.Text)
    '    End If
End Sub

Private Sub cboMatChargeTo_Change()
    If cboMatChargeTo.Text <> "" Then
        txtMatDiscount.Text = 0
        txtMatDiscount.Enabled = False
    Else
        txtMatDiscount.Enabled = True
    End If
    If cboMatChargeTo.Text = "C" Or cboMatChargeTo.Text = "S" Then
        cboAcctCodeMaterials.Enabled = True
        If PREV_MATERIALS_CHARGE_TO <> "C" And PREV_MATERIALS_CHARGE_TO <> "S" Then
            If NumericVal(txtMatUnitPrice.Text) > 0 Then
                txtMatUnitPrice.Text = ToDoubleNumber(NumericVal(txtMatUnitPrice) / 1.12)
            End If
        End If
    Else
        cboAcctCodeMaterials.ListIndex = -1
        cboAcctCodeMaterials.Enabled = False
        If PREV_MATERIALS_CHARGE_TO = "C" Or PREV_MATERIALS_CHARGE_TO = "S" Then
            If NumericVal(txtMatUnitPrice.Text) > 0 Then
                txtMatUnitPrice.Text = ToDoubleNumber(NumericVal(txtMatUnitPrice) * 1.12)
            End If
        End If
    End If
    PREV_MATERIALS_CHARGE_TO = cboMatChargeTo
End Sub

Private Sub cboMatChargeTo_Click()
    If cboMatChargeTo.Text <> "" Then
        txtMatDiscount.Text = 0
        txtMatDiscount.Enabled = False
    Else
        txtMatDiscount.Enabled = True
    End If
    If cboMatChargeTo.Text = "C" Or cboMatChargeTo.Text = "S" Then
        cboAcctCodeMaterials.Enabled = True
        If PREV_MATERIALS_CHARGE_TO <> "C" And PREV_MATERIALS_CHARGE_TO <> "S" Then
            If NumericVal(txtMatUnitPrice.Text) > 0 Then
                txtMatUnitPrice.Text = ToDoubleNumber(NumericVal(txtMatUnitPrice) / 1.12)
            End If
        End If
    Else
        cboAcctCodeMaterials.ListIndex = -1
        cboAcctCodeMaterials.Enabled = False
        If PREV_MATERIALS_CHARGE_TO = "C" Or PREV_MATERIALS_CHARGE_TO = "S" Then
            If NumericVal(txtMatUnitPrice.Text) > 0 Then
                txtMatUnitPrice.Text = ToDoubleNumber(NumericVal(txtMatUnitPrice) * 1.12)
            End If
        End If
    End If
    PREV_MATERIALS_CHARGE_TO = cboMatChargeTo
End Sub

Private Sub cboMatChargeTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cboMatChargeTo.ListIndex = -1
End Sub

Private Sub cboPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub chkAllowManDist_Click()
    If chkAllowManDist.Value = 1 Then
        fraParticipation.Enabled = True
        txtLOAAmount.Enabled = False
    Else
        fraParticipation.Enabled = False
        txtLOAAmount.Enabled = True
    End If
End Sub

Private Sub chkParticipat_Click()
    If chkParticipat.Value = 1 Then
        Screen.MousePointer = 11
        RO_OR_ESTI_OR_PART = "PART"
        'Me.Enabled = False

        'frmCustomerSearchRO.Show
        Screen.MousePointer = 0
    Else
        txtParticipat.Text = ""
        txtParticipation.Text = ""
    End If
End Sub

Private Sub cmdAccCancel_Click()
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
    Picture1.Enabled = True: Picture5.Enabled = True: pic3.Enabled = True: Frame2.Enabled = True
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

    Call SendToBack
    cmdCancel.Value = True
End Sub

Private Sub cmdAccDelete_Click()
    If Module_Access(LOGID, "DELETE ACCESSORIES ENTRY", "SYSTEM") = False Then Exit Sub

    If MsgQuestionBox("Delete This Accessories, Are you Sure?", "Delete Accessories Entry") = True Then
        SQL_STATEMENT = "delete from CSMS_RO_Det where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, labid, "ACC", "ACC NO: " & cboAccCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Dim cnt                                        As Integer
        Dim rsRo_detDup                                As New ADODB.Recordset
        Set rsRo_detDup = gconDMIS.Execute("select id,LINE_NO from CSMS_RO_Det where rep_or = " & N2Str2Null(rsREPOR!REP_OR) & " and livil = '4' order by LINE_NO asc")
        If Not rsRo_detDup.EOF And Not rsRo_detDup.BOF Then
            cnt = 0
            rsRo_detDup.MoveFirst
            Do While Not rsRo_detDup.EOF
                cnt = cnt + 1
                gconDMIS.Execute "update CSMS_RO_Det set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsRo_detDup!ID
                rsRo_detDup.MoveNext
            Loop
        End If
        rsRo_detDup.Close: Set rsRo_detDup = Nothing
        
        Call FillAccessories
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        SQL_STATEMENT = "update CSMS_RepOr set" & _
            " Accessories = " & TOTACCAMT - TOTACCTAX & "," & _
            " A_amtvalue = " & TOTACCAMT & "," & _
            " A_disc = " & TOTACCDISCVAL & "," & _
            " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
            " A_taxval = " & TOTACCTAX & "," & _
            " A_discount = " & TOTACCDISC & "," & _
            " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
            " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
            " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
            " where REP_OR = '" & txtRep_Or.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call ShowDeletedMsg
        Call rsRefresh
        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    End If
    cmdAccCancel.Value = True
End Sub

Private Sub cmdAccSave_Click()
    
    On Error GoTo ErrorCode

    Dim ACCREP_OR                                       As String
    Dim ACCLEVEL                                        As String
    Dim ACCLINE_NO                                      As String
    Dim ACCDETCDE                                       As String
    Dim ACCDETDSC                                       As String
    Dim ACCDETUNT                                       As String
    Dim ACCDETVOL                                       As Double
    Dim ACCDETPRC                                       As Double
    Dim ACCDETAMT                                       As Double
    Dim ACCCODE                                         As String
    Dim ACCWCODE                                        As String
    Dim ACCTAXRATE                                      As Double
    Dim ACCDISCRATE                                     As Double
    Dim ACCTAXVAL                                       As Double
    Dim ACCDISVAL                                       As Double
    Dim ACCPOCODE                                       As String
    Dim ACCRep_Or2                                      As String
    Dim ACCDETAIL                                       As String
    Dim ACCDET_AMT                                      As Double
    Dim ACCDIS_VAL                                      As Double
    Dim ACCDISCOUNT_2                                   As Double
    
    
    'Update by:NVB 03122010
    'Desc: Additional Validation, User Cannot Save if Account code is missing
    '      Refferences to Accounting Module
    '------------------------------------------
    If cboAccChargeTo = "C" Then
        If cboAcctCodeAccessories = "" Or IsNull(cboAcctCodeAccessories) = True Then
            MessagePop RecSaveError, "Saving Error", "You cannot Continue Please select Account Code"
            cboAcctCodeAccessories.SetFocus
            Exit Sub
        End If
    End If
    '------------------------------------------

'Upadte by: IEBV 08052010 1005Am
'Description:   Additional validation, user cannot save if discount value is greater than the jobrate
    If optAccByAmt.Value = True Then
        If NumericVal(txtAccDiscountAmt.Text) > NumericVal(txtAccAmount.Text) Then
            MessagePop RecSaveError, "Saving Error", "Discount amount is greater than the total amount"
            On Error Resume Next
            txtAccDiscountAmt.SetFocus
            Exit Sub
        End If
    End If
'----------------------------------------------------------------------------------------------------

       
    If txtAccDiscount.Text <= 100 Then
        txtAccDiscount.Text = Format(txtAccDiscount.Text, DIGIT_FORMAT)
    Else
        MessagePop RecSaveError, "Saving Error", "Percentage Discount is greater then 100%"
        txtAccDiscount.SetFocus
        Exit Sub
    End If
    

    ACCDISVAL = 0: ACCTAXVAL = 0: ACCDETAMT = 0
    ACCDIS_VAL = 0: ACCDISCOUNT_2 = 0: ACCDISCRATE = 0

    ACCREP_OR = N2Str2Null(txtRep_Or.Text)
    ACCLEVEL = "'4'"
    ACCLINE_NO = N2Str2Null(Format(txtAccLineNo.Text, "00"))
    ACCDETCDE = N2Str2Null(cboAccCode.Text)
    ACCDETDSC = N2Str2Null(Mid(cboAccessories.Text, 1, 100))
    ACCDETUNT = "NULL"
    ACCDETVOL = NumericVal(txtAccQty.Text)
    ACCDETPRC = NumericVal(txtAccUnitPrice.Text)
    ACCDETAMT = NumericVal(txtAccAmount.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    ACCCODE = N2Str2Null(SetAcctCode(cboAcctCodeAccessories.Text))
    ACCWCODE = N2Str2Null(cboAccChargeTo.Text)
    ACCTAXRATE = (VAT_RATE / 100)
    If cboAccChargeTo.Text = "C" Or cboAccChargeTo.Text = "S" Then
        If cboAcctCodeAccessories.Text = "" Then
            MsgBox "Account Code should be selected.", vbInformation, "Select Account"
            Exit Sub
        End If
    End If

    If optAccByPerc.Value = True Then
        ACCDISCRATE = NumericVal(txtAccDiscount.Text) / 100
        ACCDISVAL = (NumericVal(txtAccAmount.Text) * ACCDISCRATE) - ((NumericVal(txtAccAmount.Text) * ACCDISCRATE) * ACCTAXRATE)
    Else
        ACCDISCRATE = NumericVal(txtAccDiscountAmt.Text) / NumericVal(txtAccAmount.Text)
        ACCDISVAL = NumericVal(txtAccDiscountAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    End If
    ACCPOCODE = N2Str2Null(txtMatPOCode.Text)
    ACCRep_Or2 = "NULL"
    ACCDETAIL = "NULL"
    ACCDET_AMT = NumericVal(txtAccAmount.Text)
    ACCDIS_VAL = ACCDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
    If optAccByPerc.Value = True Then
        ACCDISCOUNT_2 = ACCDET_AMT * ACCDISCRATE
    Else
        ACCDISCOUNT_2 = NumericVal(txtAccDiscountAmt.Text)
    End If
    
    'COMMENT BY  : MJP 10162009 1030 AM
    'DESCRIPTION : DOUBLE VAT
        'ACCTAXVAL = ((ACCDETAMT - ACCDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100)
    'COMMENT BY  : MJP 10162009 1030 AM
    'UPDATE BY   : MJP 10162009 1030 AM
        ACCTAXVAL = ((ACCDET_AMT - ACCDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100)
    'UPDATE BY   : MJP 10162009 1030 AM
    
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"

    'UPDATE BY   : MJP 04212010 0244 PM
    'DESCRIPTION : TO AVOID A NEGATIVE AMOUNT IN THE DISPLAY
'    If chkParticipat.Value = 1 Then
'        If CheckIfInsuranceIsAlreadySet(txtRep_Or, labDetId, CCur(ACCDETPRC), CCur(ACCDISCOUNT_2)) = True Then
'            MsgBox "You are trying to input/Change a Values where in insurance value is already set and may result a negative value. " & vbCrLf & _
'                "Try to set first the insurance amount to zero then input the Value", vbCritical, "Error"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'    End If
    'UPDATE BY   : MJP 04212010 0244 PM
    
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into CSMS_RO_Det " & _
            "(rep_or,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,USERCDE,SAVEDATE,SAVETIME)" & _
            " values (" & ACCREP_OR & ", " & ACCLEVEL & ", " & ACCLINE_NO & "," & _
            " " & ACCDETCDE & "," & ACCDETDSC & "," & _
            " " & ACCDETUNT & ", " & ACCDETVOL & "," & _
            " " & ACCDETPRC & ", " & ACCDETAMT & ", " & ACCCODE & _
            ", " & ACCWCODE & ", " & ACCTAXRATE * 100 & ", " & ACCDISCRATE * 100 & _
            ", " & ACCTAXVAL & ", " & ACCDISVAL & ", " & ACCPOCODE & _
            ", " & ACCRep_Or2 & ", " & ACCDETAIL & ", " & ACCDET_AMT & _
            ", " & ACCDIS_VAL & ", " & ACCDISCOUNT_2 & _
            ", " & Vusercode & _
            ", " & VLastUpdate & _
            ", " & VLastUpdateTime & ")"
    Else
        SQL_STATEMENT = "update CSMS_RO_Det set" & _
            " rep_or = " & ACCREP_OR & "," & _
            " livil = " & ACCLEVEL & "," & _
            " LINE_NO = " & ACCLINE_NO & "," & _
            " detcde = " & ACCDETCDE & "," & _
            " detdsc = " & ACCDETDSC & "," & _
            " detunt = " & ACCDETUNT & "," & _
            " detvol = " & ACCDETVOL & "," & _
            " detprc = " & ACCDETPRC & "," & _
            " detamt = " & ACCDETAMT & "," & _
            " code = " & ACCCODE & "," & _
            " wcode = " & ACCWCODE & "," & _
            " taxrate = " & ACCTAXRATE * 100 & "," & _
            " discrate = " & ACCDISCRATE * 100 & "," & _
            " taxval = " & ACCTAXVAL & "," & _
            " disval = " & ACCDISVAL & "," & _
            " pocode = " & ACCPOCODE & "," & _
            " rep_or2 = " & ACCRep_Or2 & "," & _
            " detail = " & ACCDETAIL & "," & _
            " det_amt = " & ACCDET_AMT & "," & _
            " dis_val = " & ACCDIS_VAL & "," & _
            " discount_2 = " & ACCDISCOUNT_2 & "," & _
            " USERCDE = " & Vusercode & ", SAVEDATE = " & VLastUpdate & ", SAVETIME = " & VLastUpdateTime & _
            " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("EE", "BILLED OUT", SQL_STATEMENT, labid, "ACC", "ACC NO: " & cboAccCode, "", labDetID.Caption)
        'NEW LOG AUDIT-----------------------------------------------------
    End If
    
    Call FillAccessories
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
        " Accessories = " & Round(TOTACCAMT - TOTACCTAX, 2) & "," & _
        " A_amtvalue = " & Round(TOTACCAMT, 2) & "," & _
        " A_disc = " & Round(TOTACCDISCVAL, 2) & "," & _
        " A_disc2 = " & Round(TOTACCDISC * (VAT_RATE / 100), 2) & "," & _
        " A_taxval = " & Round(TOTACCTAX, 2) & "," & _
        " A_discount = " & Round(TOTACCDISC, 2) & "," & _
        " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTACCTAX, 2) & "," & _
        " WA_amt = " & ACCWarTotal & "," & _
        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
        " where REP_OR = '" & txtRep_Or.Text & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call ShowSuccessFullyUpdated
    Call rsRefresh
    rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    cmdAccCancel.Value = True
    
    If AddorEdit = "ADD" Then Call AddAccessories
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0: ShowVBError
End Sub





Private Sub cmdCancelBill_Click()
    SendToBackBill
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
    Picture1.Enabled = True: Picture5.Enabled = True: pic3.Enabled = True: Frame2.Enabled = True
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†
End Sub

Private Sub cmdCancelDisk_Click()
    Call SendToBackDisc
    
    'UPDATE BY : MJP 05 14 2008
        Picture1.Enabled = True
        Frame2.Enabled = True
        pic3.Enabled = True
        Picture5.Enabled = True
    'UPDATE BY : MJP 05 14 2008
End Sub

Function CheckIfTheresItemIssued() As Boolean
    Dim rstmp                                               As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT COUNT(DETCDE) FROM CSMS_RO_DET WHERE REP_OR = " & N2Str2Null(txtRep_Or) & " AND LIVIL <> '1'")
    If rstmp.Fields(0).Value = 0 Then
        CheckIfTheresItemIssued = False
    Else
        CheckIfTheresItemIssued = True
    End If
    Set rstmp = Nothing
End Function

Private Sub cmdCust_Click()
    'UPDATE BY   : MJP09232009
    'DESCRIPTION : THIS IS TO HAVE A VALIDITY FOR THE CUSTOMER NAME IN PART MODULE AND SERVICE MODULE
        If CheckIfTheresItemIssued = True Then
            If MsgBox("Changing the customer name when theres already issued Item, might lead to wrong customer name in Repair Order Print out and Other Invoice Paper, Proceed", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        End If
    'UPDATE BY   : MJP09232009

    Call FRMx.PassVariable("BILLING SYSTEM")
    FRMx.Show 1
End Sub

Private Sub cmdDelete_Click()
    If Module_Access(LOGID, "RO ADVANCE OPTIONS", "SYSTEM") = False Then Exit Sub

    If Function_Access(LOGID, "Acess_DELETE", "BILLING SYSTEM") = False Then Exit Sub

    If CheckIfROStillExist(txtRep_Or) = False Then
        MessagePop InfoWarning, "Repair order Information", "Repair order Cannot be found, please refresh your Billing System", 1000
        Exit Sub
    End If
    
    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Set Internal RO, please refresh your Billing System", 1000
        Exit Sub
    End If
    
    If CheckIfTheresItemIssued = True Then
        MsgBox "You cannot Delete this the RO when theres already issued Item. Kindly inform parts depatment to Unpost first the Issuances.", vbInformation, "Info."
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete this Repair Order?", vbYesNo + vbQuestion, "Warning") = vbYes Then
        SQL_STATEMENT = "delete from CSMS_RepOr where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("X", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or & " - FORCE DELETE", "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        SQL_STATEMENT = "delete from CSMS_RO_Det where rep_or = '" & txtRep_Or.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or & " - FORCE DELETE", "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        gconDMIS.Execute "DELETE FROM CSMS_PMS_Job_Det where REP_OR = '" & txtRep_Or & "'"
        gconDMIS.Execute "Delete From CSMS_RepairOrder Where RO_NO = '" & txtRep_Or.Text & "'"

        'UPDATED BY : MJP 04 14 2008
        'DESCRIPTION : ONE CAUSE OF TECHNICIAN INFO. STAYS AT THE JOB CLOCK
        gconDMIS.Execute ("Update HRMS_EMPINFO SET ASSIGNEDRO = NULL, JSTATUS = 'A' WHERE ASSIGNEDRO = '" & txtRep_Or.Text & "'")
        gconDMIS.Execute ("Update CSMS_EMPINFO SET ASSIGNEDRO = NULL, JSTATUS = 'A' WHERE ASSIGNEDRO = '" & txtRep_Or.Text & "'")
        If COMPANY_CODE = "HGC" Then
            gconDMIS.Execute "Update CSMS_Baymonitoring set ro = null, bay_status = 'Available' where ro = '" & txtRep_Or.Text & "'"
        End If
        gconDMIS.Execute ("DELETE FROM CSMS_JOBCLOCK WHERE RO_NO = '" & txtRep_Or.Text & "'")
        'UPDATED BY : MJP 04 14 2008

        Call ShowDeletedMsg
        
        'Call rsRefresh
        cmdCancel.Value = True
        
        Call txtSearch_Change
        Call Command5_Click
    End If
End Sub

Private Sub cmdInternalRO_Click()
    If Module_Access(LOGID, "RO ADVANCE OPTIONS", "SYSTEM") = False Then Exit Sub

    Dim rsACCOUNT_CODE                                          As New ADODB.Recordset
    Set rsACCOUNT_CODE = gconDMIS.Execute("SELECT CODE FROM CSMS_RO_DET WHERE REP_OR = '" & txtRep_Or & "' AND WCODE IN('S','C') AND CODE IS NULL")
    If Not rsACCOUNT_CODE.EOF And Not rsACCOUNT_CODE.BOF Then
       MsgBox "There is an internal transaction which" & vbCrLf & "Acct Code is not yet been selected.", vbExclamation, "INFORMATION"
       Exit Sub
    End If
    Set rsACCOUNT_CODE = Nothing
     
     
    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Set Internal RO, please refresh your Billing System", 1000
        Exit Sub
    End If
    
    If CheckIfAllJobIsFinish = False Then
        MsgBox "Repair Order # " & txtRep_Or.Text & " Job(s) not Yet Finish" & vbCrLf & "Please Finish All Job Before Billing this RO.", vbInformation, "Billing System"
        Exit Sub
    End If
    
    If txtInvoiceNo.Text = "" Then
        If MsgBox("Set this Repair order to a INTERNAL Repair, Are you sure?", vbYesNo + vbQuestion, "Warning") = vbYes Then
            'SQL_STATEMENT = "update CSMS_RepOr set invoice = 'INT RO', dte_comp = dte_recd, dte_rel = '" & CDate(LOGDATE) & "' Where ID = " & labid.Caption
            SQL_STATEMENT = "update CSMS_RepOr set invoice = 'INT RO' " & _
                ", dte_comp = dte_recd Where ID = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("P", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "INV NO: INT RO", "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            gconDMIS.Execute "update CSMS_RepairOrder set status = 'Billed' Where RO_NO = '" & txtRep_Or.Text & "'"
            MessagePop InfoFriend, "Repair order Information Updated", "Repair order Sucessfully tag as INT RO!", 1000

            rsRefresh
            rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
            cmdCancel.Value = True
        End If
    End If
End Sub

Private Sub cmdNoCharge_Click()
    If Module_Access(LOGID, "RO ADVANCE OPTIONS", "SYSTEM") = False Then Exit Sub

    Dim rsACCOUNT_CODE                                              As New ADODB.Recordset
    Set rsACCOUNT_CODE = gconDMIS.Execute("SELECT CODE FROM CSMS_RO_DET WHERE REP_OR = '" & txtRep_Or & "' AND WCODE IN('S','C') AND CODE IS NULL")
    If Not rsACCOUNT_CODE.EOF And Not rsACCOUNT_CODE.BOF Then
       MsgBox "There is an internal transaction which" & vbCrLf & "Acct Code is not yet been selected.", vbExclamation, "INFORMATION"
       Exit Sub
    End If
    Set rsACCOUNT_CODE = Nothing
     
    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Set to no charge, please refresh your Billing System", 1000
        Exit Sub
    End If
     
    If CheckIfAllJobIsFinish = False Then
        MsgBox "Repair Order # " & txtRep_Or.Text & " Job(s) not Yet Finish" & vbCrLf & "Please Finish All Job Before Billing this RO.", vbInformation, "Billing System"
        Exit Sub
    End If
    
    
    If MsgBox("This Repair Order will be set to NO CHARGE, Are you sure?", vbYesNo + vbQuestion, "Warning") = vbYes Then
        'SQL_STATEMENT = "update CSMS_RepOr set invoice = 'NO CHG', dte_comp = '" & CDate(LOGDATE) & "', dte_rel = '" & CDate(LOGDATE) & "' Where ID = " & labid.Caption
        SQL_STATEMENT = "update CSMS_RepOr set invoice = 'NO CHG', dte_comp = '" & CDate(LOGDATE) & "' Where ID = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("P", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "INV NO: NO CHG", "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        
        gconDMIS.Execute ("UPDATE CSMS_REPAIRORDER SET JSTATUS = 'Y', STATUS = 'Billed' WHERE RO_NO = " & N2Str2Null(txtRep_Or) & "")
        
        Dim Update_RIV                                 As Integer
        For Update_RIV = 1 To RO_RIV_Tranno_Counter
            SQL_STATEMENT = "update Ord_Hd set Status = 'B' where trantype = 'RIV' and Tranno = '" & Format(RO_RIV_Tranno(Update_RIV), "000000") & "'"
            gconDMIS.Execute SQL_STATEMENT
        Next

        MessagePop InfoFriend, "Repair order Information Updated", "Repair order Sucessfully tag as NO CHARGE!", 1000

        Call rsRefresh
        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
        cmdCancel.Value = True
    End If
End Sub

Private Sub cmdOkBill_Click()
    On Error GoTo ErrorCode
    If IsDate(txtInvoiceDate) = False Then
        MsgBox "Invalid Invoice Date Format", vbCritical, "ERROR"
        'UPDATED BY: JUN--------------------
        'DATE UPDATED: 03-13-2009
        'DESCRIPTION: TCN: HGC - 12735
            txtInvoiceDate.Text = ""
            txtInvoiceDate.Text = Date
        'UPDATED BY: JUN--------------------
        txtInvoiceDate.SetFocus
        Exit Sub
    End If

    If txtInvoiceNumber.Text <> "" And IsDate(txtInvoiceDate.Text) = True And IsDate(txtDateReleased.Text) = True Then
        If Date > CDate(txtInvoiceDate.Text) Then
            If Module_Access(LOGID, "BACK DATE", "SYSTEM") = False Then
                txtInvoiceDate.SetFocus
                Exit Sub
            End If
        End If
        Dim rsReporDup                                 As New ADODB.Recordset
        Dim rsINVOICEDUp                               As New ADODB.Recordset

        Set rsReporDup = gconDMIS.Execute("select invoice from CSMS_RepOr where invoice = '" & Format(txtInvoiceNumber.Text, "000000") & "'")
        If Not rsReporDup.BOF And Not rsReporDup.EOF Then
            MsgSpeechBox "Invoice Number Already Exist!"
        Else
            If COMPANY_CODE = "HGC" Then
                Set rsINVOICEDUp = gconDMIS.Execute("Select * from CSMS_INVOICE WHERE INVOICENO = '" & Format(txtInvoiceNumber.Text, "000000") & "'")
                If Not rsINVOICEDUp.EOF And Not rsINVOICEDUp.BOF Then
                    If Trim(Null2String(rsINVOICEDUp!Status)) = "P" Then
                        MsgSpeechBox "Invoice Number Already Exist!"
                        Exit Sub
                    End If
                    If Trim(Null2String(rsINVOICEDUp!Status)) = "C" Then
                        MsgSpeechBox "Invoice Number Already Used and Was Cancelled!"
                        Exit Sub
                    End If
                End If
            End If
            
            Dim xINVOICE_TYPE As String
            If Option1.Value = True Then
                xINVOICE_TYPE = N2Str2Null("CSH")
            Else
                xINVOICE_TYPE = N2Str2Null("CHG")
            End If
            
            SQL_STATEMENT = "update CSMS_RepOr set PRIN_DTE = '" & Date & _
                "', Term = " & xINVOICE_TYPE & _
                ", Invoice = '" & txtInvoiceNumber.Text & _
                "', Dte_comp = '" & CDate(txtInvoiceDate.Text) & _
                "' Where ID = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
                Dim rTYPE                                  As String
                Call NEW_LogAudit("B", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "INV NO: " & txtInvoiceNumber, "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            If COMPANY_CODE = "HGC" Then
                gconDMIS.Execute "Insert into CSMS_INVOICE (INVOICENO, INVOICEDATE, STATUS) " & _
                    " Values ('" & txtInvoiceNumber.Text & _
                    "', '" & CDate(txtInvoiceDate.Text) & _
                    "', 'P')"
            End If

            gconDMIS.Execute "update CSMS_RepairOrder set Status = 'Billed' " & _
                " where RO_No = '" & txtRep_Or.Text & "'"
            
            'UPDATE BY   : MJP012709 0608PM
            'DESCRIPTION : TO UPDATE THE STATUS OF JOB WHEN RO IS BEEN INVOICE
                gconDMIS.Execute ("UPDATE CSMS_RO_DET SET DONE = 'Y', STATUS = 'Y' " & _
                    "WHERE LIVIL = '1' AND REP_OR = '" & txtRep_Or & "'")
            'UPDATE BY   : MJP012709 0608PM
            
            Dim vRONO                                  As String
            vRONO = Left(txtRep_Or, 1) & Right(txtRep_Or, 6)
            Dim Update_RIV                             As Integer
            '***************************************************************************************
            'updating code:     jaa - 11132008      -   to update the RO Number directly
            'For Update_RIV = 1 To RO_RIV_Tranno_Counter
            '    gconDMIS.Execute "update PMIS_Ord_Hd set Status = 'B' where (STATUS <> 'C' OR STATUS <> 'N') AND trantype = 'RIV' and Tranno = '" & Format(RO_RIV_Tranno(Update_RIV), "000000") & "'"
            'Next
            'gconDMIS.Execute "update PMIS_Ord_Hd set Status = 'B' where (STATUS <> 'C' OR STATUS <> 'N') AND trantype = 'RIV' and RONO = '" & txtRep_Or.Text & "'"
            '***************************************************************************************
            '===[ EAP:051109: change original code. tag billed only if parts is posted.. ]====
                 gconDMIS.Execute "update PMIS_Ord_Hd set Status = 'B' where (STATUS = 'P') AND trantype = 'RIV' and RONO = '" & txtRep_Or.Text & "'"
            '===[]===

            gconDMIS.Execute "update PMIS_Ord_Hd set In_Process = 'N' where (STATUS <> 'C' OR STATUS <> 'N') AND trantype = 'RIV' and rono = '" & vRONO & "'"
            gconDMIS.Execute "update PMIS_Ord_Hist set In_Process = 'N' where (STATUS <> 'C' OR STATUS <> 'N') AND trantype = 'RIV' and rono = '" & vRONO & "'"

            'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
            'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
                Picture1.Enabled = True: Picture5.Enabled = True: pic3.Enabled = True: Frame2.Enabled = True
            'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

            'Call rsRefresh
            DoEvents
            Call SendToBackBill
            DoEvents
            
            'COMMENT BY  : MJP 062909 1157AM
            'DESCRIPTION : TO RETURN TO SEARCH MENU
            'rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
            'Call StoreMemVars
            'COMMENT BY  : MJP 062909 1157AM
            
            DoEvents

            If COMPANY_CODE = "HGC" Then
                Call SaveReprintInformation("SI", MODULENAME, txtRep_Or, "NULL", LOGDATE, LOGNAME, False)
                If CANCEL_ANS = "NO" Then Exit Sub
            End If
            
            Call UpdateROAmount
            If DiscTotal > 0 And (WARTotal > 0 Or SALESTotal > 0 Or COMTotal > 0) Then
                PRINTROWCODE
            ElseIf DiscTotal > 0 And (WARTotal = 0 And SALESTotal = 0 And COMTotal = 0) Then
                PRINTRODISCOUNT
            ElseIf DiscTotal = 0 And (WARTotal > 0 Or SALESTotal > 0 Or COMTotal > 0) Then
                PRINTROWSC
            Else
                PRINTRO
            End If
            
            If COMPANY_CODE = "HPC" Then
                rptDET.WindowShowPrintSetupBtn = True
                rptDET.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDET.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptDET.WindowTitle = "Repair Order Details"
                
                PrintSQLReport rptDET, CSMS_REPORT_PATH & "RO_Details.rpt", "{repor.rep_or} = '" & txtRep_Or & "'", CSMS_REPORT_CONNECTION, 1
            End If
            
            'NEW LOG AUDIT ----------------------------------------------------------------
                Call NEW_LogAudit("V", "BILLING SYSTEM", "", labid, "", "RO NO: " & txtRep_Or, "", "")
            'NEW LOG AUDIT ----------------------------------------------------------------
            
            'UPDATE BY   : MJP062909 1201PM †----------------------------------------------------------
            'DESCRIPTION : TO RETURN TO SEARCH OPTION
                Call Command5_Click
            'UPDATE BY   : MJP062909 1201PM ----------------------------------------------------------†
        End If
        rsReporDup.Close
        Set rsReporDup = Nothing
    Else
        MsgSpeechBox "Invalid Invoice Number"
    End If

    Exit Sub

ErrorCode:

    ShowVBError
    Exit Sub
End Sub

Private Sub cmdOkDisc_Click()
    Screen.MousePointer = 11
    Dim varREP_OR                                      As String
    Dim varDISVAL                                      As Double
    varREP_OR = N2Str2Null(rsREPOR!REP_OR)
    varDISVAL = NumericVal(txtDiscAmt.Text) / 100

    If SSTab1.SelectedItem = 1 Then
        Dim JOBID                                       As Long
        Dim JOBREP_OR                                   As String
        Dim JOBLEVEL                                    As String
        Dim JOBLINE_NO                                  As String
        Dim JOBDETCDE                                   As String
        Dim JOBDETDSC                                   As String
        Dim JOBDETUNT                                   As String
        Dim JOBDETVOL                                   As Double
        Dim JOBDETPRC                                   As Double
        Dim JOBDETAMT                                   As Double
        Dim JOBCODE                                     As String
        Dim JOBWCODE                                    As String
        Dim JOBTAXRATE                                  As Double
        Dim JOBDISCRATE                                 As Double
        Dim JOBTAXVAL                                   As Double
        Dim JOBDISVAL                                   As Double
        Dim JOBPOCODE                                   As String
        Dim JOBRep_Or2                                  As String
        Dim JOBDETAIL                                   As String
        Dim JOBDET_AMT                                  As Double
        Dim JOBDIS_VAL                                  As Double
        Dim JOBDISCOUNT_2                               As Double
        Dim JOBREMARKS                                  As String

        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_Det where rep_or = " & varREP_OR & " and livil = '1' order by LINE_NO asc")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            rsRO_DET.MoveFirst
            Do While Not rsRO_DET.EOF
                JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
                JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

                If Null2String(rsRO_DET!wCode) = "" Then
                    JOBID = rsRO_DET!ID
                    JOBREP_OR = varREP_OR
                    JOBLEVEL = "'1'"
                    JOBLINE_NO = N2Str2Null(rsRO_DET!LINE_NO)
                    JOBDETCDE = N2Str2Null(rsRO_DET!DETCDE)
                    JOBDETDSC = N2Str2Null(rsRO_DET!DETDSC)
                    JOBDETUNT = N2Str2Null(rsRO_DET!detunt)
                    JOBDETVOL = N2Str2IntZero(rsRO_DET!detvol)
                    JOBDETPRC = N2Str2Zero(rsRO_DET!DetPrc)
                    JOBCODE = N2Str2Null(rsRO_DET!Code)
                    JOBWCODE = N2Str2Null(rsRO_DET!wCode)
                    JOBTAXRATE = (VAT_RATE / 100)
                    JOBDISCRATE = varDISVAL
                    JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
                    JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
                    JOBPOCODE = N2Str2Null(rsRO_DET!pocode)
                    JOBRep_Or2 = "NULL"
                    JOBDETAIL = N2Str2Null(rsRO_DET!Detail)
                    JOBDET_AMT = JOBDETPRC
                    JOBDIS_VAL = JOBDISVAL * JOBTAXRATE
                    JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
                    
                    'JOBTAXVAL = (JOBDETAMT - JOBDISCOUNT_2) * JOBTAXRATE
                    JOBTAXVAL = (JOBDET_AMT - JOBDISCOUNT_2) * JOBTAXRATE

                    SQL_STATEMENT = "update CSMS_RO_Det set" & _
                        " rep_or = " & JOBREP_OR & "," & _
                        " livil = " & JOBLEVEL & "," & _
                        " LINE_NO = " & JOBLINE_NO & "," & _
                        " detcde = " & JOBDETCDE & "," & _
                        " detdsc = " & JOBDETDSC & "," & _
                        " detunt = " & JOBDETUNT & "," & _
                        " detvol = " & JOBDETVOL & "," & _
                        " detprc = " & JOBDETPRC & "," & _
                        " detamt = " & JOBDETAMT & "," & _
                        " code = " & JOBCODE & "," & _
                        " wcode = " & JOBWCODE & "," & _
                        " taxrate = " & (JOBTAXRATE * 100) & "," & _
                        " discrate = " & (JOBDISCRATE * 100) & "," & _
                        " taxval = " & JOBTAXVAL & "," & _
                        " disval = " & JOBDISVAL & "," & _
                        " pocode = " & JOBPOCODE & "," & _
                        " rep_or2 = " & JOBRep_Or2 & "," & _
                        " detail = " & JOBDETAIL & "," & _
                        " det_amt = " & JOBDET_AMT & "," & _
                        " dis_val = " & JOBDIS_VAL & "," & _
                        " discount_2 = " & JOBDISCOUNT_2 & _
                        " where id = " & JOBID
                    gconDMIS.Execute SQL_STATEMENT

                    'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "JOBS", "JOB CODE: " & Null2String(JOBDETCDE), "", Null2String(JOBID))
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                rsRO_DET.MoveNext
            Loop

            Call FillJobs
            ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                " labor = " & TOTJOBAMT - TOTJOBTAX & "," & _
                " l_amtvalue = " & TOTJOBAMT & "," & _
                " l_disc = " & TOTJOBDISCVAL & "," & _
                " l_disc2 = " & TOTJOBDISC * (VAT_RATE / 100) & "," & _
                " l_taxval = " & TOTJOBTAX & "," & _
                " l_discount = " & TOTJOBDISC & "," & _
                " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                " wl_amt = " & 0 & "," & _
                " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                " where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "RO NO", "RO NO: " & txtRep_Or, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    ElseIf SSTab1.SelectedItem = 2 Then
        Dim PARTSID                                     As Long
        Dim PARTSREP_OR                                 As String
        Dim PARTSLEVEL                                  As String
        Dim PARTSLINE_NO                                As String
        Dim PARTSDETCDE                                 As String
        Dim PARTSDETDSC                                 As String
        Dim PARTSDETUNT                                 As String
        Dim PARTSDETVOL                                 As Double
        Dim PARTSDETPRC                                 As Double
        Dim PARTSDETAMT                                 As Double
        Dim PARTSCODE                                   As String
        Dim PARTSWCODE                                  As String
        Dim PARTSTAXRATE                                As Double
        Dim PARTSDISCRATE                               As Double
        Dim PARTSTAXVAL                                 As Double
        Dim PARTSDISVAL                                 As Double
        Dim PARTSPOCODE                                 As String
        Dim PARTSRep_Or2                                As String
        Dim PARTSDETAIL                                 As String
        Dim PARTSDET_AMT                                As Double
        Dim PARTSDIS_VAL                                As Double
        Dim PARTSDISCOUNT_2                             As Double
        Dim PARTSREMARKS                                As String
        Dim PARTS_TRANTYPE                              As String
        Dim PARTS_TRANNO                                As String
        Dim PARTS_TRAN_ITEMNO                           As String
        
        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("select * from CSMS_RO_Det where rep_or = " & varREP_OR & " and livil = '2' order by LINE_NO asc")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            rsRO_DET.MoveFirst
            Do While Not rsRO_DET.EOF
                PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                PARTS_TRANTYPE = Left(Null2String(rsRO_DET!REF_RIV_ADB), 3)
                PARTS_TRANNO = Mid(Null2String(rsRO_DET!REF_RIV_ADB), 4, 6)
                PARTS_TRAN_ITEMNO = Right(Null2String(rsRO_DET!REF_RIV_ADB), 3)

                If Null2String(rsRO_DET!wCode) = "" Then
                    PARTSID = rsRO_DET!ID
                    PARTSREP_OR = varREP_OR
                    PARTSLEVEL = "'2'"
                    PARTSLINE_NO = Format(N2Str2Null(rsRO_DET!LINE_NO), "00")
                    PARTSDETCDE = N2Str2Null(rsRO_DET!DETCDE)
                    PARTSDETDSC = N2Str2Null(rsRO_DET!DETDSC)
                    PARTSDETUNT = N2Str2Null(rsRO_DET!detunt)
                    PARTSDETVOL = N2Str2Zero(rsRO_DET!detvol)
                    PARTSDETPRC = N2Str2Zero(rsRO_DET!DetPrc)
                    PARTSDETAMT = N2Str2Zero(rsRO_DET!DETAMT)
                    PARTSCODE = N2Str2Null(rsRO_DET!Code)
                    PARTSWCODE = N2Str2Null(rsRO_DET!wCode)
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = varDISVAL
                    PARTSDISVAL = (PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE)
                    PARTSPOCODE = N2Str2Null(rsRO_DET!pocode)
                    PARTSRep_Or2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = N2Str2Zero(rsRO_DET!DET_AMT)
                    PARTSDIS_VAL = PARTSDISVAL * PARTSTAXRATE
                    PARTSDISCOUNT_2 = PARTSDET_AMT * PARTSDISCRATE
                    PARTSTAXVAL = (PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE

                    SQL_STATEMENT = "update CSMS_RO_Det set" & _
                        " rep_or = " & PARTSREP_OR & "," & _
                        " livil = " & PARTSLEVEL & "," & _
                        " LINE_NO = " & PARTSLINE_NO & "," & _
                        " detcde = " & PARTSDETCDE & "," & _
                        " detdsc = " & PARTSDETDSC & "," & _
                        " detunt = " & PARTSDETUNT & "," & _
                        " detvol = " & PARTSDETVOL & "," & _
                        " detprc = " & PARTSDETPRC & "," & _
                        " detamt = " & PARTSDETAMT & "," & _
                        " code = " & PARTSCODE & "," & _
                        " wcode = " & PARTSWCODE & "," & _
                        " taxrate = " & PARTSTAXRATE * 100 & "," & _
                        " discrate = " & PARTSDISCRATE * 100 & "," & _
                        " taxval = " & PARTSTAXVAL & "," & _
                        " disval = " & PARTSDISVAL & "," & _
                        " pocode = " & PARTSPOCODE & "," & _
                        " rep_or2 = " & PARTSRep_Or2 & "," & _
                        " detail = " & PARTSDETAIL & "," & _
                        " det_amt = " & PARTSDET_AMT & "," & _
                        " dis_val = " & PARTSDIS_VAL & "," & _
                        " discount_2 = " & PARTSDISCOUNT_2 & _
                        " where id = " & PARTSID
                    gconDMIS.Execute SQL_STATEMENT
                    
                    'NEW LOG AUDIT-----------------------------------------------------
                        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "PARTS", "PART NO: " & Null2String(PARTSDETCDE), "", Null2String(PARTSID))
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                rsRO_DET.MoveNext
            Loop
            
            Call FillParts
            ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                " p_amtvalue = " & TOTPARTSAMT & "," & _
                " p_disc = " & TOTPARTSDISCVAL & "," & _
                " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
                " p_taxval = " & TOTPARTSTAX & "," & _
                " p_discount = " & TOTPARTSDISC & "," & _
                " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                " wp_amt = " & 0 & "," & _
                " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                " where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    ElseIf SSTab1.SelectedItem = 3 Then
        Dim MATID                                       As Long
        Dim MATREP_OR                                   As String
        Dim MATLEVEL                                    As String
        Dim MATLINE_NO                                  As String
        Dim MATDETCDE                                   As String
        Dim MATDETDSC                                   As String
        Dim MATDETUNT                                   As String
        Dim MATDETVOL                                   As Double
        Dim MATDETPRC                                   As Double
        Dim MATDETAMT                                   As Double
        Dim MatCode                                     As String
        Dim MATWCODE                                    As String
        Dim MATTAXRATE                                  As Double
        Dim MATDISCRATE                                 As Double
        Dim MATTAXVAL                                   As Double
        Dim MATDISVAL                                   As Double
        Dim MATPOCODE                                   As String
        Dim MATRep_Or2                                  As String
        Dim MATDETAIL                                   As String
        Dim MATDET_AMT                                  As Double
        Dim MATDIS_VAL                                  As Double
        Dim MATDISCOUNT_2                               As Double

        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("select * from CSMS_RO_Det where rep_or = " & varREP_OR & " and livil = '3' order by LINE_NO asc")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            rsRO_DET.MoveFirst
            Do While Not rsRO_DET.EOF
                MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
                MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

                If Null2String(rsRO_DET!wCode) = "" Then
                    MATID = rsRO_DET!ID
                    MATREP_OR = varREP_OR
                    MATLEVEL = "'3'"
                    MATLINE_NO = Format(N2Str2Null(rsRO_DET!LINE_NO), "00")
                    MATDETCDE = N2Str2Null(rsRO_DET!DETCDE)
                    MATDETDSC = N2Str2Null(rsRO_DET!DETDSC)
                    MATDETUNT = N2Str2Null(rsRO_DET!detunt)
                    MATDETVOL = N2Str2Zero(rsRO_DET!detvol)
                    MATDETPRC = N2Str2Zero(rsRO_DET!DetPrc)
                    MATDETAMT = N2Str2Zero(rsRO_DET!DETAMT)
                    MatCode = N2Str2Null(rsRO_DET!Code)
                    MATWCODE = N2Str2Null(rsRO_DET!wCode)
                    MATTAXRATE = (VAT_RATE / 100)
                    MATDISCRATE = varDISVAL
                    MATDISVAL = (MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE)
                    MATPOCODE = N2Str2Null(rsRO_DET!pocode)
                    MATRep_Or2 = "NULL"
                    MATDETAIL = "NULL"
                    MATDET_AMT = N2Str2Zero(rsRO_DET!DET_AMT)
                    MATDIS_VAL = MATDISVAL * MATTAXRATE
                    MATDISCOUNT_2 = MATDET_AMT * MATDISCRATE
                    MATTAXVAL = (MATDETAMT - MATDISCOUNT_2) * MATTAXRATE

                    SQL_STATEMENT = "update CSMS_RO_Det set" & _
                        " rep_or = " & MATREP_OR & "," & _
                        " livil = " & MATLEVEL & "," & _
                        " LINE_NO = " & MATLINE_NO & "," & _
                        " detcde = " & MATDETCDE & "," & _
                        " detdsc = " & MATDETDSC & "," & _
                        " detunt = " & MATDETUNT & "," & _
                        " detvol = " & MATDETVOL & "," & _
                        " detprc = " & MATDETPRC & "," & _
                        " detamt = " & MATDETAMT & "," & _
                        " code = " & MatCode & "," & _
                        " wcode = " & MATWCODE & "," & _
                        " taxrate = " & MATTAXRATE * 100 & "," & _
                        " discrate = " & MATDISCRATE * 100 & "," & _
                        " taxval = " & MATTAXVAL & "," & _
                        " disval = " & MATDISVAL & "," & _
                        " pocode = " & MATPOCODE & "," & _
                        " rep_or2 = " & MATRep_Or2 & "," & _
                        " detail = " & MATDETAIL & "," & _
                        " det_amt = " & MATDET_AMT & "," & _
                        " dis_val = " & MATDIS_VAL & "," & _
                        " discount_2 = " & MATDISCOUNT_2 & _
                        " where id = " & MATID
                    gconDMIS.Execute SQL_STATEMENT
                    
                    'NEW LOG AUDIT-----------------------------------------------------
                        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "MAT", "MAT NO: " & Null2String(MATDETCDE), "", Null2String(MATID))
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                rsRO_DET.MoveNext
            Loop
            
            Call FillMaterials
            ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                " material = " & TOTMATAMT - TOTMATTAX & "," & _
                " m_amtvalue = " & TOTMATAMT & "," & _
                " m_disc = " & TOTMATDISCVAL & "," & _
                " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
                " m_taxval = " & TOTMATTAX & "," & _
                " m_discount = " & TOTMATDISC & "," & _
                " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                " wm_amt = " & 0 & "," & _
                " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                " where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    ElseIf SSTab1.SelectedItem = 4 Then
        Dim ACCESSORIESID                               As Long
        Dim ACCESSORIESREP_OR                           As String
        Dim ACCESSORIESLEVEL                            As String
        Dim ACCESSORIESLINE_NO                          As String
        Dim ACCESSORIESDETCDE                           As String
        Dim ACCESSORIESDETDSC                           As String
        Dim ACCESSORIESDETUNT                           As String
        Dim ACCESSORIESDETVOL                           As Double
        Dim ACCESSORIESDETPRC                           As Double
        Dim ACCESSORIESDETAMT                           As Double
        Dim ACCESSORIESCODE                             As String
        Dim ACCESSORIESWCODE                            As String
        Dim ACCESSORIESTAXRATE                          As Double
        Dim ACCESSORIESDISCRATE                         As Double
        Dim ACCESSORIESTAXVAL                           As Double
        Dim ACCESSORIESDISVAL                           As Double
        Dim ACCESSORIESPOCODE                           As String
        Dim ACCESSORIESRep_Or2                          As String
        Dim ACCESSORIESDETAIL                           As String
        Dim ACCESSORIESDET_AMT                          As Double
        Dim ACCESSORIESDIS_VAL                          As Double
        Dim ACCESSORIESDISCOUNT_2                       As Double
        Dim ACCESSORIESREMARKS                          As String
        Dim ACCESSORIES_TRANTYPE                        As String
        Dim ACCESSORIES_TRANNO                          As String
        Dim ACCESSORIES_TRAN_ITEMNO                     As String
        
        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("select * from CSMS_RO_Det where rep_or = " & varREP_OR & " and livil = '4' order by LINE_NO asc")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            rsRO_DET.MoveFirst
            Do While Not rsRO_DET.EOF
                ACCESSORIESDISVAL = 0: ACCESSORIESTAXVAL = 0: ACCESSORIESDETAMT = 0
                ACCESSORIESDIS_VAL = 0: ACCESSORIESDISCOUNT_2 = 0: ACCESSORIESDISCRATE = 0

                ACCESSORIES_TRANTYPE = Left(Null2String(rsRO_DET!REF_RIV_ADB), 3)
                ACCESSORIES_TRANNO = Mid(Null2String(rsRO_DET!REF_RIV_ADB), 4, 6)
                ACCESSORIES_TRAN_ITEMNO = Right(Null2String(rsRO_DET!REF_RIV_ADB), 3)

                If Null2String(rsRO_DET!wCode) = "" Then
                    ACCESSORIESID = rsRO_DET!ID
                    ACCESSORIESREP_OR = varREP_OR
                    ACCESSORIESLEVEL = "'4'"
                    ACCESSORIESLINE_NO = Format(N2Str2Null(rsRO_DET!LINE_NO), "00")
                    ACCESSORIESDETCDE = N2Str2Null(rsRO_DET!DETCDE)
                    ACCESSORIESDETDSC = N2Str2Null(rsRO_DET!DETDSC)
                    ACCESSORIESDETUNT = N2Str2Null(rsRO_DET!detunt)
                    ACCESSORIESDETVOL = N2Str2Zero(rsRO_DET!detvol)
                    ACCESSORIESDETPRC = N2Str2Zero(rsRO_DET!DetPrc)
                    ACCESSORIESDETAMT = N2Str2Zero(rsRO_DET!DETAMT)
                    ACCESSORIESCODE = N2Str2Null(rsRO_DET!Code)
                    ACCESSORIESWCODE = N2Str2Null(rsRO_DET!wCode)
                    ACCESSORIESTAXRATE = (VAT_RATE / 100)
                    ACCESSORIESDISCRATE = varDISVAL
                    ACCESSORIESDISVAL = (ACCESSORIESDETPRC * ACCESSORIESDISCRATE) - ((ACCESSORIESDETPRC * ACCESSORIESDISCRATE) * ACCESSORIESTAXRATE)
                    ACCESSORIESPOCODE = N2Str2Null(rsRO_DET!pocode)
                    ACCESSORIESRep_Or2 = "NULL"
                    ACCESSORIESDETAIL = "NULL"
                    ACCESSORIESDET_AMT = N2Str2Zero(rsRO_DET!DET_AMT)
                    ACCESSORIESDIS_VAL = ACCESSORIESDISVAL * ACCESSORIESTAXRATE
                    ACCESSORIESDISCOUNT_2 = ACCESSORIESDET_AMT * ACCESSORIESDISCRATE
                    ACCESSORIESTAXVAL = (ACCESSORIESDETAMT - ACCESSORIESDISCOUNT_2) * ACCESSORIESTAXRATE

                    SQL_STATEMENT = "update CSMS_RO_Det set" & _
                        " rep_or = " & ACCESSORIESREP_OR & "," & _
                        " livil = " & ACCESSORIESLEVEL & "," & _
                        " LINE_NO = " & ACCESSORIESLINE_NO & "," & _
                        " detcde = " & ACCESSORIESDETCDE & "," & _
                        " detdsc = " & ACCESSORIESDETDSC & "," & _
                        " detunt = " & ACCESSORIESDETUNT & "," & _
                        " detvol = " & ACCESSORIESDETVOL & "," & _
                        " detprc = " & ACCESSORIESDETPRC & "," & _
                        " detamt = " & ACCESSORIESDETAMT & "," & _
                        " code = " & ACCESSORIESCODE & "," & _
                        " wcode = " & ACCESSORIESWCODE & "," & _
                        " taxrate = " & ACCESSORIESTAXRATE * 100 & "," & _
                        " discrate = " & ACCESSORIESDISCRATE * 100 & "," & _
                        " taxval = " & ACCESSORIESTAXVAL & "," & _
                        " disval = " & ACCESSORIESDISVAL & "," & _
                        " pocode = " & ACCESSORIESPOCODE & "," & _
                        " rep_or2 = " & ACCESSORIESRep_Or2 & "," & _
                        " detail = " & ACCESSORIESDETAIL & "," & _
                        " det_amt = " & ACCESSORIESDET_AMT & "," & _
                        " dis_val = " & ACCESSORIESDIS_VAL & "," & _
                        " discount_2 = " & ACCESSORIESDISCOUNT_2 & _
                        " where id = " & ACCESSORIESID
                    gconDMIS.Execute SQL_STATEMENT
                    
                    'NEW LOG AUDIT--------------------------------------------------
                        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "A", "ACC CODE: " & Null2String(rsRO_DET!DETCDE), "", Null2String(rsRO_DET!ID))
                    'NEW LOG AUDIT--------------------------------------------------
                End If
                rsRO_DET.MoveNext
            Loop
            
            Call FillAccessories
            ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                " ACCESSORIES = " & TOTACCAMT - TOTACCTAX & "," & _
                " A_amtvalue = " & TOTACCAMT & "," & _
                " A_disc = " & TOTACCDISCVAL & "," & _
                " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
                " A_taxval = " & TOTACCTAX & "," & _
                " A_discount = " & TOTACCDISC & "," & _
                " amount = " & ROTotal - TOTJOBDISC - TOTACCDISC - TOTMATDISC - TOTACCDISC & "," & _
                " rovat = " & TOTJOBTAX + TOTACCTAX + TOTMATTAX + TOTACCTAX & "," & _
                " wA_amt = " & 0 & "," & _
                " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTACCDISC - TOTMATDISC - TOTACCDISC, 2) & _
                " where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    End If


    Screen.MousePointer = 0
    Call cmdCancelDisk_Click
    
    Call rsRefresh
    rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    Call StoreMemVars
End Sub

Private Sub cmdPart_Click()
    If chkParticipat.Value = 1 Then
        Call FRMx.PassVariable("BILLING INSURANCE")
        FRMx.Show 1
    End If
End Sub

Private Sub cmdPartClose_Click()
    'UPDATE BY : MJP 05 14 2008
        Picture1.Enabled = True
        Frame2.Enabled = True
        pic3.Enabled = True
        Picture5.Enabled = True
    'UPDATE BY : MJP 05 14 2008

    Picture6.ZOrder 1
    Picture6.Visible = False
End Sub

Private Sub cmdPartSave_Click()
    Call UpdateParticipation
    Call cmdPartClose_Click
    
    MessagePop InfoFriend, "Repair order Information", "Insurance amount successfully Set", 1000
End Sub

Function CheckIfAllJobPassedQC(mREPOR As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT QC_STATUS FROM CSMS_RO_DET WHERE LIVIL = '1' AND REP_OR = '" & mREPOR & "' AND (QC_STATUS = 'N' OR QC_STATUS IS NULL)")
    If (rstmp.BOF And rstmp.EOF) Then
        CheckIfAllJobPassedQC = True
    Else
        CheckIfAllJobPassedQC = False
    End If
    Set rstmp = Nothing
End Function

Sub DisableFrame(COND As Boolean)
    Picture1.Enabled = COND
    Picture5.Enabled = COND
    pic3.Enabled = COND
    Frame2.Enabled = COND
End Sub

Function SearchRoToBilled(RONO As String)
    Dim rstmp                   As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT ID FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(RONO) & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        labid.Caption = rstmp!ID
        
        Call rsRefresh
        Call StoreMemVars
    
        Call cmdSelect_Click
    Else
        Call ShowNoRecord
    End If
    Set rstmp = Nothing
End Function

Function CHECK_JOBCOST(XXX As String) As Boolean
    Dim rsCHECK_COST As ADODB.Recordset
        Set rsCHECK_COST = gconDMIS.Execute("Select DETCOST  from CSMS_RO_DET where rep_or = '" & XXX & "' and JOBTYPE = 'BP' and  (DETCOST = 0 or DETCOST is NULL)")
        If Not rsCHECK_COST.EOF And Not rsCHECK_COST.BOF Then
            CHECK_JOBCOST = True
        Else
            CHECK_JOBCOST = False
        End If
    Set rsCHECK_COST = Nothing
End Function

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "BILLING SYSTEM") = False Then Exit Sub
    
    'UPDATED BY: JUN------------------------------------------------------------------------------------------
    'DATE UPDATED: 06-01-09
    'DESCRIPTION: VALIDATE TRANSACTION FOR CHARGE TO 'C' OR 'S' IF THEY SELECTED THE ACCOUNT CODE DESCRIPTION
        Dim rsACCOUNT_CODE As ADODB.Recordset
           Set rsACCOUNT_CODE = gconDMIS.Execute("SELECT CODE FROM CSMS_RO_DET WHERE REP_OR = '" & txtRep_Or & "' AND WCODE IN('S','C') AND CODE IS NULL")
           If Not rsACCOUNT_CODE.EOF And Not rsACCOUNT_CODE.BOF Then
               MsgBox "There is an internal transaction which" & vbCrLf & "Acct Code is not yet been selected.", vbExclamation, "INFORMATION"
               Exit Sub
           End If
        Set rsACCOUNT_CODE = Nothing
    'UPDATED BY: JUN------------------------------------------------------------------------------------------
    
    'UPDATED BY: JUN------------------------------------------------------------------------------------------
    'DATE UPDATED: 04022009
    'DESCRIPTION: DO NOT ALLOW TO PRINT BP TRANSACION IF THE JOB COST IS ZERO FOR ACCOUNTING PURPOSES
        If COMPANY_CODE = "HPI" Then
            If CHECK_JOBCOST(RTrim(LTrim(txtRep_Or))) = True Then
                MsgBox "There is a Zero Job Cost" & vbCrLf & "Please check the Job Cost.", vbInformation, "INFORMATION"
                Exit Sub
            End If
        End If
    'UPDATED BY: JUN------------------------------------------------------------------------------------------
    
    If REPRINT <> "" Then
        If AllowReprint("BILLING SYSTEM") = False Then Exit Sub
    End If
    
    If MsgBox("Print this Repair Order", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    If txtInvoiceNo.Text = "" Then
        If MsgQuestionBox("Is this for Final Billing?", "Billing System") = True Then
            'COMMENTED - FML 05072008 --> remove restrictrion for HAI For Purpose of Billing
            Option1.Value = True
            Ichg = False
            If QC_MODULE_ON = "ON" Then
                If CheckIfAllJobPassedQC(txtRep_Or) = True Then
                    Call SendToFrontBill
                    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
                    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
                        Call DisableFrame(False)
                    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†
                Else
                    MsgBox "Repair Order # " & txtRep_Or.Text & " Job(s) not Yet Approved by QC" & vbCrLf & "Please Let the CQ Approved Job Before Billing this RO.", vbInformation, "Billing System"
                    Exit Sub
                End If
            Else
                If CheckIfAllJobIsFinish = True Then
                    Call SendToFrontBill
                    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
                    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
                        Call DisableFrame(False)
                    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†
                Else
                    MsgBox "Repair Order # " & txtRep_Or.Text & " Job(s) not Yet Finish" & vbCrLf & "Please Finish All Job Before Billing this RO.", vbInformation, "Billing System"
                    Exit Sub
                End If
            End If
        Else
            Call UpdateROAmount
            'If chkParticipat.Value = 1 Then Call UpdateParticipation
            If DiscTotal > 0 And (WARTotal > 0 Or SALESTotal > 0 Or COMTotal > 0) Then
                Call PRINTROWCODE
            ElseIf DiscTotal > 0 And (WARTotal = 0 And SALESTotal = 0 And COMTotal = 0) Then
                Call PRINTRODISCOUNT
            ElseIf DiscTotal = 0 And (WARTotal > 0 Or SALESTotal > 0 Or COMTotal > 0) Then
                Call PRINTROWSC
            Else
                Call PRINTRO
            End If
            
            If COMPANY_CODE = "HPC" Then
                rptHPC.WindowTitle = "Service Invoice Attachment"
                If NumericVal(rsREPOR!INSAMT) > 0 Then
                    PrintSQLReport rptHPC, CSMS_REPORT_PATH & "RO_DETAILS_W_INS.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
                Else
                    PrintSQLReport rptHPC, CSMS_REPORT_PATH & "RO_DETAILS.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
                End If
            End If
            
            'NEW LOG AUDIT-----------------------------------------------------
                Dim rTYPE                          As String
                Call NEW_LogAudit("V", "BILLING SYSTEM", "", labid, "R", "RO NO: " & txtRep_Or, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    Else
        If COMPANY_CODE = "HGC" Then
            Call SaveReprintInformation("SI", MODULENAME, txtRep_Or, "NULL", LOGDATE, LOGNAME, False)
            If CANCEL_ANS = "NO" Then Exit Sub
        End If

        If DiscTotal > 0 And (WARTotal > 0 Or SALESTotal > 0 Or COMTotal > 0) Then
            Call PRINTROWCODE
        ElseIf DiscTotal > 0 And (WARTotal = 0 And SALESTotal = 0 And COMTotal = 0) Then
            Call PRINTRODISCOUNT
        ElseIf DiscTotal = 0 And (WARTotal > 0 Or SALESTotal > 0 Or COMTotal > 0) Then
            Call PRINTROWSC
        Else
            Call PRINTRO
        End If

        If COMPANY_CODE = "HPC" Then
            rptHPC.WindowTitle = "Service Invoice Attachment"
            If NumericVal(rsREPOR!INSAMT) > 0 Then
                PrintSQLReport rptHPC, CSMS_REPORT_PATH & "RO_DETAILS_W_INS.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
            Else
                PrintSQLReport rptHPC, CSMS_REPORT_PATH & "RO_DETAILS.rpt", "{repor.rep_or} = '" & txtRep_Or.Text & "'", CSMS_REPORT_CONNECTION, 1
            End If
        End If
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "BILLING SYSTEM", "", labid, "R", "RO NO: " & txtRep_Or, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    End If
End Sub

Private Sub cmdRelease_Click()
    If IsDate(txtDte_comp.Text) = True Then
        gconDMIS.Execute "update CSMS_RepOr set dte_rel = '" & txtDte_comp.Text & "', status = 'R' where id = " & labid.Caption
        gconDMIS.Execute "update CSMS_RO_Det set status = 'R' where rep_or = '" & txtRep_Or.Text & "'"

        gconDMIS.Execute "update CSMS_RepairOrder set Status = 'Released' where RO_No = '" & txtRep_Or.Text & "'"
        MsgBox "Repair Order Successfully Released!", vbOKOnly + vbInformation, "Released..."
        Call rsRefresh
        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
        cmdCancel.Value = True
    End If
End Sub

Private Sub cmdReleaseRO_Click()
    If Function_Access(LOGID, "Acess_Post", "BILLING SYSTEM") = False Then Exit Sub
    
    If MsgBox("Are you Sure you want to release this RO?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    Call SendToFrontRel
    Call StoreMemVars
End Sub

Private Sub cmdCancelRel_Click()
    Call SendToBackRel
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A RELEASED DATE
        Picture1.Enabled = True: Picture5.Enabled = True: pic3.Enabled = True: Frame2.Enabled = True
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†
End Sub

Private Sub cmdOkRel_Click()
    If IsDate(txtReleaseDate.Text) = True Then
        If txtDte_comp.Text <> "" Then
            If CDate(txtReleaseDate.Text) < CDate(txtDte_comp.Text) Then
                If Module_Access(LOGID, "RELEASE DATE NOT LESS THAN TODAY", "SYSTEM") = False Then
                    On Error Resume Next
                    txtReleaseDate.SetFocus
                    Exit Sub
                End If
            End If
        End If
        SQL_STATEMENT = "update CSMS_RepOr set dte_rel = '" & txtReleaseDate.Text & "', status = 'R' where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("RD", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "INV NO: " & txtInvoiceNo, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
        'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A RELEASED DATE
            Picture1.Enabled = True: Picture5.Enabled = True: pic3.Enabled = True: Frame2.Enabled = True
        'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

        SQL_STATEMENT = "update CSMS_RO_Det set Done = 'Y', status = 'R' where rep_or = '" & txtRep_Or.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT
            Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "RELEASED: " & Date, "", "")
        'NEW LOG AUDIT
        
        SQL_STATEMENT = "update CSMS_RepairOrder set JSTATUS = 'R', Status = 'Released' where RO_No = '" & txtRep_Or.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT
            Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "RELEASED: " & Date, "", "")
        'NEW LOG AUDIT

        Call SendToBackRel
        
        'COMMENT BY  : MJP062909 0113PM †----------------------------------------------------------
        'DESCRIPTION : TO GO BACK TO SEARCH
            'Call rsRefresh
            'rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
        'COMMENT BY  : MJP062909 0113PM ----------------------------------------------------------†
        cmdCancel.Value = True
        MessagePop InfoFriend, "RO Information Updated", "Repair Order Succesfully Released", 1000
        
        Call Command5_Click
    Else
        MsgSpeechBox "Invalid Release Date"
        'UPDATED BY: JUN--------------------
        'DATE UPDATED: 03-13-2009
        'DESCRIPTION: TCN: HGC - 12735
         txtReleaseDate.Text = ""
         txtReleaseDate.Text = Date
        'UPDATED BY: JUN--------------------
        txtReleaseDate.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cboDescription_LostFocus()
    On Error Resume Next
    cboDescription.Text = SetPartDisc(cboPartNo.Text)
    txtUnitPrice.Text = SetPartPrice(cboPartNo.Text)
    txtPartAmount.Text = NumericVal(txtQty.Text) * NumericVal(txtUnitPrice.Text)
End Sub

Private Sub cboJobCode_Click()
    '    If optByDescription.Value = True Then
    '        If cboJobCode.Text <> "" Then cboJcode.Text = setJobCode(cboJobCode.Text)
    '        txtJobPostCode.Text = setJobPOcode(cboJobCode.Text)
    '        txtJobRate.Text = setJobRate(cboJobCode.Text)
    '        If AddorEdit = "ADD" Then txtJobDetail.Text = setJobDetail(cboJobCode.Text)
    '    End If
End Sub

Private Sub cboMaterial_LostFocus()
    If cboMaterial.Text <> "" Then
        cboMatCode.Text = SetMatCode(cboMaterial.Text)
        txtMatUnitPrice.Text = SetMatPrice(cboMatCode.Text)
        txtMatPOCode.Text = SetMatPOCode(cboMatCode.Text)
        txtMatAmount.Text = txtMatUnitPrice.Text
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "BILLING SYSTEM") = False Then Exit Sub
    AddorEdit = "ADD"
    labAddOrEdit = "ADD"
    RO_OR_ESTI_OR_PART = "RO"
    Call SendToBack
    frmAllCustomer.Show
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "BILLING SYSTEM") = False Then Exit Sub
    AddorEdit = "EDIT"
    labAddOrEdit = "EDIT"
    Frame1.Enabled = True
    Picture5.Enabled = False:   pic3.Enabled = False
    Picture1.Visible = False:   Picture2.Visible = True
    
    PrevRoNumber = txtRep_Or
    'Updated: IEBV 06282010 1158AM
    'Description: User can edit RO Type for the HCI
    If COMPANY_CODE = "HCI" Then
        Cbo_Rotype.Visible = True
        txttype.Visible = False
        If txttype.Text = "" Then
            Cbo_Rotype.ListIndex = 0
        End If
    End If
    'Updated: IEBV 06282010 1158AM
    'Description: User can edit RO Type for the HCI
End Sub

Private Sub cmdCancel_Click()
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
    Picture1.Enabled = True: Picture5.Enabled = True: pic3.Enabled = True: Frame2.Enabled = True
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

    Call SendToBack
    Frame1.Enabled = False
    labAddOrEdit = ""
    AddorEdit = ""
    Picture1.Visible = True
    Picture2.Visible = False
    Call StoreMemVars
    'Updated: IEBV 06100310pm
    'Description:
    If COMPANY_CODE = "HCI" Then
        txttype.Visible = True
        Cbo_Rotype.Visible = False
    End If
    'Updated: IEBV 06100310pm
    'Description:
End Sub

Private Sub cmdFind_Click()
    picMain.Visible = False
    picSEARCH.Visible = True
    picCustLimit.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdJobCancel_Click()
    Call SendToBack
    cmdCancel.Value = True
    'Call FillJobs
    'Call FillDetails
End Sub

Sub ClearOrStayTechnician(vEMPNO As String, vRONO As String, VTECHCODE As String)
    Dim rsHRMS                                         As New ADODB.Recordset
    Dim rsDet                                          As New ADODB.Recordset
    Dim rstmp                                          As New ADODB.Recordset
    Dim WORKING_CNT                                    As Integer
    Dim X                                              As Integer
    X = 0
    Set rsHRMS = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = '" & vEMPNO & "'")
    If Not (rsHRMS.BOF And rsHRMS.EOF) Then
        Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & Trim(VTECHCODE) & "' AND (DONE IS NULL OR DONE <> 'Y')")
        If Not (rsDet.BOF And rsDet.EOF) Then
            Do While Not rsDet.EOF
                X = X + 1
                rsDet.MoveNext
            Loop
            If X > 0 Then
                Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & Trim(VTECHCODE) & "' AND DONE = 'W'")
                If Not (rstmp.BOF And rstmp.EOF) Then

                Else
                    SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET JSTATUS = 'S',ASSIGNEDRO = '" & vRONO & "' WHERE EMPNO = '" & vEMPNO & "'"
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                Set rstmp = Nothing

                'SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET JSTATUS = 'A',ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'"
                'gconDMIS.Execute SQL_STATEMENT
                'NEW LOG AUDIT-----------------------------------------------------
                '   Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            Else
                'gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET JSTATUS = 'S' WHERE EMPNO = '" & VEMPNO & "'")
            End If
        Else
            SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET JSTATUS = 'A',ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'"
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    Else
        Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & VTECHCODE & "' AND (DONE IS NULL OR DONE <> 'Y')")
        If Not (rsDet.BOF And rsDet.EOF) Then
            Do While Not rsDet.EOF
                X = X + 1
                rsDet.MoveNext
            Loop
            If X > 0 Then
                Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & Trim(VTECHCODE) & "' AND DONE = 'W'")
                If Not (rstmp.BOF And rstmp.EOF) Then

                Else
                    SQL_STATEMENT = "UPDATE CSMS_EMPINFO SET JSTATUS = 'S',ASSIGNEDRO = '" & vRONO & "' WHERE EMPNO = '" & vEMPNO & "'"
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW LOG AUDIT-----------------------------------------------------
                        Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "CSMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                Set rstmp = Nothing

                'SQL_STATEMENT = "UPDATE CSMS_EMPINFO SET JSTATUS = 'A',ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'"
                'gconDMIS.Execute SQL_STATEMENT
                'NEW LOG AUDIT-----------------------------------------------------
                '    Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "CSMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            Else
                'gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET JSTATUS = 'S' WHERE EMPNO = '" & VEMPNO & "'")
            End If
        Else
            'UPDATED BY: JUN---------------------------------------------------------------------------------------------------------------------------------------
            'DATE UPDATED: 03-10-2009
            'DESCRIPTION: UPDATE FOR TICKET NO: HGC - 12739 and HMH - 12685
                SQL_STATEMENT = "UPDATE CSMS_EMPINFO SET JSTATUS = 'A',ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'"
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT-----------------------------------------------------
                     Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            'UPDATED BY: JUN---------------------------------------------------------------------------------------------------------------------------------------
        End If
    End If

    Set rsDet = Nothing
End Sub

Private Sub cmdJobDelete_Click()                      '
    Dim rstmp                                          As New ADODB.Recordset

    If MsgQuestionBox("Delete This Job, Are you Sure?", "Delete Job Entry") = True Then
        SQL_STATEMENT = "delete from CSMS_RO_Det where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, labid, "JOBS", "JOB CODE: " & cboJobCode, "", labDetID)
        'NEW LOG AUDIT-----------------------------------------------------

        'UPDATE BY   : MJP 10032008 0512 PM
        'DESCRIPTION : TO CHECK WHETHER THE TECHNICIAN STAY ASSIGN OR REMAIN
        'TICKET NO   : TCN 12452
        Dim rsKUTO                                     As New ADODB.Recordset
        Dim vEMPNO_X                                   As String
        Set rsKUTO = gconDMIS.Execute("SELECT EMPNO FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & LTrim(RTrim(lblTECHCODE_X)) & "'")
        If Not rsKUTO.EOF Or Not rsKUTO.BOF Then
            vEMPNO_X = LTrim(RTrim(Null2String(rsKUTO!EMPNO)))
        End If
        Call ClearOrStayTechnician(vEMPNO_X, txtRep_Or, lblTECHCODE_X)
        'UPDATE BY   : MJP 10032008 0512 PM

        gconDMIS.Execute "delete from CSMS_JobClock where ro_nO = '" & txtRep_Or.Text & "' and detcde = '" & cboJcode.Text & "'"
        gconDMIS.Execute "delete from CSMS_PMS_Job_det where REP_OR = '" & txtRep_Or.Text & "' AND PMS_MODEL = '" & cboJobCode & "'"
    Else
        Exit Sub
    End If

    Dim cnt                                            As Integer
    Dim rsRo_detDup                                    As New ADODB.Recordset
    Set rsRo_detDup = gconDMIS.Execute("select id,LINE_NO from CSMS_RO_Det where rep_or = " & N2Str2Null(rsREPOR!REP_OR) & " and livil = '1' order by LINE_NO asc")
    If Not rsRo_detDup.EOF And Not rsRo_detDup.BOF Then
        cnt = 0
        rsRo_detDup.MoveFirst
        Do While Not rsRo_detDup.EOF
            cnt = cnt + 1
            'gconDMIS.Execute "update CSMS_RO_Det set LINE_NO = " & N2Str2Null(cnt) & " where id = " & rsRo_detDup!ID
            rsRo_detDup.MoveNext
        Loop
    End If
    Set rsRo_detDup = Nothing
    
    Call FillJobs
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
        " labor = " & TOTJOBAMT - TOTJOBTAX & "," & _
        " l_amtvalue = " & TOTJOBAMT & "," & _
        " l_disc = " & TOTJOBDISCVAL & "," & _
        " l_disc2 = " & TOTJOBDISC * (VAT_RATE / 100) & "," & _
        " l_taxval = " & TOTJOBTAX & "," & _
        " l_discount = " & TOTJOBDISC & "," & _
        " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
        " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC, 2) & _
        " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call ShowDeletedMsg
    Call rsRefresh
    rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    cmdJobCancel.Value = True
End Sub

Private Sub cmdJobSave_Click()
    If cboJcode.Text = "" Then
        MsgSpeechBox "Cannot find Job Description... Please repeat choosing Job Description"
        On Error Resume Next
        cboJcode.SetFocus
        Exit Sub
    End If
    
    'Update by:NVB 03122010
    'Desc: Additional Validation, User Cannot Save if Account code is missing
    '      Refferences to Accounting Module
    '------------------------------------------
    If cboJobChargeTo = "C" Then
        If cboAcctCodeLabor = "" Or IsNull(cboAcctCodeLabor) = True Then
            MessagePop RecSaveError, "Saving Error", "You cannot Continue Please select Account Code"
            cboAcctCodeLabor.Enabled = True
            cboAcctCodeLabor.SetFocus
            Exit Sub
        End If
    End If
    '------------------------------------------
    
'Upadte by: IEBV 08052010 1005Am
'Description:   Additional validation, user cannot save if discount value is greater than the jobrate
  If optByAmt.Value = True Then
    If txtJobDiscountAmt.Text > txtJobRate.Text Then
        MessagePop RecSaveError, "Saving Error", "Discount amount is greater than the jobrate"
        On Error Resume Next
        txtJobDiscountAmt.SetFocus
        Exit Sub
    End If
  End If
'----------------------------------------------------------------------------------------------------

    
  If txtJobDiscount.Text <= 100 Then
        txtJobDiscount.Text = Format(txtJobDiscount.Text, DIGIT_FORMAT)
    Else
        MessagePop RecSaveError, "Saving Error", "Percentage Discount is greater then 100%"
        txtJobDiscount.SetFocus
        Exit Sub
    End If


    Dim JOBREP_OR                                       As String
    Dim JOBLEVEL                                        As String
    Dim JOBLINE_NO                                      As String
    Dim JOBDETCDE                                       As String
    Dim JOBDETDSC                                       As String
    Dim JOBDETUNT                                       As String
    Dim JOBDETVOL                                       As Double
    Dim JOBDETPRC                                       As Double
    Dim JOBDETAMT                                       As Double
    Dim JOBCODE                                         As String
    Dim JOBWCODE                                        As String
    Dim JOBTAXRATE                                      As Double
    Dim JOBDISCRATE                                     As Double
    Dim JOBTAXVAL                                       As Double
    Dim JOBDISVAL                                       As Double
    Dim JOBPOCODE                                       As String
    Dim JOBRep_Or2                                      As String
    Dim JOBDETAIL                                       As String
    Dim JOBDET_AMT                                      As Double
    Dim JOBDIS_VAL                                      As Double
    Dim JOBDISCOUNT_2                                   As Double
    Dim JOBREMARKS                                      As String
    Dim JOBTECHNICIAN                                   As String
    Dim JOBDET_HRS                                      As String
    Dim QUICK_SERVICE                                   As String

    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

    JOBREP_OR = N2Str2Null(txtRep_Or.Text)
    JOBLEVEL = "'1'"
    JOBLINE_NO = N2Str2Null(txtJobLineNo.Text)
    JOBDETCDE = N2Str2Null(cboJcode.Text)
    JOBDETDSC = N2Str2Null(cboJobCode.Text)
    JOBDETUNT = "NULL"
    JOBDETVOL = NumericVal(0)
    JOBDETPRC = NumericVal(txtJobRate.Text)
    JOBCODE = N2Str2Null(SetAcctCode(cboAcctCodeLabor.Text))
    JOBWCODE = N2Str2Null(cboJobChargeTo.Text)
    JOBTAXRATE = (VAT_RATE / 100)
    JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
    If cboJobChargeTo.Text = "C" Or cboJobChargeTo.Text = "S" Then
        If cboAcctCodeLabor.Text = "" Then
            MsgBox "Account Code should be selected.", vbInformation, "Select Account"
            Exit Sub
        End If
    End If
    If optByPerc.Value = True Then
        JOBDISCRATE = NumericVal(txtJobDiscount.Text) / 100
        JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
    Else
        JOBDISCRATE = NumericVal(txtJobDiscountAmt.Text) / JOBDETPRC
        JOBDISVAL = NumericVal(txtJobDiscountAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    End If
    JOBPOCODE = N2Str2Null(txtJobPostCode.Text)
    JOBRep_Or2 = "NULL"
    JOBDETAIL = N2Str2Null(CheckChar(Replace(txtJobDetail.Text, vbCrLf, " ")))
    JOBDET_AMT = JOBDETPRC
    JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
    If optByPerc.Value = True Then
        JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
    Else
        JOBDISCOUNT_2 = NumericVal(txtJobDiscountAmt.Text)
    End If
    JOBREMARKS = N2Str2Null(CheckChar(txtJobDetail.Text))
    JOBTECHNICIAN = N2Str2Null(setTechnicianCode(cboTechnician.Text))
    JOBDET_HRS = NumericVal(txtDET_HRS.Text)
    
    'COMMENT BY  : MJP 103010162009 AM
    'DESCRIPTION : DOUBLE VAT
        'JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    'COMMENT BY  : MJP 103010162009 AM
    
    'UPDATE BY   : MJP 103010162009 AM
        JOBTAXVAL = Round(((JOBDET_AMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    'UPDATE BY   : MJP 103010162009 AM
    
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    
    'UPDATE BY   : MJP 04212010 0244 PM
    'DESCRIPTION : TO AVOID A NEGATIVE AMOUNT IN THE DISPLAY
'    If chkParticipat.Value = 1 Then
'        If CheckIfInsuranceIsAlreadySet(txtRep_Or, labDetId, CCur(JOBDETPRC), CCur(JOBDISCOUNT_2)) = True Then
'            MsgBox "You are trying to input/Change a Values where in insurance value is already set and may result a negative value. " & vbCrLf & _
'                "Try to set first the insurance amount to zero then input the Value", vbCritical, "Error"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'    End If
    'UPDATE BY   : MJP 04212010 0244 PM
    
    If optQUICK.Visible = True Then
        If optQUICK.Value = 1 Then
            QUICK_SERVICE = N2Str2Null("Y")
        Else
            QUICK_SERVICE = N2Str2Null("N")
        End If
    Else
        QUICK_SERVICE = N2Str2Null("N")
    End If

    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_RO_Det " & _
            "(rep_or,livil,LINE_NO,detcde,detdsc,technician,HRSWRK,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,USERCDE,SAVEDATE,SAVETIME)" & _
            " values (" & JOBREP_OR & ", " & JOBLEVEL & ", " & JOBLINE_NO & "," & _
            " " & JOBDETCDE & "," & JOBDETDSC & "," & JOBTECHNICIAN & "," & JOBDET_HRS & "," & _
            " " & JOBDETUNT & ", " & JOBDETVOL & "," & _
            " " & JOBDETPRC & ", " & JOBDETAMT & ", " & JOBCODE & _
            ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & ", " & (JOBDISCRATE * 100) & _
            ", " & JOBTAXVAL & ", " & JOBDISVAL & ", " & JOBPOCODE & _
            ", " & JOBRep_Or2 & ", " & JOBDETAIL & ", " & JOBDET_AMT & _
            ", " & JOBDIS_VAL & ", " & JOBDISCOUNT_2 & _
            ", " & Vusercode & _
            ", " & VLastUpdate & _
            ", " & VLastUpdateTime & ")"
        gconDMIS.Execute SQL_STATEMENT
    Else
        SQL_STATEMENT = "update CSMS_RO_Det set" & _
            " rep_or = " & JOBREP_OR & "," & _
            " livil = " & JOBLEVEL & "," & _
            " LINE_NO = " & JOBLINE_NO & "," & _
            " detcde = " & JOBDETCDE & "," & _
            " detdsc = " & JOBDETDSC & "," & _
            " technician = " & JOBTECHNICIAN & ", HRSWRK = " & JOBDET_HRS & "," & _
            " detunt = " & JOBDETUNT & "," & _
            " detvol = " & JOBDETVOL & "," & _
            " detprc = " & JOBDETPRC & "," & _
            " detamt = " & JOBDETAMT & "," & _
            " code = " & JOBCODE & "," & _
            " wcode = " & JOBWCODE & "," & _
            " taxrate = " & (JOBTAXRATE * 100) & "," & _
            " discrate = " & (JOBDISCRATE * 100) & "," & _
            " taxval = " & JOBTAXVAL & "," & _
            " disval = " & JOBDISVAL & "," & _
            " pocode = " & JOBPOCODE & "," & _
            " rep_or2 = " & JOBRep_Or2 & "," & _
            " detail = " & JOBDETAIL & "," & _
            " det_amt = " & JOBDET_AMT & "," & _
            " dis_val = " & JOBDIS_VAL & "," & _
            " discount_2 = " & JOBDISCOUNT_2 & "," & _
            " USERCDE = " & Vusercode & ", SAVEDATE = " & VLastUpdate & ", SAVETIME = " & VLastUpdateTime & ", QUICK_SERVICE = " & QUICK_SERVICE & _
            " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "JOBS", "JOB CODE: " & Null2String(JOBDETCDE), "", labDetID.Caption)
        'NEW LOG AUDIT-----------------------------------------------------
    End If
    
    Call FillJobs
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
        " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) & "," & _
        " l_amtvalue = " & Round(TOTJOBAMT, 2) & "," & _
        " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
        " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
        " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
        " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
        " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & "," & _
        " wl_amt = " & JobWarTotal & "," & _
        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
        " where REP_OR = '" & txtRep_Or.Text & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call ShowSuccessFullyUpdated
    Call rsRefresh
    rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    Screen.MousePointer = 0
    If AddorEdit = "ADD" Then Call AddJobs Else cmdJobCancel.Value = True
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0: ShowVBError
End Sub

Private Sub cmdMatCancel_Click()
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
        Picture1.Enabled = True: Picture5.Enabled = True: pic3.Enabled = True: Frame2.Enabled = True
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

    SendToBack
    cmdCancel.Value = True
    txtdetail.Text = ""
    cmdAddMaterials.Visible = False
End Sub

Private Sub cmdMatDelete_Click()
    If Module_Access(LOGID, "DELETE MATERIALS ENTRY", "SYSTEM") = False Then Exit Sub

    If MsgQuestionBox("Delete This Materials, Are you Sure?", "Delete Materials Entry") = True Then
        gconDMIS.Execute "delete from CSMS_RO_Det where id = " & labDetID.Caption
        Dim cnt                                        As Integer
        Dim rsRo_detDup                                As ADODB.Recordset
        Set rsRo_detDup = New ADODB.Recordset
        Set rsRo_detDup = gconDMIS.Execute("select id,LINE_NO from CSMS_RO_Det where rep_or = " & N2Str2Null(rsREPOR!REP_OR) & " and livil = '3' order by LINE_NO asc")
        If Not rsRo_detDup.EOF And Not rsRo_detDup.BOF Then
            cnt = 0
            rsRo_detDup.MoveFirst
            Do While Not rsRo_detDup.EOF
                cnt = cnt + 1
                gconDMIS.Execute "update CSMS_RO_Det set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsRo_detDup!ID
                rsRo_detDup.MoveNext
            Loop
        End If
        rsRo_detDup.Close: Set rsRo_detDup = Nothing
        Call FillMaterials
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        SQL_STATEMENT = "update CSMS_RepOr set" & _
            " material = " & TOTMATAMT - TOTMATTAX & "," & _
            " m_amtvalue = " & TOTMATAMT & "," & _
            " m_disc = " & TOTMATDISCVAL & "," & _
            " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
            " m_taxval = " & TOTMATTAX & "," & _
            " m_discount = " & TOTMATDISC & "," & _
            " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
            " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
            " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
            " where REP_OR = '" & txtRep_Or.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "RO NO: " & txtRep_Or, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call ShowDeletedMsg
        Call rsRefresh
        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    End If
    cmdMatCancel.Value = True
    'cmdAddMaterials.Visible = False

End Sub

Private Sub cmdMatSave_Click()
    'Screen.MousePointer = 11
    On Error GoTo ErrorCode

    Dim MATREP_OR                                       As String
    Dim MATLEVEL                                        As String
    Dim MATLINE_NO                                      As String
    Dim MATDETCDE                                       As String
    Dim MATDETDSC                                       As String
    Dim MATDETUNT                                       As String
    Dim MATDET                                          As String
    Dim MATDETVOL                                       As Double
    Dim MATDETPRC                                       As Double
    Dim MATDETAMT                                       As Double
    Dim MatCode                                         As String
    Dim MATWCODE                                        As String
    Dim MATTAXRATE                                      As Double
    Dim MATDISCRATE                                     As Double
    Dim MATTAXVAL                                       As Double
    Dim MATDISVAL                                       As Double
    Dim MATPOCODE                                       As String
    Dim MATRep_Or2                                      As String
    Dim MATDETAIL                                       As String
    Dim MATDET_AMT                                      As Double
    Dim MATDIS_VAL                                      As Double
    Dim MATDISCOUNT_2                                   As Double
    
      
    'Update by:NVB 03122010
    'Desc: Additional Validation, User Cannot Save if Account code is missing
    '      Refferences to Accounting Module
    '------------------------------------------
    If cboMatChargeTo = "C" Then
        If cboAcctCodeMaterials = "" Or IsNull(cboAcctCodeMaterials) = True Then
            MessagePop RecSaveError, "Saving Error", "You cannot Continue Please select Account Code"
            cboAcctCodeMaterials.SetFocus
            Exit Sub
        End If
    End If
    '------------------------------------------
    
'Upadte by: IEBV 08052010 0152pm
'Description:   Additional validation, user cannot save if discount value is greater than the jobrate
  If optMatByAmt.Value = True Then
    If NumericVal(txtMatDiscountAmt.Text) > NumericVal(txtMatAmount.Text) Then
        MessagePop RecSaveError, "Saving Error", "Discount amount is greater than the total amount"
        On Error Resume Next
        txtMatDiscountAmt.SetFocus
        Exit Sub
    End If
  End If
'----------------------------------------------------------------------------------------------------

    
    
    If txtMatDiscount.Text <= 100 Then
        txtMatDiscount.Text = Format(txtMatDiscount.Text, DIGIT_FORMAT)
    Else
        MessagePop RecSaveError, "Saving Error", "Percentage Discount is greater then 100%"
        Exit Sub
    End If
    

    'Updated By     : IEBV 05262010 04:42AM
    'Description    : Additional Validation, User Cannot save if detail is missing and more than 150 characters
    If COMPANY_CODE = "HQA" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HCC" Then
        If Me.cboMatCode.Text = "" Then
        Else
        If UCase(Me.cboMatCode.Text) = "MISC" Or UCase(Me.cboMatCode.Text) = "MISC." Or UCase(Me.cboMatCode.Text) = "MISCELLANEOUS" Then
                If txtdetail.Text = "" Then
                    MessagePop RecSaveError, "Saving error", " You cannot leave Detail blank"
                    On Error Resume Next
                    txtdetail.SetFocus
                    Exit Sub
                Else
                    If Len(txtdetail) > 150 Then
                        MessagePop RecSaveError, "Saving error", "Please simplify your detail"
                        On Error Resume Next
                        txtdetail.SetFocus
                    End If
                End If
        Else
        End If
        End If
    End If    'Updated By     : IEBV 05262010 04:42AM
    'Description    : Additional Validation, User Cannot save if detail is missing and more than 150 characters

    
    MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
    MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

    MATREP_OR = N2Str2Null(txtRep_Or.Text)
    MATLEVEL = "'3'"
    MATLINE_NO = N2Str2Null(Format(txtMatLineNo.Text, "00"))
    MATDETCDE = N2Str2Null(cboMatCode.Text)
    MATDETDSC = N2Str2Null(Mid(cboMaterial.Text, 1, 100))
    MATDETUNT = "NULL"
    MATDETVOL = NumericVal(txtMatQty.Text)
    MATDETPRC = NumericVal(txtMatUnitPrice.Text)
    MATDETAMT = NumericVal(txtMatAmount.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    MatCode = N2Str2Null(SetAcctCode(cboAcctCodeMaterials.Text))
    MATWCODE = N2Str2Null(cboMatChargeTo.Text)
    MATTAXRATE = (VAT_RATE / 100)
    MATDET = N2Str2Null(UCase(txtdetail.Text))
    


    If cboMatChargeTo.Text = "C" Or cboMatChargeTo.Text = "S" Then
        If cboAcctCodeMaterials.Text = "" Then
            MsgBox "Account Code should be selected.", vbInformation, "Select Account"
            Exit Sub
        End If
    End If

    If optMatByPerc.Value = True Then
        MATDISCRATE = NumericVal(txtMatDiscount.Text) / 100
        MATDISVAL = (NumericVal(txtMatAmount.Text) * MATDISCRATE) - ((NumericVal(txtMatAmount.Text) * MATDISCRATE) * MATTAXRATE)
    Else
        MATDISCRATE = NumericVal(txtMatDiscountAmt.Text) / NumericVal(txtMatAmount.Text)
        MATDISVAL = NumericVal(txtMatDiscountAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    End If
    MATPOCODE = N2Str2Null(txtMatPOCode.Text)
    MATRep_Or2 = "NULL"
    MATDETAIL = "NULL"
    MATDET_AMT = NumericVal(txtMatAmount.Text)
    MATDIS_VAL = MATDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
    If optMatByPerc.Value = True Then
        MATDISCOUNT_2 = MATDET_AMT * MATDISCRATE
    Else
        MATDISCOUNT_2 = NumericVal(txtMatDiscountAmt.Text)
    End If
    
    'COMMENT BY  : MJP 103010162009 AM
    'DESCRIPTION : DOUBLE VAT
        'MATTAXVAL = ((MATDETAMT - MATDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100)
    'COMMENT BY  : MJP 103010162009 AM
    'UPDATE BY   : MJP 103010162009 AM
        MATTAXVAL = ((MATDET_AMT - MATDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100)
    'UPDATE BY   : MJP 103010162009 AM
    
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    
    'UPDATE BY   : MJP 04212010 0244 PM
    'DESCRIPTION : TO AVOID A NEGATIVE AMOUNT IN THE DISPLAY
'    If chkParticipat.Value = 1 Then
'        If CheckIfInsuranceIsAlreadySet(txtRep_Or, labDetId, CCur(MATDETPRC), CCur(MATDISCOUNT_2)) = True Then
'            MsgBox "You are trying to input/Change a Values where in insurance value is already set and may result a negative value. " & vbCrLf & _
'                "Try to set first the insurance amount to zero then input the Value", vbCritical, "Error"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'    End If
    'UPDATE BY   : MJP 04212010 0244 PM
    
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into CSMS_RO_Det " & _
            "(rep_or,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,USERCDE,SAVEDATE,SAVETIME)" & _
            " values (" & MATREP_OR & ", " & MATLEVEL & ", " & MATLINE_NO & "," & _
            " " & MATDETCDE & "," & MATDETDSC & "," & _
            " " & MATDETUNT & ", " & MATDETVOL & "," & _
            " " & MATDETPRC & ", " & MATDETAMT & ", " & MatCode & _
            ", " & MATWCODE & ", " & MATTAXRATE * 100 & ", " & MATDISCRATE * 100 & _
            ", " & MATTAXVAL & ", " & MATDISVAL & ", " & MATPOCODE & _
            ", " & MATRep_Or2 & ", " & MATDET & ", " & MATDET_AMT & _
            ", " & MATDIS_VAL & ", " & MATDISCOUNT_2 & _
            ", " & Vusercode & _
            ", " & VLastUpdate & _
            ", " & VLastUpdateTime & ")"
    Else

        SQL_STATEMENT = "update CSMS_RO_Det set" & _
            " rep_or = " & MATREP_OR & "," & _
            " livil = " & MATLEVEL & "," & _
            " LINE_NO = " & MATLINE_NO & "," & _
            " detcde = " & MATDETCDE & "," & _
            " detdsc = " & MATDETDSC & "," & _
            " detunt = " & MATDETUNT & "," & _
            " detvol = " & MATDETVOL & "," & _
            " detprc = " & MATDETPRC & "," & _
            " detamt = " & MATDETAMT & "," & _
            " code = " & MatCode & "," & _
            " wcode = " & MATWCODE & "," & _
            " taxrate = " & MATTAXRATE * 100 & "," & _
            " discrate = " & MATDISCRATE * 100 & "," & _
            " taxval = " & MATTAXVAL & "," & _
            " disval = " & MATDISVAL & "," & _
            " pocode = " & MATPOCODE & "," & _
            " rep_or2 = " & MATRep_Or2 & "," & _
            " detail = " & MATDET & "," & _
            " det_amt = " & MATDET_AMT & "," & _
            " dis_val = " & MATDIS_VAL & "," & _
            " discount_2 = " & MATDISCOUNT_2 & "," & _
            " USERCDE = " & Vusercode & ", SAVEDATE = " & VLastUpdate & ", SAVETIME = " & VLastUpdateTime & _
            " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, labid, "MAT", "MAT CODE: " & cboMatCode, "", labDetID.Caption)
        'NEW LOG AUDIT-----------------------------------------------------
    End If
    
    Call FillMaterials
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
        " material = " & Round(TOTMATAMT - TOTMATTAX, 2) & "," & _
        " m_amtvalue = " & Round(TOTMATAMT, 2) & "," & _
        " m_disc = " & Round(TOTMATDISCVAL, 2) & "," & _
        " m_disc2 = " & Round(TOTMATDISC * (VAT_RATE / 100), 2) & "," & _
        " m_taxval = " & Round(TOTMATTAX, 2) & "," & _
        " m_discount = " & Round(TOTMATDISC, 2) & "," & _
        " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & "," & _
        " wm_amt = " & MatWarTotal & "," & _
        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
        " where REP_OR = '" & txtRep_Or.Text & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call ShowSuccessFullyUpdated
    Call rsRefresh
    rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    cmdMatCancel.Value = True
    'Screen.MousePointer = 0
    If AddorEdit = "ADD" Then AddMaterials
    Exit Sub
    cmdAddMaterials.Visible = False

ErrorCode:
    Screen.MousePointer = 0: ShowVBError
End Sub

Private Sub cmdPartsCancel_Click()
    Call SendToBack
    cmdCancel.Value = True
End Sub

Private Sub cmdPartsDelete_Click()
    If Module_Access(LOGID, "DELETE PARTS ENTRY", "SYSTEM") = False Then Exit Sub

    If MsgQuestionBox("Delete This Parts, Are you Sure?", "Delete Parts Entry") = True Then
        SQL_STATEMENT = "delete from CSMS_RO_Det where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, labid, "PARTS", "PART NO: " & cboPartNo, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        Dim cnt                                        As Integer
        Dim rsRo_detDup                                As New ADODB.Recordset
        Set rsRo_detDup = gconDMIS.Execute("select id,LINE_NO from CSMS_RO_Det where rep_or = " & N2Str2Null(rsREPOR!REP_OR) & " and livil = '2' order by LINE_NO asc")
        If Not rsRo_detDup.EOF And Not rsRo_detDup.BOF Then
            cnt = 0
            rsRo_detDup.MoveFirst
            Do While Not rsRo_detDup.EOF
                cnt = cnt + 1
                gconDMIS.Execute "update CSMS_RO_Det set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsRo_detDup!ID
                rsRo_detDup.MoveNext
            Loop
        End If
        Set rsRo_detDup = Nothing
        
        Call FillParts
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        SQL_STATEMENT = "update CSMS_RepOr set" & _
                      " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                      " p_amtvalue = " & TOTPARTSAMT & "," & _
                      " p_disc = " & TOTPARTSDISCVAL & "," & _
                      " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
                      " p_taxval = " & TOTPARTSTAX & "," & _
                      " p_discount = " & TOTPARTSDISC & "," & _
                      " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                      " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                      " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                      " where REP_OR = '" & txtRep_Or.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call ShowDeletedMsg
        Call rsRefresh
        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    End If
    cmdPartsCancel.Value = True
End Sub

Private Sub cmdPartsSave_Click()
    'Screen.MousePointer = 11
    On Error GoTo ErrorCode


    Dim PARTSREP_OR                                     As String
    Dim PARTSLEVEL                                      As String
    Dim PARTSLINE_NO                                    As String
    Dim PARTSDETCDE                                     As String
    Dim PARTSDETDSC                                     As String
    Dim PARTSDETUNT                                     As String
    Dim PARTSDETVOL                                     As Double
    Dim PARTSDETPRC                                     As Double
    Dim PARTSDETAMT                                     As Double
    Dim PARTSCODE                                       As String
    Dim PARTSWCODE                                      As String
    Dim PARTSTAXRATE                                    As Double
    Dim PARTSDISCRATE                                   As Double
    Dim PARTSTAXVAL                                     As Double
    Dim PARTSDISVAL                                     As Double
    Dim PARTSPOCODE                                     As String
    Dim PARTSRep_Or2                                    As String
    Dim PARTSDETAIL                                     As String
    Dim PARTSDET_AMT                                    As Double
    Dim PARTSDIS_VAL                                    As Double
    Dim PARTSDISCOUNT_2                                 As Double
    Dim PARTSREMARKS                                    As String
    
    'Update by:NVB 03122010
    'Desc: Additional Validation, User Cannot Save if Account code is missing
    '      Refferences to Accounting Module
    '------------------------------------------
    If cboChargeTo = "C" Then
        If cboAcctCodeParts = "" Or IsNull(cboAcctCodeParts) = True Then
            MessagePop RecSaveError, "Saving Error", "You cannot Continue Please Select Account Code"
            cboAcctCodeParts.Enabled = True
            cboAcctCodeParts.SetFocus
            Exit Sub
        End If
    End If
    '------------------------------------------
    
'Upadte by: IEBV 08052010 0230pm
'Description:   Additional validation, user cannot save if discount value is greater than the jobrate
  If optPartsbyAmt.Value = True Then
    If NumericVal(txtPartDiscountAmt.Text) > NumericVal(txtPartAmount.Text) Then
        MessagePop RecSaveError, "Saving Error", "Discount amount is greater than the total amount"
        On Error Resume Next
        txtPartDiscountAmt.SetFocus
        Exit Sub
    End If
  End If
'----------------------------------------------------------------------------------------------------
    
    If txtPartDiscount.Text <= 100 Then
        txtPartDiscount.Text = Format(txtPartDiscount.Text, DIGIT_FORMAT)
    Else
        MessagePop RecSaveError, "Saving Error", "Percentage Discount is greater then 100%"
        txtPartDiscount.SetFocus
        Exit Sub
    End If
    
    
    

    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

    PARTSREP_OR = N2Str2Null(txtRep_Or.Text)
    PARTSLEVEL = "'2'"
    PARTSLINE_NO = N2Str2Null(Format(txtPartsLineNo.Text, "00"))
    PARTSDETCDE = N2Str2Null(cboPartNo.Text)
    PARTSDETDSC = N2Str2Null(Mid(cboDescription.Text, 1, 100))
    PARTSDETUNT = "NULL"
    PARTSDETVOL = NumericVal(txtQty.Text)
    PARTSDETPRC = NumericVal(txtUnitPrice.Text)
    PARTSDETAMT = Round(NumericVal(txtPartAmount.Text) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
    PARTSCODE = N2Str2Null(SetAcctCode(cboAcctCodeParts.Text))
    PARTSWCODE = N2Str2Null(cboChargeTo.Text)
    PARTSTAXRATE = (VAT_RATE / 100)
    If cboChargeTo.Text = "C" Or cboChargeTo.Text = "S" Then
        If cboAcctCodeParts.Text = "" Then
            MsgBox "Account Code should be selected.", vbInformation, "Select Account"
            Exit Sub
        End If
    End If
    If optPartsByPerc.Value = True Then
        PARTSDISCRATE = NumericVal(txtPartDiscount.Text) / 100
        PARTSDISVAL = (NumericVal(txtPartAmount.Text) * PARTSDISCRATE) - ((NumericVal(txtPartAmount.Text) * PARTSDISCRATE) * PARTSTAXRATE)
    Else
        PARTSDISCRATE = NumericVal(txtPartDiscountAmt.Text) / NumericVal(txtPartAmount.Text)
        PARTSDISVAL = NumericVal(txtPartDiscountAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE)
    End If

    PARTSPOCODE = N2Str2Null(txtPartCode.Text)
    PARTSRep_Or2 = "NULL"
    PARTSDETAIL = "NULL"
    PARTSDET_AMT = NumericVal(txtPartAmount.Text)
    PARTSDIS_VAL = Round(PARTSDISVAL * ConvertToBIRDecimalFormat(VAT_RATE), 2)

    If optPartsByPerc.Value = True Then
        PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
    Else
        PARTSDISCOUNT_2 = NumericVal(txtPartDiscountAmt.Text)
    End If
    
    'COMMENT BY  : MJP 103010162009 AM
    'DESCRIPTION : DOUBLE VAT
        'PARTSTAXVAL = Round(((PARTSDETAMT - PARTSDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    'COMMENT BY  : MJP 103010162009 AM
    'UPDATE BY   : MJP 103010162009 AM
        PARTSTAXVAL = Round(((PARTSDET_AMT - PARTSDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    'UPDATE BY   : MJP 103010162009 AM
    
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"

    'UPDATE BY   : MJP 04212010 0244 PM
    'DESCRIPTION : TO AVOID A NEGATIVE AMOUNT IN THE DISPLAY
'    If chkParticipat.Value = 1 Then
'        If CheckIfInsuranceIsAlreadySet(txtRep_Or, labDetId, CCur(PARTSDETPRC), CCur(PARTSDISCOUNT_2)) = True Then
'            MsgBox "You are trying to input/Change a Values where in insurance value is already set and may result a negative value. " & vbCrLf & _
'                "Try to set first the insurance amount to zero then input the Value", vbCritical, "Error"
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'    End If
    'UPDATE BY   : MJP 04212010 0244 PM
    
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into CSMS_RO_Det " & _
            "(rep_or,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,USERCDE,SAVEDATE,SAVETIME)" & _
            " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
            " " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
            " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
            " " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
            ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
            ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
            ", " & PARTSRep_Or2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
            ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & _
            ", " & Vusercode & _
            ", " & VLastUpdate & _
            ", " & VLastUpdateTime & ")"
    Else
        SQL_STATEMENT = "update CSMS_RO_Det set" & _
            " rep_or = " & PARTSREP_OR & "," & _
            " livil = " & PARTSLEVEL & "," & _
            " LINE_NO = " & PARTSLINE_NO & "," & _
            " detcde = " & PARTSDETCDE & "," & _
            " detdsc = " & PARTSDETDSC & "," & _
            " detunt = " & PARTSDETUNT & "," & _
            " detvol = " & PARTSDETVOL & "," & _
            " detprc = " & PARTSDETPRC & "," & _
            " detamt = " & PARTSDETAMT & "," & _
            " code = " & PARTSCODE & "," & _
            " wcode = " & PARTSWCODE & "," & _
            " taxrate = " & PARTSTAXRATE * 100 & "," & _
            " discrate = " & PARTSDISCRATE * 100 & "," & _
            " taxval = " & PARTSTAXVAL & "," & _
            " disval = " & PARTSDISVAL & "," & _
            " pocode = " & PARTSPOCODE & "," & _
            " rep_or2 = " & PARTSRep_Or2 & "," & _
            " detail = " & PARTSDETAIL & "," & _
            " det_amt = " & PARTSDET_AMT & "," & _
            " dis_val = " & PARTSDIS_VAL & "," & _
            " discount_2 = " & PARTSDISCOUNT_2 & "," & _
            " USERCDE = " & Vusercode & ", SAVEDATE = " & VLastUpdate & ", SAVETIME = " & VLastUpdateTime & _
            " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("EE", "BILLED OUT", SQL_STATEMENT, labid, "PARTS", "PART NO: " & cboPartNo, "", labDetID.Caption)
        'NEW LOG AUDIT-----------------------------------------------------
    End If
    
    Call FillParts
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
        " parts = " & Round(TOTPARTSAMT - TOTPARTSTAX, 2) & "," & _
        " p_amtvalue = " & Round(TOTPARTSAMT, 2) & "," & _
        " p_disc = " & Round(TOTPARTSDISCVAL, 2) & "," & _
        " p_disc2 = " & Round(TOTPARTSDISC * (VAT_RATE / 100), 2) & "," & _
        " p_taxval = " & Round(TOTPARTSTAX, 2) & "," & _
        " p_discount = " & Round(TOTPARTSDISC, 2) & "," & _
        " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & "," & _
        " wp_amt = " & PartsWarTotal & "," & _
        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
        " where REP_OR = '" & txtRep_Or.Text & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLED OUT", SQL_STATEMENT, labid, "R", "RO NO : " & txtRep_Or, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call ShowSuccessFullyUpdated
    Call rsRefresh
    rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
    cmdPartsCancel.Value = True
    'Screen.MousePointer = 0
    If AddorEdit = "ADD" Then AddParts
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0: ShowVBError
End Sub

Private Sub cmdROVatExempt_Click()
    If Module_Access(LOGID, "RO ADVANCE OPTIONS", "SYSTEM") = False Then Exit Sub

    'If txtInvoiceNo.Text = "" Then
    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = False Then
        If Trim(UCase(cmdROVatExempt.Caption)) = "SET TO NOT ZERO RATED" Then
            If MsgBox("This is a Not Zero Rated Customer, Are you sure?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
                Screen.MousePointer = 11
                SQL_STATEMENT = "update CSMS_RepOr set VAT_EXEMPT = 0 Where ID = " & labid.Caption
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("P", "BILLING SYSTEM", labid, "", "", "INV NO: NOT ZERO RATED", "", "")
                'NEW LOG AUDIT-----------------------------------------------------

                Call SetROTransToNonZeroRatedVat(txtRep_Or.Text)
                Screen.MousePointer = 0
                MessagePop InfoFriend, "Repair order Information Updated", "Repair order Sucessfully tag as NOT ZERO RATED!", 1000

                Call rsRefresh
                rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
                cmdCancel.Value = True
            End If
        Else
            If MsgBox("This is a Zero Rated Customer, Are you sure?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
                Screen.MousePointer = 11
                SQL_STATEMENT = "update CSMS_RepOr set L_TAXVAL = 0, P_TAXVAL = 0, M_TAXVAL = 0, A_TAXVAL = 0, VAT_EXEMPT = 1 Where ID = " & labid.Caption
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("P", "BILLING SYSTEM", labid, "", "", "INV NO: ZERO RATED", "", "")
                'NEW LOG AUDIT-----------------------------------------------------

                Call SetROTransToZeroRatedVat(txtRep_Or.Text)
                Screen.MousePointer = 0

                MessagePop InfoFriend, "Repair order Information Updated", "Repair order Sucessfully tag as ZERO RATED!", 1000
                Call rsRefresh
                rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
                cmdCancel.Value = True
            End If
        End If
    Else
        MsgBox "Repair order already invoiced.", vbInformation, "Info"
    End If
End Sub

Private Sub cmdSave_Click()
    If txtNiym.Text = "" Then
        MsgSpeechBox "Customer must have a name"
        On Error Resume Next
        txtNiym.SetFocus
        Exit Sub
    End If
    If cboRecd_by.Text = "" Then
        MsgSpeechBox "Service Advisor must not be Empty!"
        On Error Resume Next
        cboRecd_by.SetFocus
        Exit Sub
    Else
        Set rsEmpNo = New ADODB.Recordset
        Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo where naym = '" & cboRecd_by.Text & "'")
        If rsEmpNo.EOF And rsEmpNo.BOF Then
            MsgSpeechBox "Invalid Service Advisor"
            On Error Resume Next
            cboRecd_by.SetFocus
            Exit Sub
        End If
        Set rsEmpNo = Nothing
    End If
    If txtPlate_No.Text = "" Then
        ShowIsRequiredMsg "Plate no cannot be blank"
        Exit Sub
    End If
    
    Dim rsDupRepor                                     As New ADODB.Recordset
    If PrevRoNumber <> "" And PrevRoNumber <> LTrim(RTrim(txtRep_Or)) Then
        Set rsDupRepor = New ADODB.Recordset
        Set rsDupRepor = gconDMIS.Execute("select rep_or from CSMS_RepOr where rep_or = " & N2Str2Null(txtRep_Or.Text))
        If Not rsDupRepor.EOF And Not rsDupRepor.BOF Then
            MsgSpeechBox "Repair Order Number Already Exist!"
            On Error Resume Next
            txtRep_Or.SetFocus
            Exit Sub
        End If
        Set rsDupRepor = Nothing
    End If
    If (AddorEdit = "ADD" And labAddOrEdit = "ADD") Then
        Set rsDupRepor = New ADODB.Recordset
        Set rsDupRepor = gconDMIS.Execute("select rep_or from CSMS_RepOr where rep_or = " & N2Str2Null(txtRep_Or.Text))
        If Not rsDupRepor.EOF And Not rsDupRepor.BOF Then
            MsgSpeechBox "Repair Order Number Already Exist!"
            On Error Resume Next
            txtRep_Or.SetFocus
            Exit Sub
        End If
        Set rsDupRepor = Nothing
    End If
    'Updated by: IEBV 06282010 1131AM
    'Description:
    If COMPANY_CODE = "HCI" Then
        If Cbo_Rotype.Text = "" Then
           Cbo_Rotype.ListIndex = 0
        End If
    End If
    'Updated by: IEBV 06282010 1131AM
    'Description:

    'UPDATE BY   : MJP 10062008 0327 PM
    'DESCRIPTION : TO ENSURE WHEN THE USER EDIT THE REPAIR ORDER THAT THE PLATE NO HE WILL BE ENCODE IS EXISTING IN THE VEHICLE MASTER FILE
    If AddorEdit = "EDIT" Then
        Dim RSPLATE                                    As New ADODB.Recordset
        Set RSPLATE = gconDMIS.Execute("SELECT PLATE_NO FROM CSMS_CUSVEH WHERE PLATE_NO = '" & txtPlate_No & "'")
        If (RSPLATE.BOF And RSPLATE.EOF) Then
            MsgBox "Plate no. not existing in Vehicle Master File", vbExclamation, "CSMS"
            txtPlate_No.SetFocus
            Exit Sub
        End If
        Set RSPLATE = Nothing
    End If
    'UPDATE BY   : MJP 10062008 0327 PM

    Dim VTXTREP_OR                                      As String
    Dim VTXTestimateno                                  As String
    Dim VTXTROType                                      As String
    Dim VTXTSvc_No                                      As String
    Dim VTXTAcct_No                                     As String
    Dim VTXTNiym                                        As String
    Dim VTXTPlate_No                                    As String
    Dim VcboModel                                       As String
    Dim VTXTMake                                        As String
    Dim VTXTTerm                                        As String
    Dim VTXTSektion                                     As String
    Dim VTXTKm_rdg                                      As String
    Dim VTXTDte_recd                                    As String
    Dim VTXTCertific8                                   As String
    Dim VTXTDte_comp                                    As String
    Dim VTXTDte_Rel                                     As String
    Dim VtxtAddress                                     As String
    Dim VtxtVIN                                         As String
    Dim VTXTParticipat                                  As String
    Dim VcboRecd_by                                     As String
    Dim vtxtParticipation                               As String

    VTXTREP_OR = N2Str2Null(txtRep_Or.Text)
    VTXTestimateno = N2Str2Null(txtEstimateno.Text)
    'Updated by: IEBV 06282010 1133AM
    'Description:
    VTXTROType = N2Str2Null(Cbo_Rotype.Text)
    'Updated: IEBV 05100225PM
    'Description:
    'VTXTROType = N2Str2Null(txtROType.Text)
    VTXTSvc_No = N2Str2Null(txtSvc_No.Text)
    VTXTAcct_No = N2Str2Null(txtAcct_No.Text)
    VTXTNiym = N2Str2Null(txtNiym.Text)
    VTXTPlate_No = N2Str2Null(txtPlate_No.Text)
    VcboModel = N2Str2Null(cboModel.Text)
    VTXTMake = N2Str2Null(txtMake.Text)
    VTXTTerm = N2Str2Null(txtTerm.Text)
    VTXTSektion = N2Str2Null(txtSektion.Text)
    VTXTKm_rdg = N2Str2Null(txtKm_rdg.Text)
    VTXTDte_recd = N2Date2Null(txtDte_recd.Value)
    VTXTCertific8 = N2Str2Null(txtCertific8.Text)
    VTXTDte_comp = N2Date2Null(txtDte_comp.Text)
    VTXTDte_Rel = N2Date2Null(txtDte_Rel.Text)
    VtxtVIN = N2Str2Null(txtVIN.Text)
    VTXTParticipat = N2Str2Null(txtParticipat.Text)
    VcboRecd_by = N2Str2Null(SetCodeSA(cboRecd_by.Text))
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    vtxtParticipation = N2Str2Null(txtParticipation.Text)
    
    If AddorEdit = "ADD" And labAddOrEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_RepOr " & _
            "(rep_or,estimateno,rotype,svc_no,acct_no,niym,plate_no,model,term,sektion,Recd_by,km_rdg,dte_recd,certific8,dte_comp,dte_rel,VIN,participat,INSCDE,status,USERCDE,SAVEDATE,SAVETIME)" & _
            " values (" & VTXTREP_OR & ", " & VTXTestimateno & _
            ", " & VTXTROType & ", " & VTXTSvc_No & _
            ", " & VTXTAcct_No & ", " & VTXTNiym & _
            ", " & VTXTPlate_No & ", " & VcboModel & _
            ", " & VTXTTerm & ", " & VTXTSektion & _
            ", " & VcboRecd_by & ", " & VTXTKm_rdg & _
            ", " & VTXTDte_recd & ", " & VTXTCertific8 & _
            ", " & VTXTDte_comp & ", " & VTXTDte_Rel & _
            ", " & VtxtVIN & ", " & VTXTParticipat & _
            "," & vtxtParticipation & ", 'N' " & _
            "," & Vusercode & "," & VLastUpdate & _
            "," & VLastUpdateTime & ")"
        gconDMIS.Execute SQL_STATEMENT
    Else
        SQL_STATEMENT = "update CSMS_RepOr set" & _
            " rep_or = " & VTXTREP_OR & "," & _
            " estimateno = " & VTXTestimateno & "," & _
            " rotype = " & VTXTROType & "," & _
            " svc_no = " & VTXTSvc_No & "," & _
            " acct_no = " & VTXTAcct_No & "," & _
            " niym = " & VTXTNiym & "," & _
            " plate_no = " & VTXTPlate_No & "," & _
            " model = " & VcboModel & "," & _
            " term = " & VTXTTerm & "," & _
            " sektion = " & VTXTSektion & "," & _
            " recd_by = " & VcboRecd_by & "," & _
            " km_rdg = " & VTXTKm_rdg & "," & _
            " dte_recd = " & VTXTDte_recd & "," & _
            " certific8 = " & VTXTCertific8 & "," & _
            " dte_comp = " & VTXTDte_comp & "," & _
            " dte_rel = " & VTXTDte_Rel & "," & _
            " VIN = " & VtxtVIN & "," & _
            " participat = " & VTXTParticipat & "," & _
            " INSCDE = " & vtxtParticipation & "," & _
            " status = 'N'" & "," & _
            " USERCDE = " & Vusercode & "," & _
            " SAVEDATE = " & VLastUpdate & "," & _
            " SAVETIME = " & VLastUpdateTime & _
            " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "RO NO: " & txtRep_Or, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        If VTXTREP_OR <> "'" & PrevRoNumber & "'" Then
            SQL_STATEMENT = "update CSMS_RO_Det set rep_or = " & VTXTREP_OR & " where rep_or = '" & PrevRoNumber & "'"
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("EE", "BIILING SYSTEM", SQL_STATEMENT, labid, "", "DETAILS", "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            gconDMIS.Execute "update CSMS_JobClock set ro_no = " & VTXTREP_OR & " where ro_no = '" & PrevRoNumber & "'"
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("EE", "BIILING SYSTEM", SQL_STATEMENT, labid, "", "JOB CLOCK", "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET ASSIGNEDRO = " & VTXTREP_OR & " WHERE ASSIGNEDRO = '" & PrevRoNumber & "'")
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("EE", "BIILING SYSTEM", SQL_STATEMENT, labid, "", "EMPLOYEE", "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If

        gconDMIS.Execute "update CSMS_RepairOrder set Writer = " & N2Str2Null(cboRecd_by.Text) & ", model = " & VcboModel & ", plate_no = " & VTXTPlate_No & ", acct_no = " & VTXTAcct_No & ", RO_NO = " & VTXTREP_OR & ", AppointmentDate = " & VTXTDte_recd & " where ro_no = '" & PrevRoNumber & "'"
        'gconDMIS.Execute "Update CSMS_CUSVEH SET CUSCDE = " & VTXTAcct_No & " where Plate_No = " & VTXTPlate_No & ""

        '************************************************************************************
        'UPDATE BY   : MJP 08042008 11:49 PM
        'DESCRIPTION : TO UPDATE ALSO THE PMS JOBS
            gconDMIS.Execute ("UPDATE CSMS_PMS_JOB_DET SET REP_OR = " & VTXTREP_OR & " WHERE REP_OR = '" & PrevRoNumber & "'")
        'UPDATE BY   : MJP 08042008 11:49 PM
        '************************************************************************************
        'Updated: IEVB 06100248PM
        'Description: For HCI
        If COMPANY_CODE = "HCI" Then
            Me.Cbo_Rotype.Visible = False
            Me.txttype.Visible = True
        End If
        'Updated: IEBV 06100248PM
        'Description: For HCI
    End If

CONT:
    Call rsRefresh
    rsREPOR.Find "REP_OR = " & VTXTREP_OR
    Picture5.Enabled = True:    pic3.Enabled = True
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0: ShowVBError
End Sub

Private Sub cmdSaveFollowUp_Click()
    If Trim(txtCALLED_RESULT.Text) <> "" Then
        If MsgBox("Follow up Result will be saved, are you sure?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
            
        SQL_STATEMENT = "update CSMS_RepOr Set CALLED_FOLLOWUP = 1, CALLED_RESULT = " & N2Str2Null(txtCALLED_RESULT.Text) & " Where ID = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "INV NO: " & txtInvoiceNo & " - FOLLOW UP", "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call Command6_Click
        MessagePop InfoFriend, "RO Information Updated", "Follow up Succesfully Save", 1000
        
        Call rsRefresh
        rsREPOR.Find "ID = " & labid & ""
        Call StoreMemVars
    End If
End Sub

Private Sub cmdSelect_Click()
    Call rsRefresh
    Call StoreMemVars
    
    picMain.Visible = True
    picSEARCH.Visible = False
End Sub

Private Sub Command1_Click()
    If Module_Access(LOGID, "RO ADVANCE OPTIONS", "SYSTEM") = False Then Exit Sub

    Dim rsACCOUNT_CODE                                      As New ADODB.Recordset
    Set rsACCOUNT_CODE = gconDMIS.Execute("SELECT CODE FROM CSMS_RO_DET WHERE REP_OR = '" & txtRep_Or & "' AND WCODE IN('S','C') AND CODE IS NULL")
    If Not rsACCOUNT_CODE.EOF And Not rsACCOUNT_CODE.BOF Then
       MsgBox "There is an internal transaction which" & vbCrLf & "Acct Code is not yet been selected.", vbExclamation, "INFORMATION"
       Exit Sub
    End If
    Set rsACCOUNT_CODE = Nothing
     
     
    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Set PDI Ro, please refresh your Billing System", 1000
        Exit Sub
    End If
    
    If CheckIfAllJobIsFinish = False Then
        MsgBox "Repair Order # " & txtRep_Or.Text & " Job(s) not Yet Finish" & vbCrLf & "Please Finish All Job Before Billing this RO.", vbInformation, "Billing System"
        Exit Sub
    End If
    
    If txtInvoiceNo.Text = "" Then
        If MsgBox("Set this Repair order to a PDI Repair, Are you sure?", vbYesNo + vbQuestion, "Warning") = vbYes Then
            SQL_STATEMENT = "update CSMS_RepOr set invoice = 'PDI RO', dte_comp = dte_recd Where ID = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT

            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("P", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "INV NO: PDI RO", "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            gconDMIS.Execute "update CSMS_RepairOrder set status = 'Billed' Where RO_NO = '" & txtRep_Or.Text & "'"
            MessagePop InfoFriend, "Repair order Information Updated", "Repair order Sucessfully tag as PDI RO!", 1000

            Call rsRefresh
            rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
            cmdCancel.Value = True
        End If
    End If
End Sub

Private Sub Command2_Click()
    If Module_Access(LOGID, "RO ADVANCE OPTIONS", "SYSTEM") = False Then Exit Sub

    If txtInvoiceNo.Text = "" Then
        If MsgBox("This is a Warranty Repair Order and will be Release, Are you sure?", vbYesNo + vbQuestion, "Warning") = vbYes Then
            gconDMIS.Execute "update CSMS_RepOr set invoice = 'WAR RO', dte_comp = dte_recd, dte_rel = dte_recd Where ID = " & labid.Caption
            gconDMIS.Execute "update CSMS_RepairOrder set status = 'Released' Where RO_NO = '" & txtRep_Or.Text & "'"
            MsgBox "Repair Order Successfully Release...", vbOKOnly + vbInformation, "Deleted..."
            
            Call rsRefresh
            rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
            cmdCancel.Value = True
        End If
    End If
End Sub

Private Sub Command3_Click()
    If Module_Access(LOGID, "EDIT HOURS WORK", "SYSTEM") = False Then Exit Sub
    On Error Resume Next

    txtDET_HRS.Enabled = True
    txtDET_HRS.SetFocus
End Sub

Private Sub Command4_Click()
    'UPDATE BY   : MJP09232009
    'DESCRIPTION : THIS IS TO HAVE A VALIDITION FOR THE RO NO IN PART MODULE AND SERVICE MODULE
        If CheckIfTheresItemIssued = True Then
            MsgBox "You cannot change the RO no when theres already issued Item", vbInformation, "Info."
        Else
            txtRep_Or.Locked = False
            txtRep_Or.SetFocus
        End If
    'UPDATE BY   : MJP09232009
End Sub

Private Sub Command5_Click()
    picMain.Visible = False
    picSEARCH.Visible = True
End Sub

Private Sub Command6_Click()
    Frame1.Enabled = False
    Frame2.Enabled = True
    Picture1.Enabled = True
    Picture2.Enabled = True
    Picture3.Enabled = True
    cmdFollowUp.Visible = False
End Sub

Private Sub Command7_Click()
    If Module_Access(LOGID, "CHANGE SERVICE ADVISER", "SYSTEM") = False Then Exit Sub
    cboRecd_by.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Picture1.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (BILLING SYSTEM)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "BILLING SYSTEM")

        Case Shift = 1 And vbKeyF3:
            Call cmdFind_Click

        Case vbKeyDelete
            If Left(Me.ActiveControl.Name, 3) = "cbo" Then Me.ActiveControl.ListIndex = -1
            
        Case vbKeyReturn
            'If Me.ActiveControl.Name = "txtRep_or" Then
            '    Dim rsReporDup      As ADODB.Recordset
            '    Set rsReporDup = New ADODB.Recordset
            '    Set rsReporDup = gconDMIS.Execute("select rep_or from CSMS_RepOr where rep_or = " & N2Str2Null(txtRep_Or.Text))
            '    If rsReporDup.EOF And rsReporDup.BOF Then SendKeys "{TAB}"
            '    Set rsReporDup = Nothing
            'Else
            If Picture1.Visible = True Then
                If Me.ActiveControl.Name = "txtVIN" Then
                    If txtPlate_No.Text <> "" Then
                        Me.Enabled = False: frmCSMSROCusveh.Show: frmCSMSROCusveh.ZOrder 0
                    Else
                        MsgSpeechBox "Plate Number must be inputed! Please enter 000000 if unknown"
                        On Error Resume Next
                        txtPlate_No.SetFocus
                    End If
                ElseIf Me.ActiveControl.Name = "txtAcct_No" Then
                    If txtAcct_No.Text = "" Then SendToBack Else SendKeys "{TAB}"
                ElseIf Me.ActiveControl.Name = "txtParticipat" Then
                    If chkParticipat.Value = 1 And txtParticipat.Text = "" Then
                        Call SendToBack
                        RO_OR_ESTI_OR_PART = "PART"
                    Else
                        SendKeys "{TAB}"
                    End If
                Else
                    'If cmdAddJobs.ZOrder = 1 Then
                    MoveKeyPress KeyCode
                    'End If
                End If
            End If
            
        Case vbKeyEscape
            SSTab1.SelectedItem = 0:
            Call SendToBackDisc
            Call SendToBack
            
        Case vbKeyF1                                        'Notes After Follow Up
            If Picture1.Visible = True Then
                If txtInvoiceNo.Text <> "" Then
                    Frame1.Enabled = False:         Frame2.Enabled = False
                    Picture1.Enabled = False:       Picture2.Enabled = False
                    Picture3.Enabled = False
                    cmdFollowUp.Visible = True
                    cmdFollowUp.ZOrder 0:
                    
                    cmdFollowUp.Enabled = True:     chkCALLED_FOLLOWUP.Enabled = True
                    txtCALLED_RESULT.Enabled = True
                    
                    On Error Resume Next
                    txtCALLED_RESULT.SetFocus
                End If
            End If
            
        Case vbKeyF2
            If Picture1.Visible = True Then
                'If txtDte_Rel.Text = "" Then
                If txtInvoiceNo.Text = "" Then
                    If chkParticipat.Value = 1 Then
                        If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
                            MessagePop InfoFriend, "Repair order Information", "Repair order already Set Internal RO, please refresh your Billing System", 1000
                            Exit Sub
                        End If
                        
                        Call SendToBack
                        'UPDATE BY   : MJP012609 0137PM
                        'DESCRIPTION : TO DISABLE THE NAVIGATION BUTTON AND SOME OPTION BUTTON
                            Picture1.Enabled = False
                            Frame2.Enabled = False
                            pic3.Enabled = False
                            Picture5.Enabled = False
                        'UPDATE BY   : MJP012609 0137PM
                        Picture6.ZOrder 0
                        Picture6.Visible = True
                        Call StoreParticipationEntry(txtRep_Or)
                    End If
                End If
            End If
            
        Case vbKeyF3
            If picSEARCH.Visible = True Then
                txtSearch.SetFocus
                Exit Sub
            End If
            SSTab1.SelectedItem = 1
            
        Case vbKeyF4
            If picSEARCH.Visible = True Then Exit Sub
            SSTab1.SelectedItem = 2
            
        Case vbKeyF5
            If picSEARCH.Visible = True Then Exit Sub
            If SSTab1.SelectedItem <> 3 Then
                SSTab1.SelectedItem = 3
                Exit Sub
            End If
            
            If Picture1.Visible = True Then
                SSTab1.SelectedItem = 3
                If txtInvoiceNo.Text = "" Then
                    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
                        MessagePop InfoFriend, "Repair order Information", "Repair order already Set Internal RO, please refresh your Billing System", 1000
                        Exit Sub
                    End If
    
                    If MsgBox("Add Miscelleneous Charges?", vbQuestion + vbYesNo, "Add Miscelleneous") = vbYes Then
                        Call AddMaterials
                        'UPDATE BY   : IEBV 05262010 0405:PM
                        'DESCRIPTION : To enable and disable the miscellaneous text area if the company code is HOT or HQA
                        
                        If COMPANY_CODE = "HCA" Or COMPANY_CODE = "HQA" Or COMPANY_CODE = "HCC" Then
                            lbldetail.Visible = True
                            txtdetail.Visible = True
                            cmdAddMaterials.Height = 6435
                            piccontol.Top = 5160
                        Else
                            lbldetail.Visible = False
                            txtdetail.Visible = False
                            piccontol.Top = 4200
                            cmdAddMaterials.Height = 5475

                        End If
                        'UPDATE BY   : IEBV 05262010 0405:PM
                        'DESCRIPTION : To enable and disable the miscellaneous text area if the company code is HOT or HQA

                    End If
                End If
            End If
            
        Case vbKeyF6
            If picSEARCH.Visible = True Then Exit Sub
            SSTab1.SelectedItem = 4
            
        Case vbKeyF7
            If picSEARCH.Visible = True Then Exit Sub
            If SSTab1.SelectedItem <> 0 And txtInvoiceNo.Text = "" Then
                If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
                    MessagePop InfoFriend, "Repair order Information", "Repair order already Set Internal RO, please refresh your Billing System", 1000
                    Exit Sub
                End If
    
                'UPDATE BY   : MJP012609 0137PM
                'DESCRIPTION : TO DISABLE THE NAVIGATION BUTTON AND SOME OPTION BUTTON
                    Picture1.Enabled = False
                    Frame2.Enabled = False
                    pic3.Enabled = False
                    Picture5.Enabled = False
                'UPDATE BY   : MJP012609 0137PM
                Call SendToFrontDisc
            End If
            
        Case vbKeyF8
            
            If picSEARCH.Visible = True Then Exit Sub
            If Picture1.Visible = False Then Exit Sub
            If txtInvoiceNo.Text <> "" Then Exit Sub
            'If Picture1.Visible = True And txtInvoiceNo.Text = "" Then
            If Picture1.Visible = True Then
                SSTab1.SelectedItem = 0
                rsREPOR.Find "ID = " & labid.Caption & ""
                Call StoreMemVars
                
                Call SendToBack
'                Call ImportParts
'                Call ImportMaterials
'                Call ImportAccessories

'                If chkParticipat.Value = 1 Then Call UpdateParticipation
            End If
            
        Case vbKeyF9
            Dim ReferenceNumber                    As String
            If Function_Access(LOGID, "Acess_CancelEntry", "BILLING SYSTEM") = False Then Exit Sub
            
            If picSEARCH.Visible = True Then Exit Sub
            If pic3.Enabled = False Then Exit Sub
            
            If Picture1.Visible = True Then
                If txtInvoiceNo.Text <> "" Then
                    If txtInvoiceNo.Text = "INT RO" Then
                        If CheckSJINTRONum(Null2String(rsREPOR!REP_OR)) <> "" Then
                            MsgBox "Warning: No Invoice is Allowed to be cancelled once its already Imported in Accounting!", vbExclamation, "Not Allowed!"
                            Exit Sub
                        End If
                    Else
                        If CheckORNum(Null2String(rsREPOR!invoice)) <> "" Then
                            MsgBox "Warning: No Invoice is Allowed to be cancelled once its already paid in Cashier!", vbExclamation, "Not Allowed!"
                            Exit Sub
                        End If
                        If CheckSJNum(Null2String(rsREPOR!invoice)) <> "" Then
                            MsgBox "Warning: No Invoice is Allowed to be cancelled once its already Imported in Accounting!", vbExclamation, "Not Allowed!"
                            Exit Sub
                        End If
                    End If

                    If txtInvoiceNo.Text = "INT RO" Then
                        ReferenceNumber = "Repair Order No: " & txtRep_Or.Text
                    Else
                        ReferenceNumber = "Invoice No: " & txtInvoiceNo.Text
                    End If

                    If MsgBox("Cancel " & ReferenceNumber, vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
                        If COMPANY_CODE = "HGC" Then
                            With FrmCancelTransaction
                                CANCEL_ANS = "NO"
                                .lblState.Caption = "Cancel SI NO: " & txtInvoiceNo
                                .lblTransaction_type = "SI"
                                .LblTransactionNo = txtRep_Or.Text
                                FrmCancelTransaction.Show 1
                            End With

                            If CANCEL_ANS = "NO" Then Exit Sub
                        End If

                        SQL_STATEMENT = "update CSMS_RepOr set PRIN_DTE = NULL, invoice = NULL, dte_comp = NULL, dte_rel = NULL, status = 'N' where id = " & labid.Caption
                        gconDMIS.Execute SQL_STATEMENT
                        'NEW LOG AUDIT-----------------------------------------------------
                            Call NEW_LogAudit("C", "BILLING SYSTEM", SQL_STATEMENT, labid, "R", "INV NO: " & txtInvoiceNo, "", "")
                        'NEW LOG AUDIT-----------------------------------------------------

                        SQL_STATEMENT = "Update CSMS_Repairorder set Status = 'Finish Job' where ro_no='" & txtRep_Or.Text & "'"
                        gconDMIS.Execute SQL_STATEMENT
                        'NEW LOG AUDIT-----------------------------------------------------
                            Call NEW_LogAudit("C", "SERVICE COUNTER", SQL_STATEMENT, labid, "", "INV NO: " & txtInvoiceNo, "", "")
                        'NEW LOG AUDIT-----------------------------------------------------

                        'UPDATE BY   : MJP012709 0611PM
                        'DESCRIPTION :
                            gconDMIS.Execute ("UPDATE CSMS_RO_DET SET DONE = 'Y', STATUS = 'Y' WHERE LIVIL = '1' AND REP_OR = '" & txtRep_Or & "'")
                        'UPDATE BY   : MJP012709 0611PM
                        
                        If COMPANY_CODE = "HGC" Then
                            gconDMIS.Execute "Update CSMS_INVOICE set Status = 'C',CANCELLEDDATE = '" & CDate(LOGDATE) & "' where INVOICENO='" & txtInvoiceNo.Text & "'"
                        End If

                        gconDMIS.Execute "update PMIS_Ord_Hd set Status = 'P' where (STATUS <> 'C' OR STATUS <> 'N') AND trantype = 'RIV' and rono = '" & txtRep_Or.Text & "'"
                        gconDMIS.Execute "update PMIS_Ord_Hd set In_Process = 'Y' where (STATUS <> 'C' OR STATUS <> 'N') AND trantype = 'RIV' and rono = '" & txtRep_Or.Text & "'"
                        gconDMIS.Execute "update PMIS_Ord_Hist set In_Process = 'Y' where (STATUS <> 'C' OR STATUS <> 'N') AND trantype = 'RIV' and rono = '" & txtRep_Or.Text & "'"

                        'gconDMIS.Execute ("UPDATE CSMS_REPOR SET PRIN_DTE = NULL WHERE ID = " & LABID.Caption)
                        
                        MessagePop InfoFriend, "RO Information Updated", "Repair Order Information Sucessfully Unbilled!", 1000
                        rsRefresh
                        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
                        StoreMemVars
                    End If
                End If
            End If

        Case vbKeyF10
            If Module_Access(LOGID, "UNRELEASED REPAIR ORDER", "SYSTEM") = False Then Exit Sub
            
            If picSEARCH.Visible = True Then Exit Sub
            If Picture1.Visible = True Then
                If txtDte_Rel.Text <> "" Then
                    If MsgBox("Unreleased Repair Order Number " & txtRep_Or.Text & "", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
                        SQL_STATEMENT = "update CSMS_RepOr set status='N', dte_rel = NULL where id = " & labid.Caption
                        gconDMIS.Execute SQL_STATEMENT
                        'NEW LOG AUDIT-----------------------------------------------------
                            Call NEW_LogAudit("UR", "BILLING SYSTEM", SQL_STATEMENT, labid, "", "INV NO: " & txtInvoiceNo, "", "")
                        'NEW LOG AUDIT-----------------------------------------------------

                        Dim vRONO              As String
                        vRONO = Left(txtRep_Or, 1) & Right(txtRep_Or, 6)
                        SQL_STATEMENT = "update PMIS_Ord_Hd set In_Process = 'Y' where trantype = 'RIV' and Tranno = '" & vRONO & "'"
                        gconDMIS.Execute SQL_STATEMENT
                        'NEW LOG AUDIT-----------------------------------------------------
                        '    Call NEW_LogAudit("UB", "TECHNICIAN REPORT", "", "", "", "TECHNICIAN ATTENDANCE : " & cboMonth & " " & cboYEAR, "", "")
                        'NEW LOG AUDIT-----------------------------------------------------

                        gconDMIS.Execute "update PMIS_Ord_Hist set In_Process = 'Y' where trantype = 'RIV' and rono = '" & vRONO & "'"
                        
                        SQL_STATEMENT = "Update CSMS_RO_dET set Status = 'Y' where REP_OR = '" & txtRep_Or.Text & "'"
                        gconDMIS.Execute SQL_STATEMENT
                        'NEW LOG AUDIT-----------------------------------------------------
                            Call NEW_LogAudit("EE", "BILLING SYSTEM", "", "", "", "UNRELEASED DET: " & Date, "", "")
                        'NEW LOG AUDIT-----------------------------------------------------
                        
                        gconDMIS.Execute "Update CSMS_Repairorder set Status = 'Billed' where ro_no = '" & txtRep_Or.Text & "'"
                        'NEW LOG AUDIT-----------------------------------------------------
                            Call NEW_LogAudit("EE", "BILLING SYSTEM", "", "", "", "UNRELEASED RO: " & Date, "", "")
                        'NEW LOG AUDIT-----------------------------------------------------
    
                        MessagePop InfoFriend, "RO Information Updated", "Repair Order Information Sucessfully Unreleased!", 1000
                        Call rsRefresh
                        rsREPOR.Find "REP_OR = '" & txtRep_Or.Text & "'"
                        Call StoreMemVars
                    End If
                End If
            End If
            
        Case vbKeyF11
            'If txtInvoiceNo.Text = "" Then Picture5.ZOrder 0
            'MsgBox "Sorry, this feature is not yet available", vbInformation, "For Update"
            If picSEARCH.Visible = True Then Exit Sub
            'If Label5.Enabled = True Then
                picCustLimit.Visible = True
                picCustLimit.ZOrder 0
                SetCustLimit
            'Else
             '   picCustLimit.Visible = False
             '   picCustLimit.ZOrder 1
            'End If

        Case vbKeyF12
            If picSEARCH.Visible = True Then Exit Sub
            If Picture1.Visible = False Then Exit Sub
            
            If txtInvoiceNo.Text <> "" Then
                'do nothing
                MsgBox "Repair Order Information Cannot be Edited once it is invoiced", vbInformation, "CSMS"
                Exit Sub

            Else
                If txtPlate_No.Text <> "" Then
                    
                    
                    'JBF Aug232010
                    'Call frm.SelectSQl("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(txtPlate_No) & "", "BILLING SYSTEM", GetPlateId(txtPlate_No), txtAcct_No, txtNiym, txtPlate_No)
                    'frm.Show 1
                    'Me.Enabled = False: frmCSMSROCusveh.Show: frmCSMSROCusveh.ZOrder 0: frmCSMSROCusveh.Frame1.Enabled = False: frmCSMSROCusveh.cmdSave.Enabled = False
                    '*************
                       
                    Call frm.SelectSQl("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(txtPlate_No) & "", "BILLING SYSTEM", GetPlateId(txtPlate_No), txtAcct_No, txtNiym, txtPlate_No)
                    frm.Show 1

                
                Else
                    MsgSpeechBox "Plate Number must be inputed! Please enter 000000 if unknown"
                    Exit Sub
                End If
            End If
            
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    flag = False
    Screen.MousePointer = 11
    SSTab1.SelectedItem = 0
    Set frm = New frmCSMSROCusveh
    Set FRMx = New frmCSMS_MasterSearchCustomer
    
    Call CenterMe(frmMain, Me, 1)
    Call SendToBack
    
    'Updated by:    IEBV 06282010 1125AM
    'Description:   To make Rotype Visible
    If COMPANY_CODE = "HCI" Then
        lbl_rotype.Visible = True
        Cbo_Rotype.Visible = True
        txttype.Visible = True
    End If
    'Updated by:    IEBV 06282010 1125AM
    'Description:
    
    ROSHOW = True
    cmdFollowUp.Enabled = False
    
    If UCase(LOGLEVEL) <> "ADM" And UCase(LOGLEVEL) <> "AUTHOR" Then
        Picture3.Visible = False
    End If
    
    Frame1.Enabled = False
    Frame2.ZOrder 0
    Frame1.ZOrder 0
    Picture1.Visible = True
    Picture2.Visible = False
    
    DoEvents
    Call InitializeRC
    Call initMemvars
    Call InitCbo
    cboJobCode.Enabled = False
    Call InitJobs
    Call InitParts
    Call InitMaterials
    Call InitAccessories
    
    DoEvents
    Call txtSearch_Change
    DoEvents
    Screen.MousePointer = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.FontUnderline = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ROSHOW = False: Set frmCSMSDataEntry = Nothing
End Sub

Private Sub frm_SaveChanges(xPLATE_NO As String, xWARR_CER As String, xMake As String, xMODEL As String, xSERIAL As String, xDESCRIPTION As Variant, FromFrom As String)
    If FromFrom = "BILLING SYSTEM" Then
        txtPlate_No.Text = xPLATE_NO
        txtCertific8.Text = xWARR_CER
        txtVIN.Text = xSERIAL
        cboModel.Text = xMODEL
        txtMake.Text = xDESCRIPTION
        
        Unload frm
    End If
End Sub

Private Sub FRMx_SelectionMade(ByVal Xcode As String, xName As String, FromForm As String)
    If FromForm = "BILLING SYSTEM" Then
        txtAcct_No.Text = Xcode
        txtNiym.Text = xName
        
        Unload FRMx
    ElseIf FromForm = "BILLING INSURANCE" Then
        txtParticipat.Text = Xcode
        txtParticipation.Text = xName
        
        Unload FRMx
    End If
End Sub

Private Sub grdDetails_DblClick()
    If grdDetails.Row = 1 Then SSTab1.SelectedItem = 1
    If grdDetails.Row = 2 Then SSTab1.SelectedItem = 2
    If grdDetails.Row = 3 Then SSTab1.SelectedItem = 3
    If grdDetails.Row = 4 Then SSTab1.SelectedItem = 4
End Sub

Private Sub cboJcode_Click()
    If optByCode.Value = True Then
        If cboJcode.Text <> "" Then cboJobCode.Text = setJobDesc(cboJcode.Text)
        txtJobPostCode.Text = setJobPOcode(cboJobCode.Text)
        txtJobRate.Text = setJobRate(cboJobCode.Text)
        If AddorEdit = "ADD" Then txtJobDetail.Text = setJobDetail(cboJobCode.Text)
    End If
End Sub

Private Sub cbomatcode_LostFocus()
    If cboMatCode.Text <> "" Then cboMaterial.Text = SetMatDisc(cboMatCode.Text)
    txtMatUnitPrice.Text = SetMatPrice(cboMatCode.Text)
    txtMatPOCode.Text = SetMatPOCode(cboMatCode.Text)
    txtMatAmount.Text = txtMatUnitPrice.Text
End Sub

Private Sub Label79_Click()
    If txtInvoiceNo <> "" Then
        MsgBox "Repair Order Information Cannot be Edited once it is invoiced", vbInformation, "CSMS"
        Exit Sub
    End If

    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Billing System", 1000
        Exit Sub
    End If
    
    If txtPlate_No.Text <> "" Then
        AddorEdit = "EDIT": labAddOrEdit.Caption = "EDIT"
        'UPDATE BY   : MJP 09212009 0357PM
            Call frm.SelectSQl("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(txtPlate_No) & "", "BILLING SYSTEM", GetPlateId(txtPlate_No), txtAcct_No, txtNiym, txtPlate_No)
            frm.Show 1
        'UPDATE BY   : MJP 09212009 0357PM
    Else
        MsgSpeechBox "Plate Number must be inputed! Please enter 000000 if unknown"
        On Error Resume Next
        txtPlate_No.SetFocus
    End If
End Sub

Function GetPlateId(XPLATENO As String) As Long
    Dim rstmp                                   As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT ID FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(XPLATENO) & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GetPlateId = rstmp!ID
    End If
    Set rstmp = Nothing
End Function

Private Sub Label79_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label79.FontUnderline = True
End Sub

Private Sub lblStatus_Click(Index As Integer)
    Screen.MousePointer = 11
    Dim RSUPLOAD                                        As New ADODB.Recordset
    Dim REC                                             As XtremeReportControl.ReportRecord
    Dim SQLTXT As String
    
    If Index = 0 Then
        Set RSUPLOAD = gconDMIS.Execute("SELECT REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL, NULL as DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND DTE_COMP IS NULL ORDER BY REP_OR DESC")
    ElseIf Index = 1 Then
        Set RSUPLOAD = gconDMIS.Execute("SELECT REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,'Y' as DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND DTE_COMP IS NOT NULL AND DTE_REL IS NULL AND INVOICE IS NOT NULL ORDER BY REP_OR DESC")
    ElseIf Index = 2 Then
        Set RSUPLOAD = gconDMIS.Execute("SELECT REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,'Y' as DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND DTE_REL IS NOT NULL  AND INVOICE IS NOT NULL ORDER BY REP_OR DESC")
    Else
        'UPDATE BY: NVB 12/5/2009
        'DESCRIPTION: AS REQUESTED BY HLI,
        '             IN BILLING THEY WANT TO SEE THE REPAIR ORDER THAT IS FINISHED JOB
        SQLTXT = "SELECT A.REP_OR, A.INVOICE, A.NIYM, A.PLATE_NO, A.VIN, A.MODEL," & vbCrLf
        SQLTXT = SQLTXT & "A.ID, A.DTE_REL,B.DONE FROM CSMS_REPOR A INNER JOIN (" & vbCrLf
        SQLTXT = SQLTXT & "SELECT REP_OR,DONE FROM(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT REP_OR,DONE FROM (" & vbCrLf
        SQLTXT = SQLTXT & "SELECT DISTINCT(A.REP_OR) AS REP_OR, A.INVOICE, A.NIYM, A.PLATE_NO, A.VIN, A.MODEL," & vbCrLf
        SQLTXT = SQLTXT & "A.ID, A.DTE_REL,B.DONE FROM CSMS_REPOR A INNER JOIN CSMS_RO_DET B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.REP_OR = B.REP_OR WHERE A.TRANSTYPE = 'R' AND A.INVOICE IS NULL" & vbCrLf
        SQLTXT = SQLTXT & ")T GROUP BY REP_OR,DONE HAVING COUNT(REP_OR) = 1" & vbCrLf
        SQLTXT = SQLTXT & ")X WHERE DONE = 'Y'" & vbCrLf
        SQLTXT = SQLTXT & ")B ON A.REP_OR = B.REP_OR" & vbCrLf

        Set RSUPLOAD = gconDMIS.Execute(SQLTXT)
    End If
    
    rptRO.Records.DeleteAll
    While Not RSUPLOAD.EOF
        Set REC = rptRO.Records.Add
        REC.AddItem (Trim(RSUPLOAD!NIYM))
        REC.AddItem (Trim(RSUPLOAD!REP_OR))
        REC.AddItem (Trim(RSUPLOAD!invoice))
        REC.AddItem (Trim(RSUPLOAD!PLATE_NO))
        REC.AddItem (Trim(RSUPLOAD!Vin))
        REC.AddItem (Trim(RSUPLOAD!Model))
        REC.AddItem (Trim(RSUPLOAD!DTE_rel))
        REC.AddItem (Trim(RSUPLOAD!ID))
        REC.AddItem (Trim(RSUPLOAD!DONE))
    
        RSUPLOAD.MoveNext
        Set REC = Nothing
    Wend
    rptRO.Populate
    txtSearch.SetFocus
    Screen.MousePointer = 0
End Sub

Function CHECKIFFINISHEDJOB(ro As String, FIELD_NIYM As String) As Boolean
        Dim SQLTXT As String
        Dim rstmp As New ADODB.Recordset
        
        SQLTXT = "SELECT top 10 A.REP_OR, A.INVOICE, A.NIYM, A.PLATE_NO, A.VIN, A.MODEL," & vbCrLf
        SQLTXT = SQLTXT & "A.ID, A.DTE_REL,B.DONE FROM CSMS_REPOR A INNER JOIN (" & vbCrLf
        SQLTXT = SQLTXT & "SELECT REP_OR,DONE FROM(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT REP_OR,DONE FROM (" & vbCrLf
        SQLTXT = SQLTXT & "SELECT DISTINCT(A.REP_OR) AS REP_OR, A.INVOICE, A.NIYM, A.PLATE_NO, A.VIN, A.MODEL," & vbCrLf
        SQLTXT = SQLTXT & "A.ID, A.DTE_REL,B.DONE FROM CSMS_REPOR A INNER JOIN CSMS_RO_DET B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.REP_OR = B.REP_OR WHERE A.TRANSTYPE = 'R'" & vbCrLf
        SQLTXT = SQLTXT & ")T GROUP BY REP_OR,DONE HAVING COUNT(REP_OR) = 1" & vbCrLf
        SQLTXT = SQLTXT & ")X WHERE DONE = 'Y'" & vbCrLf
        SQLTXT = SQLTXT & ")B ON A.REP_OR = B.REP_OR WHERE " & FIELD_NIYM & "  LIKE '%" & Null2String(ro) & "%'" & vbCrLf

        Set rstmp = gconDMIS.Execute(SQLTXT)
        
        If Not (rstmp.EOF And rstmp.BOF) Then
            CHECKIFFINISHEDJOB = True
        Else
            CHECKIFFINISHEDJOB = False
        End If
        
        Set rstmp = Nothing
End Function

Private Sub lstJObs_DblClick()
    If kcnt = 0 Then Exit Sub
        
'    If txtInvoiceNo <> "" Then
'        MessagePop InfoFriend, "Repair order Information", "Repair Order Already Invoiced. Details Information cannot be edit", 1000
'        Exit Sub
'    End If
    
'    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
'        MessagePop InfoFriend, "Repair order Information", "Repair order details cannot be edit once its already invoice", 1000
'        Exit Sub
'    End If
    
    AddorEdit = "EDIT"
    Call SendToBack
    cmdAddJobs.ZOrder 0: fraAddJobs.ZOrder 0
    If txtInvoiceNo.Text = "" Then
        fraAddJobs.Enabled = True
        'If CheckifDetailsIsSublet(lstJobs.SelectedItem.SubItems(6)) = False Then
            'Call EnabledSubletJobObject(True)
        'Else
        '    Call EnabledSubletJobObject(False)
        '    MessagePop InfoFriend, "Info", "Sublet details cannot be edited here, please go to sublet purchasing for the editing of values"
        'End If
    Else
        fraAddJobs.Enabled = False
    End If

    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
        Picture1.Enabled = False: Picture5.Enabled = False: pic3.Enabled = False: Frame2.Enabled = False
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

    txtDET_HRS.Enabled = False
    fraAddJobs.Caption = "Edit Jobs"
    Call StoreJobsEntry(Trim(Me.lstJObs.SelectedItem.SubItems(6)))
End Sub

Sub EnabledSubletJobObject(COND As Boolean)
    txtJobLineNo.Enabled = COND
    cboJcode.Enabled = COND
    cboJobCode.Enabled = COND
    txtJobRate.Enabled = COND
    cboJobChargeTo.Enabled = COND
    'optByPerc.Enabled = COND
    'txtJobDiscount.Enabled = COND
    'optByAmt.Enabled = COND
    'txtJobDiscountAmt.Enabled = COND
    cboTechnician.Enabled = COND
    txtDET_HRS.Enabled = COND
    Command3.Enabled = COND
    txtJobDetail.Enabled = COND
    cmdJobDelete.Enabled = COND
End Sub

Private Sub lstMaterials_DblClick()
'UPDATE BY   : IEBV 05262010 0405:PM
'DESCRIPTION : Stops the program closing when errors occur specialy when overflow error occurs

    On Error Resume Next
'UPDATE BY   : IEBV 05262010 0405:PM
'DESCRIPTION : Stops the program closing when errors occur specialy when overflow error occurs

    If Mcnt = 0 Then Exit Sub
    
'    If txtInvoiceNo <> "" Then
'        MessagePop InfoFriend, "Repair order Information", "Repair Order Already Invoiced. Details Information cannot be edit", 1000
'        Exit Sub
'    End If
    
'    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
'        MessagePop InfoFriend, "Repair order Information", "Repair order details cannot be edit once its already invoice", 1000
'        Exit Sub
'    End If
    'Updated By : IEBV 05262010
     cmdAddMaterials.Visible = True
    'Updated By : IEBV 05262010
    AddorEdit = "EDIT"
    Call SendToBack
    cmdAddMaterials.ZOrder 0: fraAddMaterials.ZOrder 0
    If txtInvoiceNo.Text = "" Then
        fraAddMaterials.Enabled = True
        'If CheckifDetailsIsSublet(lstMaterials.SelectedItem.SubItems(8)) = False Then
        '    Call EnabledSubletMaterialObject(True)
        'Else
        '    Call EnabledSubletMaterialObject(False)
        '    MessagePop InfoFriend, "Info", "Sublet details cannot be edited here, please go to sublet purchasing for the editing of values"
        'End If
    Else
        fraAddMaterials.Enabled = False
    End If
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
        Picture1.Enabled = False: Picture5.Enabled = False: pic3.Enabled = False: Frame2.Enabled = False
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†
    
    
    'UPDATE BY   : IEBV 05262010 0405:PM
    'DESCRIPTION : To enable and disable the miscellaneous text area if the company code is HOT or HQA
        If COMPANY_CODE = "HQA" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HCC" Then
            If UCase(Me.lstMaterials.SelectedItem.SubItems(1)) = "MISC" Or UCase(Me.lstMaterials.SelectedItem.SubItems(1)) = "MISC." Or UCase(Me.lstMaterials.SelectedItem.SubItems(1)) = "MISCELLANEOUS" Then
                lbldetail.Visible = True
                txtdetail.Visible = True
                cmdAddMaterials.Height = 6435
                piccontol.Top = 5160
    
            Else
                lbldetail.Visible = False
                txtdetail.Visible = False
                cmdAddMaterials.Height = 5475
                piccontol.Top = 4200
                
            End If
        Else
            lbldetail.Visible = False
            txtdetail.Visible = False
            piccontol.Top = 4200
            cmdAddMaterials.Height = 5475
    End If
    'DESCRIPTION : To enable and disable the miscellaneous text area if the company code is HOT or HQA
    'UPDATE BY  : IEBV 05262010 0405:PM
    
    fraAddMaterials.Caption = "Edit Materials"
    Call StoreMatEntry(Trim(Me.lstMaterials.SelectedItem.SubItems(8)))
End Sub

Sub EnabledSubletMaterialObject(COND As Boolean)
    txtMatLineNo.Enabled = COND
    Combo1.Enabled = COND
    cboMatCode.Enabled = COND
    cboMaterial.Enabled = COND
    txtMatQty.Enabled = COND
    txtMatUnitPrice.Enabled = COND
    txtMatAmount.Enabled = COND
    cboMatChargeTo.Enabled = COND
    'optMatByPerc.Enabled = COND
    'txtMatDiscount.Enabled = COND
    'optMatByAmt.Enabled = COND
    'txtMatDiscountAmt.Enabled = COND
    txtdetail.Enabled = COND
    cmdMatDelete.Enabled = COND
End Sub

Private Sub lstParts_DblClick()
    If Pcnt = 0 Then Exit Sub
    
'    If txtInvoiceNo <> "" Then
'        MessagePop InfoFriend, "Repair order Information", "Repair Order Already Invoiced. Details Information cannot be edit", 1000
'        Exit Sub
'    End If
    
'    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
'        MessagePop InfoFriend, "Repair order Information", "Repair order details cannot be edit once its already invoice", 1000
'        Exit Sub
'    End If
    
    AddorEdit = "EDIT"
    Call SendToBack
    cmdAddParts.ZOrder 0: fraAddParts.ZOrder 0
    If txtInvoiceNo.Text = "" Then
        fraAddParts.Enabled = True
        'If CheckifDetailsIsSublet(lstParts.SelectedItem.SubItems(8)) = False Then
        '    Call EnabledSubletPartObject(True)
        'Else
        '    Call EnabledSubletPartObject(False)
            'MessagePop InfoFriend, "Info", "Sublet details cannot be edited here, please go to sublet purchasing for the editing of values"
        'End If
    Else
        fraAddParts.Enabled = False
    End If
    
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
        Picture1.Enabled = False: Picture5.Enabled = False: pic3.Enabled = False: Frame2.Enabled = False
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

    fraAddParts.Caption = "Edit Parts"
    Call StorePartsEntry(Trim(Me.lstParts.SelectedItem.SubItems(8)))
End Sub

Sub EnabledSubletPartObject(COND As Boolean)
    txtPartLabor.Enabled = COND
    cboPartNo.Enabled = COND
    cboDescription.Enabled = COND
    txtQty.Enabled = COND
    txtUnitPrice.Enabled = COND
    txtPartAmount.Enabled = COND
    'cboChargeTo.Enabled = COND
    'optPartsByPerc.Enabled = COND
    'txtPartDiscount.Enabled = COND
    'optPartsbyAmt.Enabled = COND
    'txtPartDiscountAmt.Enabled = COND
    cmdPartsDelete.Enabled = COND
End Sub

Private Sub lstAccessories_DblClick()
    If Acnt = 0 Then Exit Sub
    
'    If txtInvoiceNo <> "" Then
'        MessagePop InfoFriend, "Repair order Information", "Repair Order Already Invoiced. Details Information cannot be edit", 1000
'        Exit Sub
'    End If
    
'    If CheckIfRoIsAlreadyInvoice(txtRep_Or) = True Then
'        MessagePop InfoFriend, "Repair order Information", "Repair order details cannot be edit once its already invoice", 1000
'        Exit Sub
'    End If
    
    AddorEdit = "EDIT"
    Call SendToBack
    cmdAddAccessories.ZOrder 0: fraAddAccessories.ZOrder 0
    If txtInvoiceNo.Text = "" Then
        If CheckifDetailsIsSublet(lstAccessories.SelectedItem.SubItems(8)) = False Then
            fraAddAccessories.Enabled = True
        Else
            fraAddAccessories.Enabled = False
            MessagePop InfoFriend, "Info", "Sublet details cannot be edited here, please go to sublet purchasing for the editing of values"
        End If
    Else
        fraAddAccessories.Enabled = False
    End If
    
    'UPDATE BY   : MJP 08292008 †----------------------------------------------------------
    'DESCRIPTION : NEXT/PREVIOUS BUTTON CAN BE STILL CLICK WHEN ENTERING A INVOICE NO
        Picture1.Enabled = False: Picture5.Enabled = False: pic3.Enabled = False: Frame2.Enabled = False
    'UPDATE BY   : MJP 08292008 ----------------------------------------------------------†

    fraAddAccessories.Caption = "Edit Accessories"
    Call StoreAccEntry(Trim(Me.lstAccessories.SelectedItem.SubItems(8)))
End Sub

Private Sub lsvSearch_DblClick()
    If Not lsvSearch.ListItems.Count = 0 Then
        Call cmdSelect_Click
    End If
End Sub

Private Sub lsvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    labid.Caption = Item.ListSubItems(6).Text
End Sub

Private Sub lsvSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not lsvSearch.ListItems.Count = 0 Then
            Call cmdSelect_Click
        End If
    End If
End Sub

Private Sub optAccByAmt_Click()
    txtAccDiscount.Enabled = False
    txtAccDiscountAmt.Enabled = True
    txtAccDiscount.Text = 0
End Sub

Private Sub optAccByPerc_Click()
    txtAccDiscount.Enabled = True
    txtAccDiscountAmt.Enabled = False
    txtAccDiscountAmt.Text = ZERO
End Sub

Private Sub optByAmt_Click()
    txtJobDiscount.Enabled = False
    txtJobDiscountAmt.Enabled = True
    txtJobDiscount.Text = 0
End Sub

Private Sub optByCode_Click()
    cboJcode.Enabled = True
    DoEvents
    On Error Resume Next
    cboJcode.SetFocus
    cboJobCode.Enabled = False
End Sub

Private Sub optByDescription_Click()
    '    cboJobCode.Enabled = True
    '    DoEvents
    '    On Error Resume Next
    '    cboJobCode.SetFocus
    '    cboJcode.Enabled = False
End Sub

Private Sub optByPerc_Click()
    txtJobDiscount.Enabled = True
    txtJobDiscountAmt.Enabled = False
    txtJobDiscountAmt.Text = ZERO
End Sub

Private Sub Option1_Click()
    txtTerm.Text = "CSH"
    cmdOkBill.Enabled = True
End Sub

Private Sub Option2_Click()
'updated by:    IEBV_03282011AM
'description:   To check if customer had a terms for charge
'------------------------------------------------------------------------------------------------------------------------------------------------------
    If (gconDMIS.Execute("Select count(*) from csms_ro_det where rep_or = '" & RTrim(LTrim(txtRep_Or.Text)) & "' and isnull(wcode,'N') not in('C','S','W')").Fields(0).Value) >= 1 Then
        Dim rsHasCredit                              As New ADODB.Recordset
        Set rsHasCredit = New ADODB.Recordset
        Set rsHasCredit = gconDMIS.Execute("Select isnull(round(creditlimit,2),0) as creditlimit from ALL_Customer where cuscde = '" & Null2String(rsREPOR!ACCT_NO) & "'")
        If Not (rsHasCredit.EOF And rsHasCredit.BOF) Then
              If Null2String(rsHasCredit!CreditLimit) = 0 Then
                MsgBox "Credit is not yet configured.", vbInformation + vbOKOnly
                Option1.Value = True
                Exit Sub
              ElseIf (ROTotal - DiscTotal) > Null2String(rsHasCredit!CreditLimit) Then
                If MsgBox("Credit is over the limit, Do you want to continue?", vbQuestion + vbYesNo) = vbYes Then
                    fraBillBut.Enabled = False
                    picoverride.Visible = True
                    picoverride.ZOrder 0
                    ictr = 3
                    Exit Sub
                Else
                    cmdOkBill.Enabled = False
                    Option1.Value = True
                    'do nothing
                End If
              Else
                cmdOkBill.Enabled = True
                txtTerm.Text = "CHG"
              End If
        End If
    End If
End Sub

Private Sub optMatByAmt_Click()
    txtMatDiscount.Enabled = False
    txtMatDiscountAmt.Enabled = True
    txtMatDiscount.Text = 0
End Sub

Private Sub optMatByPerc_Click()
    txtMatDiscount.Enabled = True
    txtMatDiscountAmt.Enabled = False
    txtMatDiscountAmt.Text = ZERO
End Sub

Private Sub optPartsbyAmt_Click()
    txtPartDiscount.Enabled = False
    txtPartDiscountAmt.Enabled = True
    txtPartDiscount.Text = 0
End Sub

Private Sub optPartsByPerc_Click()
    txtPartDiscount.Enabled = True
    txtPartDiscountAmt.Enabled = False
    txtPartDiscountAmt.Text = ZERO
End Sub

Private Sub optSearch_Click(Index As Integer)
    txtSearch.SetFocus
End Sub

Private Sub rptRO_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Not Row.Record(6).Value = "" And Row.Record(8).Value = "Y" Then
        Metrics.ForeColor = &H8000&
    ElseIf Not Row.Record(2).Value = "" And Row.Record(8).Value = "Y" Then
        Metrics.ForeColor = &H800080
    ElseIf Row.Record(8).Value = "Y" Then
        Metrics.ForeColor = vbBlue
    ElseIf Row.Record(8).Value = "" And Row.Record(2).Value = "" And Row.Record(6).Value = "" Then
        Metrics.ForeColor = vbBlack
    End If
    
End Sub

Private Sub rptRO_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then: Exit Sub
    
    labid.Caption = Row.Record(7).Value                       'id
    Call rsRefresh
    Call StoreMemVars
    
    Call cmdSelect_Click
End Sub

'updated By:    IEBV 02222011_0525pm
'-----------------------------------------------------
Private Sub txtoverride_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ictr = 0 Then
        Ichg = False
        cmdBillBut.ZOrder 1
    Else
        If txtoverride.Text <> "ALONE" Then
            ictr = ictr - 1
            MsgBox "Invalid Code, You have ( " & ictr & " ) tries left!", vbInformation
            If ictr = 0 Then
                 cmdBillBut.ZOrder 1
                 Ichg = False
            Else
                Ichg = False
                txtoverride.Text = ""
                On Error Resume Next
                txtoverride.SetFocus
                Exit Sub
            End If
        Else
            txtTerm.Text = "CHG"
            txtoverride.Text = ""
            picoverride.Visible = False
            fraBillBut.Enabled = True
            cmdOkBill.Enabled = True
            Ichg = True
            Exit Sub
        End If
    End If
ElseIf KeyAscii = 27 Then
    txtoverride.Text = ""
    picoverride.Visible = False
    fraBillBut.Enabled = True
    Option1.Value = True
    Exit Sub
End If
'-----------------------------------------------------
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If N2Str2Zero(rsREPOR!VAT_EXEMPT) = 1 Then
        If labZeroRated.Visible = True Then
            labZeroRated.Visible = False
        Else
            labZeroRated.Visible = True
        End If
    Else
        labZeroRated.Visible = False
    End If
End Sub

Private Sub txtAccDiscountAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("0123456789.", KeyAscii)
    End If
End Sub

Private Sub txtAcct_No_KeyPress(KeyAscii As Integer)
'    If Trim(txtAcct_No.Text) = "" Then
'        RO_OR_ESTI_OR_PART = "CUST"
'        frmCSMSRO_CustomerSearch.Show 1
'    End If
End Sub

Private Sub txtAccUnitPrice_GotFocus()
    txtAccUnitPrice.Text = NumericVal(txtAccUnitPrice.Text)
End Sub

Private Sub txtCertific8_Click()
    If txtPlate_No.Text <> "" Then
        Me.Enabled = False
        frmCSMSROCusveh.Show
        frmCSMSROCusveh.ZOrder 0
    Else
        MsgSpeechBox "Plate Number must be inputed! Please enter 000000 if unknown"
    End If
End Sub

Private Sub txtDateReleased_LostFocus()
    If txtDateReleased.Text <> "" Then txtDateReleased.Text = Format(txtDateReleased.Text, "Short Date")
End Sub

Private Sub txtDte_comp_LostFocus()
    If txtDte_comp.Text <> "" Then txtDte_comp.Text = Format(txtDte_comp.Text, "Short Date")
End Sub

Private Sub txtDte_recd_LostFocus()
    If txtDte_recd.Value <> "" Then txtDte_recd.Value = Format(txtDte_recd.Value, "Short Date")
End Sub

Private Sub txtDte_Rel_LostFocus()
    If txtDte_Rel.Text <> "" Then txtDte_Rel.Text = Format(txtDte_Rel.Text, "Short Date")
End Sub

Private Sub txtInvoiceDate_LostFocus()
    If txtInvoiceDate.Text <> "" Then txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "Short Date")
End Sub

Private Sub txtInvoiceNumber_LostFocus()
    If NumericVal(txtInvoiceNumber.Text) > 0 Then
        txtInvoiceNumber.Text = Format(txtInvoiceNumber.Text, "000000")
        cmdOkBill.Enabled = True
    Else
        If Len(txtInvoiceNumber.Text) = 6 Then

            cmdOkBill.Enabled = True
        Else
            cmdOkBill.Enabled = False
        End If
    End If
'updated By:    IEBV 02222011_0525pm
'-----------------------------------------------------
    If Option2.Value = True Then
            If Ichg = False Then
                Call Option2_Click
            Else
                'do nothing
            End If
    End If
'-----------------------------------------------------
End Sub

Private Sub txtJobDiscountAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("0123456789.", KeyAscii)
    End If
End Sub

Private Sub txtJobRate_LostFocus()
    txtJobRate.Text = Format(txtJobRate.Text, MAXIMUM_DIGIT)
    If NumericVal(txtJobRate) < 0 Then
        txtJobRate = "0.00"
    End If
End Sub

Private Sub txtKm_rdg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tabs}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtLOAAmount_Change()
    Dim RO_JOB                                      As Double
    Dim RO_PARTS                                    As Double
    Dim RO_MATS                                     As Double
    Dim RO_ACCS                                     As Double
    
    
    RO_JOB = TOTJOBAMT + JobInsTotal
    RO_PARTS = TOTPARTSAMT + PartsInsTotal
    RO_MATS = TOTMATAMT + MatInsTotal
    RO_ACCS = TOTACCAMT + AccInsTotal

    If chkAllowManDist.Value = 0 Then
        If NumericVal(txtLOAAmount.Text) > NumericVal((RO_JOB - TOTJOBDISC) + (RO_PARTS - TOTPARTSDISC) + (RO_MATS - TOTMATDISC) + (RO_ACCS - TOTACCDISC)) Then
            MsgBox "Warning: LOA Amount should not Exceed Repair Order Total Amount.", vbCritical, "Not Allowed!"
            txtLOAAmount.Text = NumericVal(txtPartTotal.Text)
            Exit Sub
        End If

        If NumericVal(txtLOAAmount.Text) > (RO_JOB - TOTJOBDISC) Then
            txtPartLabor.Text = RO_JOB - TOTJOBDISC
            If NumericVal(txtLOAAmount.Text) - (RO_JOB - TOTJOBDISC) > (RO_PARTS - TOTPARTSDISC) Then
                txtPartParts.Text = RO_PARTS - TOTPARTSDISC
                If NumericVal(txtLOAAmount.Text) - (RO_JOB - TOTJOBDISC) - (RO_PARTS - TOTPARTSDISC) > (RO_MATS - TOTMATDISC) Then
                    txtPartMaterials.Text = RO_MATS - TOTMATDISC
                Else
                    txtPartMaterials.Text = NumericVal(txtLOAAmount.Text) - (RO_JOB - TOTJOBDISC) - (RO_PARTS - TOTPARTSDISC)
                    If NumericVal(txtLOAAmount.Text) - (RO_JOB - TOTJOBDISC) - (RO_PARTS - TOTPARTSDISC) - (RO_MATS - TOTMATDISC) > (RO_ACCS - TOTACCDISC) Then
                        txtPartAccessories.Text = NumericVal(txtLOAAmount.Text) - (RO_JOB - TOTJOBDISC) - (RO_PARTS - TOTPARTSDISC) - (RO_MATS - TOTMATDISC) - (RO_ACCS - TOTACCDISC)
                    Else
                        txtPartAccessories.Text = RO_ACCS - TOTACCDISC
                    End If
                End If
            Else
                txtPartParts.Text = NumericVal(txtLOAAmount.Text) - (RO_JOB - TOTJOBDISC)
            End If
        Else
            txtPartLabor.Text = NumericVal(txtLOAAmount.Text)
            txtPartParts.Text = ZERO
            txtPartMaterials.Text = ZERO
            txtPartAccessories.Text = ZERO
        End If
    End If
End Sub

Private Sub txtLOAAmount_GotFocus()
    txtLOAAmount.Text = NumericVal(txtLOAAmount.Text)
End Sub

Private Sub txtLOAAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtMatAmount_GotFocus()
    txtMatAmount.Text = ToDoubleNumber(NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text))
End Sub

Private Sub txtMatAmount_LostFocus()
    txtMatAmount.Text = Format(NumericVal(txtMatAmount.Text), MAXIMUM_DIGIT)
End Sub

Private Sub txtMatDiscount_LostFocus()
    txtMatDiscount.Text = Format(txtMatDiscount.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtMatDiscountAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("0123456789.", KeyAscii)
    End If
End Sub

Private Sub txtMatQty_Change()
    If NumericVal(txtMatQty.Text) > 0 Then
        txtMatAmount.Text = ToDoubleNumber(NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text))
    End If
End Sub

Private Sub txtMatQty_LostFocus()
    If NumericVal(txtMatQty.Text) > 0 Then
        txtMatAmount.Text = Format(NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text), "#####0.0")
    End If
End Sub

Private Sub txtMatUnitPrice_Change()
    If NumericVal(txtMatUnitPrice.Text) > 0 Then
        txtMatAmount.Text = ToDoubleNumber(NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text))
    End If
End Sub

Private Sub txtMatUnitPrice_GotFocus()
    txtMatUnitPrice.Text = NumericVal(txtMatUnitPrice.Text)
End Sub

Private Sub txtMatUnitPrice_LostFocus()
    If NumericVal(txtMatUnitPrice.Text) > 0 Then
        txtMatAmount.Text = Format(NumericVal(txtMatQty.Text) * NumericVal(txtMatUnitPrice.Text), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtPartAccessories_Change()
    If chkAllowManDist.Value = 1 Then
        If NumericVal(txtPartAccessories.Text) > (TOTACCAMT - TOTACCDISC) + AccInsTotal Then
            MsgBox "Warning: System Doesn't allow Participation to Exceed Actual Accessories Amount" & vbCrLf & "                Actual Accessories Amount will be set as default", vbCritical, "Not Allowed!"
            txtPartAccessories.Text = (TOTACCAMT - TOTACCDISC) + AccInsTotal
        End If
    End If
    Call SetTotalParticipation
End Sub

Private Sub txtPartAccessories_GotFocus()
    If NumericVal(txtPartAccessories) > 0 Then
        txtPartAccessories = NumericVal(txtPartAccessories)
    Else
        txtPartAccessories = ""
    End If
End Sub

Private Sub txtPartAccessories_LostFocus()
    If NumericVal(txtPartAccessories) > 0 Then
        txtPartAccessories = ToDoubleNumber(txtPartAccessories)
    Else
        txtPartAccessories = "0.00"
    End If
End Sub

Private Sub txtPartAmount_GotFocus()
    txtPartAmount.Text = ToDoubleNumber(NumericVal(txtQty.Text) * NumericVal(txtUnitPrice.Text))
End Sub

Private Sub cboPartno_LostFocus()
    If cboPartNo.Text <> "" Then cboDescription.Text = SetPartDisc(cboPartNo.Text)
    txtUnitPrice.Text = SetPartPrice(cboPartNo.Text)
    txtPartAmount.Text = ToDoubleNumber(NumericVal(txtQty.Text) * NumericVal(txtUnitPrice.Text))
End Sub

Private Sub txtAccAmount_GotFocus()
    txtAccAmount.Text = ToDoubleNumber(NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text))
End Sub

Private Sub txtAccAmount_LostFocus()
    txtAccAmount.Text = Format(NumericVal(txtAccAmount.Text), MAXIMUM_DIGIT)
End Sub

Private Sub txtAccDiscount_LostFocus()
    txtAccDiscount.Text = Format(txtAccDiscount.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtAccQty_Change()
    If NumericVal(txtAccQty.Text) > 0 Then
        txtAccAmount.Text = ToDoubleNumber(NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text))
    End If
End Sub

Private Sub txtAccQty_LostFocus()
    If NumericVal(txtAccQty.Text) > 0 Then
        txtAccAmount.Text = Format(NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text), "#####0.0")
    End If
End Sub

Private Sub txtAccUnitPrice_Change()
    If NumericVal(txtAccUnitPrice.Text) > 0 Then
        txtAccAmount.Text = ToDoubleNumber(NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text))
    End If
End Sub

Private Sub txtAccUnitPrice_LostFocus()
    If NumericVal(txtAccUnitPrice.Text) > 0 Then
        txtAccAmount.Text = Format(NumericVal(txtAccQty.Text) * NumericVal(txtAccUnitPrice.Text), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtPartAmount_LostFocus()
    txtPartAmount.Text = Format(txtPartAmount.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtPartDiscount_LostFocus()
    txtPartDiscount.Text = Format(txtPartDiscount.Text, DIGIT_FORMAT)
End Sub

Private Sub txtPartDiscountAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("0123456789.", KeyAscii)
    End If
End Sub

Private Sub txtParticipat_LostFocus()
    '    If txtParticipat.Text <> "" And chkParticipat.Value = 1 Then
    '        txtParticipat.Text = UCase(txtParticipat.Text)
    '        Set rsCustomer = New ADODB.Recordset
    '        Set rsCustomer = gconDMIS.Execute("select cuscde,cusnam,cusadd from All_Cusmas where cuscde = '" & txtAcct_No.Text & "'")
    '        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
    '            txtNiym.Text = Null2String(rsCustomer!cusnam)
    '            txtAddress.Text = Null2String(rsCustomer!Cusadd)
    '        End If
    '        Set rsCustomer = New ADODB.Recordset
    '        Set rsCustomer = gconDMIS.Execute("select cuscde,cusnam from All_Cusmas where cuscde = '" & txtParticipat.Text & "'")
    '        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
    '            txtNiym.Text = txtNiym.Text & "/" & Null2String(rsCustomer!cusnam)
    '        End If
    '        Set rsCustomer = Nothing
    '    End If
End Sub

Private Sub txtPartLabor_Change()
    If chkAllowManDist.Value = 1 Then
        If NumericVal(txtPartLabor.Text) > (TOTJOBAMT - TOTJOBDISC) + JobInsTotal Then
            MsgBox "Warning: System Doesn't allow Participation to Exceed Actual Job Amount" & vbCrLf & "                Actual Job Amount will be set as default", vbCritical, "Not Allowed!"
            txtPartLabor.Text = (TOTJOBAMT - TOTJOBDISC) + JobInsTotal
        End If
    End If
    Call SetTotalParticipation
End Sub

Private Sub txtPartLabor_GotFocus()
    If NumericVal(txtPartLabor) > 0 Then
        txtPartLabor = NumericVal(txtPartLabor)
    Else
        txtPartLabor = ""
    End If
End Sub

Private Sub txtPartLabor_LostFocus()
    If NumericVal(txtPartLabor) > 0 Then
        txtPartLabor = ToDoubleNumber(txtPartLabor)
    Else
        txtPartLabor = "0.00"
    End If
End Sub

Private Sub txtPartMaterials_Change()
    If chkAllowManDist.Value = 1 Then
        If NumericVal(txtPartMaterials.Text) > (TOTMATAMT - TOTMATDISC) + MatInsTotal Then
            MsgBox "Warning: System Doesn't allow Participation to Exceed Actual Materials Amount" & vbCrLf & "                Actual Materials Amount will be set as default", vbCritical, "Not Allowed!"
            txtPartMaterials.Text = (TOTMATAMT - TOTMATDISC) + MatInsTotal
        End If
    End If
    Call SetTotalParticipation
End Sub

Private Sub txtPartMaterials_GotFocus()
    If NumericVal(txtPartMaterials) > 0 Then
        txtPartMaterials = NumericVal(txtPartMaterials)
    Else
        txtPartMaterials = ""
    End If
End Sub

Private Sub txtPartMaterials_LostFocus()
    If NumericVal(txtPartMaterials) > 0 Then
        txtPartMaterials = ToDoubleNumber(txtPartMaterials)
    Else
        txtPartMaterials = "0.00"
    End If
End Sub

Private Sub txtPartParts_Change()
    If chkAllowManDist.Value = 1 Then
        If NumericVal(txtPartParts.Text) > (TOTPARTSAMT - TOTPARTSDISC) + PartsInsTotal Then
            MsgBox "Warning: System Doesn't allow Participation to Exceed Actual Parts Amount" & vbCrLf & "                Actual Parts Amount will be set as default", vbCritical, "Not Allowed!"
            txtPartParts.Text = (TOTPARTSAMT - TOTPARTSDISC) + PartsInsTotal
        End If
    End If
    Call SetTotalParticipation
End Sub

Private Sub txtPartParts_GotFocus()
    If NumericVal(txtPartParts) > 0 Then
        txtPartParts = NumericVal(txtPartParts)
    Else
        txtPartParts = ""
    End If
End Sub

Private Sub txtPartParts_LostFocus()
    If NumericVal(txtPartParts) > 0 Then
        txtPartParts = ToDoubleNumber(txtPartParts)
    Else
        txtPartParts = "0.00"
    End If
End Sub

Private Sub txtQTY_Change()
    If txtQty.Text <> "" Then
        txtPartAmount.Text = ToDoubleNumber(NumericVal(txtQty.Text) * NumericVal(txtUnitPrice.Text))
    End If
End Sub

Private Sub txtQty_LostFocus()
    txtQty.Text = Format(txtQty.Text, DIGIT_FORMAT)
End Sub

Private Sub txtReleaseDate_LostFocus()
    If IsDate(txtReleaseDate) = True Then
        txtReleaseDate.Text = Format(txtReleaseDate.Text, "Short Date")
    Else
        txtReleaseDate = ""
    End If
End Sub

Private Sub txtRep_Or_LostFocus()
    Dim Rep_Or2, rep_or3                               As String
    Dim k                                              As Integer
    If Left(txtRep_Or.Text, 2) = "R-" Then
        txtRep_Or.Text = "R-" & Format(NumericVal(Right(txtRep_Or.Text, Len(txtRep_Or.Text) - 2)), "00000000")
    Else
        txtRep_Or.Text = "R-" & Format(NumericVal(Right(txtRep_Or.Text, Len(txtRep_Or.Text))), "00000000")
    End If
    If AddorEdit = "ADD" Then
        Dim rsReporDup                                 As New ADODB.Recordset
        Set rsReporDup = gconDMIS.Execute("select rep_or from CSMS_RepOr where rep_or = " & N2Str2Null(txtRep_Or.Text))
        If Not rsReporDup.EOF And Not rsReporDup.BOF Then
            MsgSpeechBox "Warning: Repair Order Number Already Exist!"
            On Error Resume Next
            txtRep_Or.SetFocus
        End If
        Set rsReporDup = Nothing
    End If
End Sub

Private Sub txtSearch_Change()
    'rptRO.FilterText = txtSearch.Text
    'rptRO.Populate
    
    Call FillSearchGrid(txtSearch)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSUPLOAD                                        As New ADODB.Recordset
    Dim REC                                             As XtremeReportControl.ReportRecord
    Dim FIELD                                           As String
    XXX = Replace(XXX, "'", "")
    If XXX = "" Then
        Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,NULL AS DONE  FROM CSMS_REPOR WHERE DTE_COMP IS NULL AND TRANSTYPE = 'R' ORDER BY REP_OR DESC")
    Else
        'customer
        If optSearch(0).Value = True Then
            FIELD = "A.NIYM"
            If CHECKIFFINISHEDJOB(XXX, FIELD) = True Then
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,'Y' AS DONE  FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND NIYM LIKE '%" & XXX & "%' ORDER BY NIYM ASC")
            Else
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL, NULL AS DONE  FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND NIYM LIKE '%" & XXX & "%' ORDER BY NIYM ASC")
            End If
        'repairorder
        ElseIf optSearch(1).Value = True Then
             FIELD = "A.REP_OR"
             If CHECKIFFINISHEDJOB(XXX, FIELD) = True Then
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,'Y' AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND REP_OR LIKE '%" & XXX & "%' ORDER BY REP_OR DESC")
             Else
                 Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,NULL AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND REP_OR LIKE '%" & XXX & "%' ORDER BY REP_OR DESC")
             End If
        'invoice
        ElseIf optSearch(2).Value = True Then
             FIELD = "A.INVOICE"
            If CHECKIFFINISHEDJOB(XXX, FIELD) = True Then
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,'Y' AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND INVOICE LIKE '%" & XXX & "%' ORDER BY INVOICE ASC")
            Else
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,NULL AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND INVOICE LIKE '%" & XXX & "%' ORDER BY INVOICE ASC")
            End If
        'plate number
        ElseIf optSearch(3).Value = True Then
             FIELD = "A.PLATE_NO"
            If CHECKIFFINISHEDJOB(XXX, FIELD) = True Then
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,'Y' AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND PLATE_NO LIKE '%" & XXX & "%' ORDER BY PLATE_NO ASC")
            Else
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,NULL AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND PLATE_NO LIKE '%" & XXX & "%' ORDER BY PLATE_NO ASC")
            End If
        'vin number
        ElseIf optSearch(4).Value = True Then
             FIELD = "A.VIN"
             If CHECKIFFINISHEDJOB(XXX, FIELD) = True Then
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,'Y' AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND VIN LIKE '%" & XXX & "%' ORDER BY VIN ASC")
             Else
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,NULL AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND VIN LIKE '%" & XXX & "%' ORDER BY VIN ASC")
             End If
        'MODEL
        Else
             FIELD = "A.MODEL"
            If CHECKIFFINISHEDJOB(XXX, FIELD) = True Then
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL, 'Y' AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND MODEL LIKE '%" & XXX & "%' ORDER BY MODEL ASC")
            Else
                Set RSUPLOAD = gconDMIS.Execute("SELECT TOP 100 REP_OR, INVOICE, NIYM, PLATE_NO, VIN, MODEL, ID, DTE_REL,NULL AS DONE FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND MODEL LIKE '%" & XXX & "%' ORDER BY MODEL ASC")
            End If
        End If
    End If
    
    rptRO.Records.DeleteAll
    While Not RSUPLOAD.EOF
        Set REC = rptRO.Records.Add
        REC.AddItem (Trim(RSUPLOAD!NIYM))
        REC.AddItem (Trim(RSUPLOAD!REP_OR))
        REC.AddItem (Trim(RSUPLOAD!invoice))
        REC.AddItem (Trim(RSUPLOAD!PLATE_NO))
        REC.AddItem (Trim(RSUPLOAD!Vin))
        REC.AddItem (Trim(RSUPLOAD!Model))
        REC.AddItem (Trim(RSUPLOAD!DTE_rel))
        REC.AddItem (Trim(RSUPLOAD!ID))
        REC.AddItem (Trim(RSUPLOAD!DONE))
        RSUPLOAD.MoveNext
        Set REC = Nothing
    Wend
    rptRO.Populate
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFFF
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

Private Sub txtUnitPrice_Change()
    If txtUnitPrice.Text <> "" Then
        txtPartAmount.Text = ToDoubleNumber(NumericVal(txtQty.Text) * NumericVal(txtUnitPrice.Text))
    End If
End Sub

Private Sub txtUnitPrice_LostFocus()
    txtUnitPrice.Text = Format(txtUnitPrice.Text, MAXIMUM_DIGIT)
End Sub

Sub InitializeRC()
    With rptRO
        .Columns.DeleteAll
        .Columns.Add 0, "Customer Name", 300, True::    .Columns(0).Alignment = xtpAlignmentLeft:       .Columns(0).AllowRemove = False
        .Columns.Add 1, "Repair Order no", 100, True:   .Columns(1).Alignment = xtpAlignmentLeft:       .Columns(1).AllowRemove = False
        .Columns.Add 2, "Invoice No", 80, True:         .Columns(2).Alignment = xtpAlignmentLeft:       .Columns(2).AllowRemove = False
        .Columns.Add 3, "Plate no", 60, True:           .Columns(3).Alignment = xtpAlignmentLeft:       .Columns(3).AllowRemove = False
        .Columns.Add 4, "Vin no", 120, True:            .Columns(4).Alignment = xtpAlignmentLeft:       .Columns(4).AllowRemove = False
        .Columns.Add 5, "Vehicle Model", 180, True:     .Columns(5).Alignment = xtpAlignmentLeft:       .Columns(5).AllowRemove = False
        .Columns.Add 6, "Date Released", 95, True:      .Columns(6).Alignment = xtpAlignmentLeft:       .Columns(6).AllowRemove = False:    .Columns(6).Resizable = False
        .Columns.Add 7, "id", 0, True:                  .Columns(7).Alignment = xtpAlignmentLeft:       .Columns(7).AllowRemove = False:    .Columns(7).Resizable = False
        .Columns.Add 8, "done", 0, True:                .Columns(8).Alignment = xtpAlignmentLeft:       .Columns(8).AllowRemove = False:    .Columns(8).Resizable = False
        
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
    End With
End Sub
'Updated: IEBV 06282010 1200PM
'Description: Add RoType
Function ADDROTYPE()
    Cbo_Rotype.Clear
    Cbo_Rotype.AddItem ""
    Cbo_Rotype.AddItem "WTY"
    Cbo_Rotype.AddItem "BRP"
    Cbo_Rotype.AddItem "JET"
    Cbo_Rotype.AddItem "GJ"
    Cbo_Rotype.AddItem "O/H"
    Cbo_Rotype.AddItem "PP"
    Cbo_Rotype.AddItem "RF"
    Cbo_Rotype.AddItem "QS"
    Cbo_Rotype.AddItem "AC"
    Cbo_Rotype.AddItem "PDI"
    Cbo_Rotype.AddItem "QC"
    Cbo_Rotype.AddItem "FI"
    Cbo_Rotype.AddItem "DET"
'Updated: IEBV IEBV 06282010 1200PM
'Description: Add RoType
End Function
'Updated: IEBV IEBV 06282010 1200PM
'Description: Shows The Ro type of the specified Ro number
Function ShorROtype(RONO As String)
    Dim rsROTYPE            As New ADODB.Recordset
    Set rsROTYPE = gconDMIS.Execute("Select ROTYPE from csms_repor where Rep_or = '" & RONO & "'")
    If (rsROTYPE!ROTYPE) <> "" Then
        txttype.Text = N2String(rsROTYPE!ROTYPE)
    Else
        txttype.Text = ""
    End If
End Function
'Updated: IEBV IEBV 06282010 1200PM
'Description: Shows The Ro type of the specified Ro number

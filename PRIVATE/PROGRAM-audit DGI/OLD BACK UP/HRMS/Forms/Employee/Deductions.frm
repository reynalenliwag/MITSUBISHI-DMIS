VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMSDeductions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deductions Entry"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8595
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Deductions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   8595
   Begin VB.PictureBox picDeductions 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   3210
      ScaleHeight     =   4215
      ScaleWidth      =   4035
      TabIndex        =   12
      Top             =   1020
      Width           =   4065
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   2010
         TabIndex        =   28
         Top             =   3600
         Width           =   1785
      End
      Begin VB.TextBox txtNoOfMinutes 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3180
         Width           =   1785
      End
      Begin VB.TextBox txtNoMin_Computed 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2010
         TabIndex        =   26
         Top             =   2820
         Width           =   1785
      End
      Begin VB.TextBox txtNoHours_Computed 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2010
         TabIndex        =   24
         Top             =   2430
         Width           =   1785
      End
      Begin VB.ComboBox cboParticular 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1980
         Width           =   3555
      End
      Begin VB.ComboBox cboQuensina 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   270
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   19
         Text            =   "cboQuensina"
         Top             =   660
         Width           =   3195
      End
      Begin MSComCtl2.DTPicker dt_deduct 
         Height          =   345
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54460417
         CurrentDate     =   40179
      End
      Begin VB.Label Label8 
         Caption         =   "Cut-Off"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   45
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label Label7 
         Caption         =   "Date"
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
         Left            =   240
         TabIndex        =   44
         Top             =   1050
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   4095
         _Version        =   655364
         _ExtentX        =   7223
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Add/Edit Deductions"
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
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Hours"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1035
         TabIndex        =   15
         Top             =   2550
         Width           =   930
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1230
         TabIndex        =   16
         Top             =   2940
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Computed Minutes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   17
         Top             =   3300
         Width           =   1770
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         TabIndex        =   14
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1140
         TabIndex        =   18
         Top             =   3690
         Width           =   720
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         Caption         =   "ID"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   -210
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin wizButton.cmd cmdDeductions 
      Height          =   4335
      Left            =   3150
      TabIndex        =   11
      Top             =   990
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   7646
      TX              =   "cmd1"
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
      MICON           =   "Deductions.frx":030A
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2865
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   21
      Top             =   5505
      Width           =   5580
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
         Left            =   4860
         MouseIcon       =   "Deductions.frx":0326
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":0478
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   4170
         MouseIcon       =   "Deductions.frx":07DE
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":0930
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Print this Record"
         Top             =   30
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
         Left            =   3480
         MouseIcon       =   "Deductions.frx":0C96
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":0DE8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
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
         Left            =   2790
         MouseIcon       =   "Deductions.frx":1113
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":1265
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Left            =   2100
         MouseIcon       =   "Deductions.frx":15C1
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":1713
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Add Record"
         Top             =   30
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
         Left            =   1410
         MouseIcon       =   "Deductions.frx":1A26
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":1B78
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Find a Record"
         Top             =   30
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
         Left            =   720
         MouseIcon       =   "Deductions.frx":1E72
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":1FC4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move to Next Record"
         Top             =   30
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
         Left            =   30
         MouseIcon       =   "Deductions.frx":231C
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":246E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picSearch 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6465
      Left            =   0
      Picture         =   "Deductions.frx":27CD
      ScaleHeight     =   6435
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   0
      Width           =   2505
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
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
         Left            =   30
         MaxLength       =   35
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   60
         Width           =   2415
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   5775
         Left            =   30
         TabIndex        =   2
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   10186
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
         BorderStyle     =   1
         Appearance      =   0
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
         MouseIcon       =   "Deductions.frx":5509
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture11 
      Height          =   6405
      Left            =   0
      Picture         =   "Deductions.frx":566B
      ScaleHeight     =   6345
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   0
      Width           =   2505
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1080
      Left            =   2580
      ScaleHeight     =   1080
      ScaleWidth      =   8460
      TabIndex        =   7
      Top             =   0
      Width           =   8460
      Begin VB.TextBox txtEmpno 
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
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   60
         Width           =   1785
      End
      Begin VB.TextBox txtPosition 
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
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   660
         Width           =   5685
      End
      Begin VB.TextBox txtName 
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
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   60
         Width           =   3855
      End
      Begin Crystal.CrystalReport rptDeductions 
         Left            =   5280
         Top             =   1140
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
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   30
         TabIndex        =   8
         Top             =   450
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   4425
      Left            =   2520
      ScaleHeight     =   4425
      ScaleWidth      =   5955
      TabIndex        =   9
      Top             =   1080
      Width           =   5955
      Begin VB.TextBox txtYTDAbsent 
         Alignment       =   1  'Right Justify
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   4020
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtYTDUTLate 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   4020
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid grdDeductions 
         Height          =   3975
         Left            =   30
         TabIndex        =   10
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   6
         ForeColor       =   0
         ForeColorFixed  =   0
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Deductions.frx":193C8
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Absent"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3060
         TabIndex        =   41
         Top             =   4095
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Late"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   60
         TabIndex        =   40
         Top             =   4095
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7005
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   35
      Top             =   5505
      Width           =   1440
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
         Left            =   720
         MouseIcon       =   "Deductions.frx":196E2
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":19834
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Left            =   30
         MouseIcon       =   "Deductions.frx":19B72
         MousePointer    =   99  'Custom
         Picture         =   "Deductions.frx":19CC4
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label LABMONTH 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   3900
      TabIndex        =   43
      Top             =   6810
      Width           =   3825
   End
End
Attribute VB_Name = "frmHRMSDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo                                                         As ADODB.Recordset
Dim rsDeductions                                                      As ADODB.Recordset
Dim rsSalaryGrade                                                     As ADODB.Recordset
Dim rsSETUPDEDUCTION                                                  As ADODB.Recordset
Dim rsfindDup                                                         As ADODB.Recordset
Dim DAYS_OF_WORK                                                      As Integer
Dim ADDOREDIT                                                         As String
Attribute ADDOREDIT.VB_VarUserMemId = 1073938436
Dim Diyt                                                              As String
Dim EMPLIVIL                                                          As String
Dim xCUT_OFF                                                          As String
Attribute xCUT_OFF.VB_VarUserMemId = 1073938438
Dim xPAY_MONTH                                                        As Integer
Dim xPAY_YEAR                                                         As Integer

'Sub AddMonthName()
'    Dim X As Integer
'    cboMOnth.Clear
'    For X = 1 To 12
'        cboMOnth.AddItem MonthName(X)
'    Next
'End Sub

Function GetBasicPay()

    Dim RSTMP                                                         As ADODB.Recordset
    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT BASICSALARY, EMPSTATUS FROM HRMS_EMPINFO WHERE EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!EMPSTATUS) = "M" Then
            GetBasicPay = N2Str2Zero(RSTMP!BASICSALARY)
        ElseIf Null2String(RSTMP!EMPSTATUS) = "D" Then
            GetBasicPay = (N2Str2Zero(RSTMP!BASICSALARY) * DAYS_OF_WORK) / 12
        End If
    End If
    Set RSTMP = Nothing
End Function

Function StoreEntry(ByVal ID As Variant)
    Dim MM                                                            As String
    Dim YY                                                            As String
    Dim TheDeyt                                                       As String
    Dim SPACES                                                        As String
    Dim ENTRYBY                                                       As String
    Dim totmin
    Dim RSTMP                                                         As New ADODB.Recordset

    Set rsDeductions = New ADODB.Recordset
    rsDeductions.Open "SELECT * FROM HRMS_DEDUCTIONS WHERE ID = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsDeductions.EOF And Not rsDeductions.BOF Then
        SPACES = "                    "
        LABID.Caption = rsDeductions!ID
        TheDeyt = Null2Date(rsDeductions!DEYT)
        totmin = N2Str2IntZero(rsDeductions!nomin)
        txtNoOfMinutes = totmin
        txtNoHours_Computed = totmin \ 60
        txtNoMin_Computed = totmin Mod 60
        txtAmount.Text = N2Str2Zero(rsDeductions!AMOUNT)
        
        'cboDay.Text = Day(TheDeyt)
        'cboMonth.Text = The_month(MONTH(TheDeyt))
        
        dt_deduct = Null2Date(rsDeductions!DEYT)
        
        Set RSTMP = gconDMIS.Execute("SELECT DESCRIPTION,ENTRYBY FROM HRMS_DEDUCTIONCODE WHERE CODE = '" & rsDeductions!PARTICULAR & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            If Null2String(RSTMP!ENTRYBY) = "TIME" Then ENTRYBY = "T"
            If Null2String(RSTMP!ENTRYBY) = "AMOUNT" Then ENTRYBY = "A"
            If Null2String(RSTMP!ENTRYBY) = "SALARY" Then ENTRYBY = "S"
            cboParticular.Text = Null2String(RSTMP!Description) & SPACES & " - " & Null2String(rsDeductions!PARTICULAR) & "-" & ENTRYBY
        End If
        Set RSTMP = Nothing
    End If
End Function

Function GetDeductionDescription(DEDID As String)
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DESCRIPTION FROM HRMS_DEDUCTIONCODE WHERE CODE= '" & DEDID & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetDeductionDescription = Null2String(RSTMP!Description)
    Else
        GetDeductionDescription = ""
    End If
    Set RSTMP = Nothing
End Function

Sub EnablePics(COND As Boolean)
    picSearch.Enabled = COND
    Picture4.Enabled = COND
End Sub

Sub EnableDisabledTimeAndAmount()
    Dim ENTRYBY                                                       As String
    Dim rperday                                                       As Double
    Dim rperHour                                                      As Double
    Dim rperMinute                                                    As Double
    ENTRYBY = Right(cboParticular, 1)
    If ENTRYBY = "T" Then
        txtNoMin_Computed.Enabled = True
        If UCase(Mid(Right(cboParticular, 4), 1, 2)) = "WD" Or UCase(Mid(Right(cboParticular, 4), 1, 2)) = "HD" Then
            txtNoMin_Computed.Enabled = False
        End If
        txtNoOfMinutes.Enabled = False
        txtNoHours_Computed.Enabled = True
        txtAmount.Enabled = False
        Label4.Caption = "Amount"
    ElseIf ENTRYBY = "S" Then
        rperday = Round(((GetBasicPay * 12) / DAYS_OF_WORK), 2)
        rperHour = Round(rperday / 8, 2)
        rperMinute = Round(rperHour / 60, 2)
        txtNoOfMinutes.Enabled = False
        txtNoHours_Computed.Enabled = False
        txtNoMin_Computed.Enabled = False
        txtAmount.Enabled = True
        If UCase(Mid(Right(cboParticular, 4), 1, 2)) = "WD" Then
            txtAmount.Text = Format(rperday, "#,###,##0.00")
            Label4.Caption = "Rate/Day"
        ElseIf UCase(Mid(Right(cboParticular, 4), 1, 2)) = "HD" Then
            txtAmount.Text = Format((rperday / 2), "#,###,##0.00")
            Label4.Caption = "Rate/Half Day"
        End If
    Else
        txtAmount.Enabled = True
        Label4.Caption = "Amount"
        txtNoOfMinutes.Enabled = False
        txtNoHours_Computed.Enabled = False
        txtNoMin_Computed.Enabled = False
    End If
End Sub

Sub rsSETUP()
    Set rsSETUPDEDUCTION = gconDMIS.Execute("SELECT WORKING_DAY FROM HRMS_SETUPDEDUCTION")
    If Not (rsSETUPDEDUCTION.EOF And rsSETUPDEDUCTION.BOF) Then
        DAYS_OF_WORK = N2Str2Zero(rsSETUPDEDUCTION!WORKING_DAY)
    End If
    Set rsSETUPDEDUCTION = Nothing
End Sub

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT EMPNO,[POSITION],LASTNAME,FIRSTNAME,MIDDLENAME,EMPLEVEL,SALARYCODE FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & EMPINFOEMPNO.Caption & "'", gconDMIS
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT EMPNO,[POSITION],LASTNAME,FIRSTNAME,MIDDLENAME,EMPLEVEL,SALARYCODE FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & frmHRMSEmpInfo.LABID.Caption & "'", gconDMIS
    Else
        Set rsEmpInfo = New ADODB.Recordset
        'rsEmpInfo.Open "SELECT EMPNO,[POSITION],LASTNAME,FIRSTNAME,MIDDLENAME,EMPLEVEL,RESIGNED,SALARYCODE FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL ORDER BY LASTNAME,FIRSTNAME,MIDDLENAME ASC", gconDMIS
        rsEmpInfo.Open "SELECT EMPNO,[POSITION],LASTNAME,FIRSTNAME,MIDDLENAME,EMPLEVEL,RESIGNED,SALARYCODE FROM HRMS_EMPINFO WHERE RESIGNED IS NULL ORDER BY LASTNAME,FIRSTNAME,MIDDLENAME ASC", gconDMIS
    End If
End Sub

Sub InitGrid()
    With grdDeductions
        .Rows = 2
        .ColWidth(0) = 1300
        .ColWidth(1) = 2700
        .ColWidth(2) = 800
        .ColWidth(3) = 1
        .ColWidth(4) = 300
        .ColWidth(5) = 300
        .Row = 0
        .Col = 0
        .Text = "Date"
        .Col = 1
        .Text = "Particular"
        .Col = 2
        .Text = "Amount"
        .Col = 3
        .Text = "ID"
        .Col = 4
        .Text = "H"
        .Col = 5
        .Text = "M"
    End With
End Sub

Sub InitMemvars()
    Dim rsCutoff                                                      As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM HRMS_PAYROLLSETUP")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            cboQuensina.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            cboQuensina.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        'Call AddMonthName
        'cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        LABMONTH.Caption = MonthName(Null2String(rsCutoff!PERIODMONTH))
            
        'cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        'cboYear.Text = Null2String(rsCutoff!PERIODYEAR)

    End If
    'fillcboDay cboDay
    'cboDay.Text = Day(Now)
    txtNoOfMinutes.Text = ""
    txtNoOfMinutes.Enabled = False
    txtAmount.Text = 0
    txtNoHours_Computed = 0
    txtNoMin_Computed = 0
    FIllCbo
End Sub

Sub StoreMemVars()
    Dim DEDDESC                                                       As String
    Dim CNT                                                           As Integer
    Dim VYTDUTLate                                                    As Double
    Dim VYTDAbsent                                                    As Double

    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Set rsDeductions = New ADODB.Recordset
        If DEDUCTION_OPTION = "ATTENDANCE DEDUCTION" Then
            rsDeductions.Open "SELECT * FROM HRMS_DEDUCTIONS WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND CUT_OFF = '" & CUTTOFF_CODE & "' AND PAY_MONTH = '" & PAY_MONTH & "' AND PAY_YEAR = '" & PAY_YEAR & "' AND (PARTICULAR = 'LT' OR PARTICULAR = 'UT' OR PARTICULAR = 'WD' OR PARTICULAR = 'HD')  ORDER BY DEYT DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsDeductions.Open "SELECT * FROM HRMS_DEDUCTIONS WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND CUT_OFF = '" & CUTTOFF_CODE & "' AND PAY_MONTH = '" & PAY_MONTH & "' AND PAY_YEAR = '" & PAY_YEAR & "' AND (PARTICULAR <> 'LT' AND PARTICULAR <> 'UT' AND PARTICULAR <> 'HD' AND PARTICULAR <> 'WD')  ORDER BY DEYT DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        CNT = 0
        VYTDUTLate = 0
        VYTDAbsent = 0
        If Not rsDeductions.EOF And Not rsDeductions.BOF Then
            rsDeductions.MoveFirst
            cleargrid grdDeductions
            grdDeductions.Rows = grdDeductions.Rows
            Do While Not rsDeductions.EOF
                CNT = CNT + 1
                LABID.Caption = rsDeductions!ID
                DEDDESC = GetDeductionDescription(Null2String(rsDeductions!PARTICULAR))
                grdDeductions.AddItem Format(Null2Date(rsDeductions!DEYT), "mm/dd/yyyy") & Chr(9) & _
                    DEDDESC & Chr(9) & _
                    N2Str2Zero(rsDeductions!AMOUNT) & Chr(9) & _
                    rsDeductions!ID & Chr(9) & _
                    (N2Str2Zero(rsDeductions!nomin) \ 60) & Chr(9) & _
                    (N2Str2Zero(rsDeductions!nomin) Mod 60)
                    
                If (Null2String(rsDeductions!PARTICULAR) = "LT" Or Null2String(rsDeductions!PARTICULAR) = "UT") And YEAR(Null2Date(rsDeductions!DEYT)) = YEAR(LOGDATE) Then
                    VYTDUTLate = VYTDUTLate + N2Str2Zero(rsDeductions!AMOUNT)
                End If
                If Null2String(rsDeductions!PARTICULAR) = "HD" Or Null2String(rsDeductions!PARTICULAR) = "WD" And YEAR(Null2Date(rsDeductions!DEYT)) = YEAR(LOGDATE) Then
                    VYTDAbsent = VYTDAbsent + N2Str2Zero(rsDeductions!AMOUNT)
                End If
                rsDeductions.MoveNext
            Loop
            grdDeductions.RemoveItem 1
        Else
            cleargrid grdDeductions
        End If
        txtYTDUTLate.Text = N2Str2Zero(VYTDUTLate)
        txtYTDAbsent.Text = N2Str2Zero(VYTDAbsent)
        txtPosition.Text = Null2String(rsEmpInfo!Position)
        txtName.Text = Cap1st(Null2String(rsEmpInfo!lastname)) & ", " & Cap1st(Null2String(rsEmpInfo!FIRSTNAME)) & " " & Cap1st(Null2String(rsEmpInfo!MIDDLENAME))
        txtEmpno = Null2String(rsEmpInfo!EMPNO)
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub FIllCbo()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim ENTRYBY                                                       As String
    Dim SPACES                                                        As String
    If DEDUCTION_OPTION = "ATTENDANCE DEDUCTION" Then
        'Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_DEDUCTIONCODE WHERE ENTRYBY = 'TIME' ORDER BY DESCRIPTION ASC")
        Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_DEDUCTIONCODE WHERE (CODE = 'LT' OR CODE = 'UT' OR CODE = 'WD' OR CODE = 'HD') ORDER BY DESCRIPTION ASC")
    Else
        'Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_DEDUCTIONCODE WHERE ENTRYBY <> 'TIME' ORDER BY DESCRIPTION ASC")
        Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_DEDUCTIONCODE WHERE (CODE <> 'LT' AND CODE <> 'UT' AND CODE <> 'WD' AND CODE <> 'HD') ORDER BY DESCRIPTION ASC")
    End If
    cboParticular.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            SPACES = "                    "
            If Null2String(RSTMP!ENTRYBY) = "TIME" Then
                ENTRYBY = "T"
            End If
            If Null2String(RSTMP!ENTRYBY) = "AMOUNT" Then
                ENTRYBY = "A"
            End If
            If Null2String(RSTMP!ENTRYBY) = "SALARY" Then
                ENTRYBY = "S"
            End If
            cboParticular.AddItem Null2String(RSTMP!Description) & SPACES & " - " & Null2String(RSTMP!CODE) & "-" & ENTRYBY
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    'Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME,EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno ,ID from HRMS_EmpInfo WHERE RESIGNED IS NULL order by lastname+', '+firstname asc")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    'Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME,EMPNO FROM HRMS_EMPINFO  WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL AND LASTNAME+', '+FIRSTNAME LIKE '" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME,EMPNO FROM HRMS_EMPINFO  WHERE RESIGNED IS NULL AND LASTNAME+', '+FIRSTNAME LIKE '" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
        
        If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub ComputeHours()
    On Error Resume Next
    txtNoOfMinutes = NumericVal(txtNoHours_Computed) * 60 + (NumericVal(txtNoMin_Computed))
End Sub

Private Sub cboParticular_Change()
    EnableDisabledTimeAndAmount
End Sub

Private Sub cboParticular_Click()
    EnableDisabledTimeAndAmount
End Sub

Private Sub cboParticular_LostFocus()
    EnableDisabledTimeAndAmount
End Sub

Private Sub cmdAdd_Click()
    If DEDUCTION_OPTION = "OTHER DEDUCTIONS" Then
        If Function_Access(LOGID, "Acess_Add", "EMPLOYEE MAINTAIN OTHER DEDUCTIONS") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Add", "EMPLOYEE MAINTAIN DEDUCTIONS") = False Then Exit Sub
    End If
    ADDOREDIT = "ADD"
    picDeductions.Visible = True
    cmdDeductions.Visible = True
    picSearch.Enabled = False
    picDeductions.Enabled = True
    cmdDeductions.ZOrder 0
    picDeductions.ZOrder 0
    Picture1.Visible = False
    Picture2.Visible = True
    EnablePics False
    InitMemvars
    
   'JBF: update aug 31 2010
    dt_deduct = LOGDATE
   '******************
    
End Sub

Private Sub cmdCancel_Click()
    lsAdjustment.Enabled = True
    ADDOREDIT = ""
    picDeductions.Visible = False
    cmdDeductions.Visible = False
    picSearch.Enabled = True
    picDeductions.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    cmdDeductions.ZOrder 1
    picDeductions.ZOrder 1
    EnablePics True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If DEDUCTION_OPTION = "OTHER DEDUCTIONS" Then
        If Function_Access(LOGID, "Acess_Delete", "EMPLOYEE MAINTAIN OTHER DEDUCTIONS") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Delete", "EMPLOYEE MAINTAIN DEDUCTIONS") = False Then Exit Sub
    End If
    grdDeductions.Col = 3
    If grdDeductions.Text <> "" Then
        If MsgBox("Delete this Selected record, Are you sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub '
        
        gconDMIS.Execute "DELETE FROM HRMS_DEDUCTIONS WHERE ID = " & grdDeductions.Text
        LogAudit "D", "EMPLOYEE MAINTAIN DEDUCTIONS", grdDeductions.Text
        Call ShowDeletedMsg
    Else
        ShowNothingToDeleteMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If DEDUCTION_OPTION = "OTHER DEDUCTIONS" Then
        If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE MAINTAIN OTHER DEDUCTIONS") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE MAINTAIN DEDUCTIONS") = False Then Exit Sub
    End If
    Dim fild                                                          As String
    grdDeductions.Row = grdDeductions.Row
    grdDeductions.Col = 3
    fild = grdDeductions.Text
    If fild <> "" Then
        ADDOREDIT = "EDIT"
        cmdDeductions.Visible = True
        cmdDeductions.ZOrder 0
        picSearch.Enabled = False
        picDeductions.Visible = True
        picDeductions.ZOrder 0
        picDeductions.Enabled = True
        Picture1.Visible = False
        Picture2.Visible = True
        EnablePics False
        StoreEntry fild
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    picSearch.ZOrder 0
    On Error Resume Next
    txtsearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsEmpInfo.MoveNext
    If rsEmpInfo.EOF Then
        rsEmpInfo.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsEmpInfo.MovePrevious
    If rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    Dim MM As Integer
    Dim YY As Integer
    Dim DD As Integer
    
    
    If DEDUCTION_OPTION = "OTHER DEDUCTIONS" Then
        If Function_Access(LOGID, "Acess_Print", "EMPLOYEE MAINTAIN OTHER DEDUCTIONS") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Print", "EMPLOYEE MAINTAIN DEDUCTIONS") = False Then Exit Sub
    End If
    Screen.MousePointer = 11
    Dim VCUT_OFF                                                      As Integer
    If cboQuensina.Text = "1st Cut-Off" Then
        VCUT_OFF = 1
    End If
    If cboQuensina.Text = "2nd Cut-Off" Then
        VCUT_OFF = 2
    End If
    
   

    MM = MONTH(dt_deduct)
    YY = YEAR(dt_deduct)
    DD = Day(dt_deduct)
    
    
    
    rptDeductions.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptDeductions.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptDeductions.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptDeductions.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
    If DEDUCTION_OPTION = "ATTENDANCE DEDUCTION" Then
        'PrintSQLReport rptDeductions, HRMS_REPORT_PATH & "DEDUCTIONS.rpt", "{Deductions.empno} = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND {Deductions.CUT_OFF} = " & N2Str2Null(VCUT_OFF) & " AND {Deductions.PAY_MONTH} = " & MM & " AND {Deductions.PAY_YEAR} = " & YY & " AND ({Deductions.PARTICULAR} = 'WD' OR {Deductions.PARTICULAR} = 'HD' OR {Deductions.PARTICULAR} = 'LT' OR {Deductions.PARTICULAR} = 'UT')", DMIS_REPORT_Connection, 1
        PrintSQLReport rptDeductions, HRMS_REPORT_PATH & "DEDUCTIONS.rpt", "{Deductions.empno} = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND {Deductions.CUT_OFF} = '" & CUTTOFF_CODE & " ' AND {Deductions.PAY_MONTH} = " & PAY_MONTH & " AND {Deductions.PAY_YEAR} = " & PAY_YEAR & " AND ({Deductions.PARTICULAR} = 'WD' OR {Deductions.PARTICULAR} = 'HD' OR {Deductions.PARTICULAR} = 'LT' OR {Deductions.PARTICULAR} = 'UT')", DMIS_REPORT_Connection, 1
    Else
        'PrintSQLReport rptDeductions, HRMS_REPORT_PATH & "DEDUCTIONS.rpt", "{Deductions.empno} = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND {Deductions.CUT_OFF} = " & N2Str2Null(VCUT_OFF) & " AND {Deductions.PAY_MONTH} = " & MM & " AND {Deductions.PAY_YEAR} = " & YY & " AND ({Deductions.PARTICULAR} <> 'WD' AND {Deductions.PARTICULAR} <> 'HD' AND {Deductions.PARTICULAR} <> 'LT' AND {Deductions.PARTICULAR} <> 'UT')", DMIS_REPORT_Connection, 1
        PrintSQLReport rptDeductions, HRMS_REPORT_PATH & "DEDUCTIONS.rpt", "{Deductions.empno} = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND {Deductions.CUT_OFF} = '" & CUTTOFF_CODE & " ' AND {Deductions.PAY_MONTH} = " & PAY_MONTH & " AND {Deductions.PAY_YEAR} = " & PAY_YEAR & " AND ({Deductions.PARTICULAR} <> 'WD' and {Deductions.PARTICULAR} <> 'HD' and {Deductions.PARTICULAR} <> 'LT' and {Deductions.PARTICULAR} <> 'UT')", DMIS_REPORT_Connection, 1
    End If
    LogAudit "V", "PRINT EMPLOYEE DEDUCTION", ""
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    Dim DEDID                                                         As String
    Dim MM                                                            As String
    Dim YY                                                            As String
    Dim VCUT_OFF                                                      As Integer
    Dim DD                                                            As String
    
    
    MM = MONTH(dt_deduct)
    YY = YEAR(dt_deduct)
    DD = Day(dt_deduct)
        
    Diyt = DateSerial(YY, MM, DD)
    DEDID = Mid(cboParticular, Len(cboParticular) - 3, 2)
    
    If cboParticular.Text = "" Then
        On Error Resume Next
        MsgBox "Select type of deduction", vbInformation, "Info."
        cboParticular.SetFocus
        Exit Sub
    End If
    
    
     ' *************************************
     'JBF Check if data exist
        
        If ADDOREDIT = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select * from HRMS_DEDUCTIONS where EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND deyt = " & N2Date2Null(Diyt) & " and particular = '" & DEDID & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Data already exist!"
                Exit Sub
            End If
        
        Else
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select * from HRMS_DEDUCTIONS where EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND deyt = " & N2Date2Null(Diyt) & " and particular = '" & DEDID & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Data already exist!"
                Exit Sub
            End If
        
        End If
    
    ' *************************************
    
    If cboQuensina.Text = "1st Cut-Off" Then
        VCUT_OFF = 1
    End If
    If cboQuensina.Text = "2nd Cut-Off" Then
        VCUT_OFF = 2
    End If
    


    If ADDOREDIT = "ADD" Then
        'COMMENT BY  : MJP 010909 1003AM
        'DESCRIPTION :
            'gconDMIS.Execute ("INSERT INTO HRMS_DEDUCTIONS " & _
            '                  "(EMPLEVEL, EMPNO, DEYT, PARTICULAR, AMOUNT, NOMIN, CUT_OFF, PAY_MONTH, PAY_YEAR, MANUAL) " & _
            '                  "VALUES (" & EMPLIVIL & _
            '                  "," & N2Str2Null(RSEMPINFO!EMPNO) & _
            '                  ", " & N2Date2Null(Diyt) & _
            '                  ", '" & DEDID & _
            '                  "', " & NumericVal(txtAmount.Text) & _
            '                  ", " & NumericVal(txtNoOfMinutes) & _
            '                  "," & VCUT_OFF & _
            '                  "," & What_month(LABMONTH) & _
            '                  "," & cboYear & ",'Y')")
        'COMMENT BY  : MJP 010909 1003AM
        
        'UPDATED BY  : MJP 010909 1003AM
        'DESCRIPTION :
            gconDMIS.Execute ("INSERT INTO HRMS_DEDUCTIONS " & _
                "(EMPLEVEL, EMPNO, DEYT, PARTICULAR, AMOUNT, NOMIN, CUT_OFF, PAY_MONTH, PAY_YEAR, MANUAL) " & _
                "VALUES (" & EMPLIVIL & _
                "," & N2Str2Null(rsEmpInfo!EMPNO) & _
                ", " & N2Date2Null(Diyt) & _
                ", '" & DEDID & _
                "', " & NumericVal(txtAmount.Text) & _
                ", " & NumericVal(txtNoOfMinutes) & _
                "," & VCUT_OFF & _
                "," & What_month(LABMONTH) & _
                "," & PAY_YEAR & ",'Y')")
        'UPDATED BY  : MJP 010909 1003AM
        
        LogAudit "A", "EMPLOYEE MAINTAIN DEDUCTIONS", rsEmpInfo!EMPNO
        ShowSuccessFullyAdded
    Else
        'COMMENT BY  : MJP 010909 1003AM
        'DESCRIPTION :
            'gconDMIS.Execute "UPDATE HRMS_DEDUCTIONS SET" & _
            '           " EMPLEVEL = " & EMPLIVIL & "," & _
            '           " EMPNO = " & N2Str2Null(RSEMPINFO!EMPNO) & "," & _
            '           " DEYT = " & N2Date2Null(Diyt) & "," & _
            '           " PARTICULAR = '" & DEDID & "'," & _
            '           " NOMIN = " & NumericVal(txtNoOfMinutes) & "," & _
            '           " AMOUNT = " & NumericVal(txtAmount.Text) & "," & _
            '           " CUT_OFF = " & VCUT_OFF & "," & _
            '           " PAY_MONTH = " & What_month(LABMONTH) & "," & _
            '           " PAY_YEAR = " & cboYear & _
            '           " WHERE ID = " & labID.Caption
        'COMMENT BY  : MJP 010909 1003AM
            
        'UPDATE BY   : MJP 010909 1003AM
        'DESCRIPTION :
            gconDMIS.Execute "UPDATE HRMS_DEDUCTIONS SET" & _
                       " EMPLEVEL = " & EMPLIVIL & "," & _
                       " EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & "," & _
                       " DEYT = " & N2Date2Null(Diyt) & "," & _
                       " PARTICULAR = '" & DEDID & "'," & _
                       " NOMIN = " & NumericVal(txtNoOfMinutes) & "," & _
                       " AMOUNT = " & NumericVal(txtAmount.Text) & "," & _
                       " CUT_OFF = " & VCUT_OFF & "," & _
                       " PAY_MONTH = " & What_month(LABMONTH) & "," & _
                       " PAY_YEAR = " & PAY_YEAR & _
                       " WHERE ID = " & LABID.Caption
        'UPDATE BY   : MJP 010909 1003AM
        
        LogAudit "E", "EMPLOYEE MAINTAIN DEDUCTIONS", rsEmpInfo!EMPNO
        ShowSuccessFullyUpdated
    End If
    cmdCancel.Value = True
    StoreMemVars
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            cmdAdd.Value = True
        Case vbKeyEscape
            cmdCancel.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then
        EMPLIVIL = "'C'"
    End If
    If EMP_TYPE = "ALLOWANCE BASE" Then
        EMPLIVIL = "'A'"
    End If
    
    txtsearch.Text = ""
    rsSETUP
    rsrefresh
    InitGrid
    InitMemvars
    cmdCancel_Click
    'DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHRMSDeductions = Nothing
End Sub

Private Sub grdDeductions_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    rsEmpInfo.Bookmark = rsFIND(rsEmpInfo.Clone, "empno", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lsAdjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsAdjustment
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

Private Sub lsAdjustment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtNoOfMinutes_Change()
    Dim rperday                                                       As Double
    Dim rperHour                                                      As Double
    Dim rperMinute                                                    As Double
    rperday = (GetBasicPay * 12) / DAYS_OF_WORK
    rperHour = rperday / 8
    rperMinute = rperHour / 60
    'If UCase(Mid(cboParticular, 1, 4)) = "WHOL" Or UCase(Mid(cboParticular, 1, 4)) = "HALF" Then
    If UCase(Mid(Right(cboParticular, 4), 1, 2)) = "WD" Or UCase(Mid(Right(cboParticular, 4), 1, 2)) = "HD" Then
        txtAmount.Text = Round((NumericVal(txtNoOfMinutes.Text) / 480) * Round(rperday, 2), 2)
    Else
        txtAmount.Text = Round(NumericVal(txtNoOfMinutes.Text) * Round(rperMinute, 2), 2)
    End If
End Sub

Private Sub txtsearch_Change()
    If Trim(txtsearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtsearch.Text)
    End If
End Sub

Private Sub txtNoMin_Computed_Change()
    If NumericVal(txtNoMin_Computed) > 59 Then txtNoMin_Computed = 59
    ComputeHours
End Sub

Private Sub txtNoMin_Computed_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub txtNoHours_Computed_Change()
    ComputeHours
End Sub

Private Sub txtNoHours_Computed_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmHRMSCommission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Commission Entry"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8985
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Commission.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   8985
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3270
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   26
      Top             =   5310
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
         MouseIcon       =   "Commission.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":045C
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
         MouseIcon       =   "Commission.frx":07C2
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":0914
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
         MouseIcon       =   "Commission.frx":0C7A
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":0DCC
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
         MouseIcon       =   "Commission.frx":10F7
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":1249
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
         MouseIcon       =   "Commission.frx":15A5
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":16F7
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
         MouseIcon       =   "Commission.frx":1A0A
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":1B5C
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Move to Next Record"
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
         MouseIcon       =   "Commission.frx":1E56
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":1FA8
         Style           =   1  'Graphical
         TabIndex        =   28
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
         MouseIcon       =   "Commission.frx":2300
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":2452
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox fraCommission 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   4110
      ScaleHeight     =   2295
      ScaleWidth      =   3465
      TabIndex        =   10
      Top             =   2010
      Width           =   3465
      Begin VB.ComboBox cboQuensina 
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
         Left            =   30
         Style           =   1  'Simple Combo
         TabIndex        =   38
         Text            =   "cboQuensina"
         Top             =   60
         Width           =   3195
      End
      Begin VB.ComboBox cboDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   930
         Style           =   1  'Simple Combo
         TabIndex        =   18
         Text            =   "cboMonth"
         Top             =   720
         Width           =   1545
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   2520
         Style           =   1  'Simple Combo
         TabIndex        =   17
         Text            =   "cboYear"
         Top             =   720
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   315
         Left            =   1470
         TabIndex        =   0
         Top             =   1140
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTax 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   1500
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNetAmount 
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   1860
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   150
         TabIndex        =   22
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Left            =   1260
         TabIndex        =   21
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   2550
         TabIndex        =   20
         Top             =   480
         Width           =   555
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         Caption         =   "ID"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2250
         TabIndex        =   14
         Top             =   1500
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Amount"
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
         Height          =   285
         Left            =   30
         TabIndex        =   13
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "w/ Tax"
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
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   1530
         Width           =   1155
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   1890
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1365
      Left            =   2610
      ScaleHeight     =   1365
      ScaleWidth      =   6375
      TabIndex        =   4
      Top             =   45
      Width           =   6375
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
         Left            =   930
         TabIndex        =   7
         Top             =   480
         Width           =   5355
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
         Left            =   90
         TabIndex        =   6
         Top             =   60
         Width           =   6195
      End
      Begin VB.TextBox txtYTDCommission 
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
         Left            =   1710
         TabIndex        =   5
         Top             =   900
         Width           =   1605
      End
      Begin Crystal.CrystalReport rptCommission 
         Left            =   30
         Top             =   0
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
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Commission"
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
         Left            =   120
         TabIndex        =   8
         Top             =   930
         Width           =   1755
      End
   End
   Begin wizButton.cmd cmdCommission 
      Height          =   2415
      Left            =   4050
      TabIndex        =   3
      Top             =   1950
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   4260
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
      MICON           =   "Commission.frx":27B1
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   2610
      ScaleHeight     =   3885
      ScaleWidth      =   6375
      TabIndex        =   15
      Top             =   1455
      Width           =   6375
      Begin MSFlexGridLib.MSFlexGrid grdCommission 
         Height          =   3735
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   5
         ForeColor       =   0
         BackColorFixed  =   14606302
         ForeColorFixed  =   0
         BackColorSel    =   14606302
         ForeColorSel    =   0
         BackColorBkg    =   14606302
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
         MouseIcon       =   "Commission.frx":27CD
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6120
      Left            =   60
      Picture         =   "Commission.frx":2AE7
      ScaleHeight     =   6090
      ScaleWidth      =   2445
      TabIndex        =   23
      Top             =   45
      Width           =   2475
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
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   30
         Width           =   2385
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   5625
         Left            =   30
         TabIndex        =   25
         Top             =   420
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   9922
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
         MouseIcon       =   "Commission.frx":5823
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
         Picture         =   "Commission.frx":5985
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7425
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   35
      Top             =   5310
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
         MouseIcon       =   "Commission.frx":196F2
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":19844
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
         MouseIcon       =   "Commission.frx":19B82
         MousePointer    =   99  'Custom
         Picture         =   "Commission.frx":19CD4
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo, rsCommission                                           As ADODB.Recordset
Attribute rsCommission.VB_VarUserMemId = 1073938432
Dim AddorEdit, Diyt                                                   As String
Attribute AddorEdit.VB_VarUserMemId = 1073938434
Attribute Diyt.VB_VarUserMemId = 1073938434
Dim EMPLIVIL                                                          As String
Attribute EMPLIVIL.VB_VarUserMemId = 1073938436

Function StoreEntry(ByVal ID As Variant)
    Dim MM, DD, YY                                                    As String
    Dim TheDeyt                                                       As String
    Set rsCommission = New ADODB.Recordset
    rsCommission.Open "select * from HRMS_Commission where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCommission.EOF And Not rsCommission.BOF Then
        If rsCommission!CUT_OFF = 1 Then cboQuensina.Text = "1st Cut-Off"
        If rsCommission!CUT_OFF = 2 Then cboQuensina.Text = "2nd Cut-Off"
        labID.Caption = rsCommission!ID
        TheDeyt = Null2Date(rsCommission!DEYT)
        DD = Day(TheDeyt)
        MM = The_month(MONTH(TheDeyt))
        YY = YEAR(TheDeyt)
        cboDay.Text = DD
        cboMonth.Text = MM
        cboYear.Text = YY
        txtAmount.Text = N2Str2Zero(rsCommission!AMOUNT)
        txtTax.Text = N2Str2Zero(rsCommission!TAX)
    End If
End Function

Sub EnablePics(COND As Boolean)
    picSearch.Enabled = COND
    Picture5.Enabled = COND
End Sub

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select empno,[position],lastname,firstname,middlename,emplevel,resigned from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & EMPINFOEMPNO.Caption & "'", gconDMIS
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select empno,[position],lastname,firstname,middlename,emplevel,resigned from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & frmHRMSEmpInfo.labID.Caption & "'", gconDMIS
    Else
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select empno,[position],lastname,firstname,middlename,emplevel,resigned from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL order by lastname asc", gconDMIS
    End If
End Sub

Sub InitGrid()
    With grdCommission
        .Rows = 2
        .ColWidth(0) = 1300
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1
        .Row = 0
        .Col = 0
        .Text = "Date"
        .Col = 1
        .Text = "Amount"
        .Col = 2
        .Text = "Tax"
        .Col = 3
        .Text = "Total"
        .Col = 4
        .Text = "ID"
    End With
End Sub

Sub InitMemvars()
    '    cboQuensina.Clear
    '    cboQuensina.AddItem "1st Cut-Off"
    '    cboQuensina.AddItem "2nd Cut-Off"
    Dim rsCutoff                                                      As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM ALL_PROFILE WHERE MODULENAME = 'HRMS'")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            cboQuensina.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            cboQuensina.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        cboYear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If
    fillcboDay cboDay
    '    fillcbomonth cboMonth
    '    FillcboYear cboYear
    '    cboYear.Text = Year(LOGDATE)
    '    cboMonth.Text = The_month(Month(LOGDATE))
    cboDay.Text = Day(Now)
    txtAmount.Text = "0.00"
    txtTax.Text = "0.00"
    txtNetAmount.Text = "0.00"
End Sub

Sub StoreMemVars()
    On Error GoTo Errorcode
    Dim CNT                                                           As Integer
    Dim VYTDCommission, VYTDCommissionTax                             As Double
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Set rsCommission = New ADODB.Recordset
        'rsCommission.Open "select * from HRMS_Commission where EMPLEVEL = " & EMPLIVIL & " AND empno = " & N2Str2Null(rsEmpInfo!EMPNO) & " order by deyt desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        rsCommission.Open "SELECT * FROM HRMS_COMMISSION WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND CUT_OFF = '" & CUTTOFF_CODE & "' AND PAY_MONTH = '" & PAY_MONTH & "' AND PAY_YEAR = '" & PAY_YEAR & "' ORDER BY DEYT DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        CNT = 0: VYTDCommission = 0: VYTDCommissionTax = 0
        If Not rsCommission.EOF And Not rsCommission.BOF Then
            rsCommission.MoveFirst
            cleargrid grdCommission
            grdCommission.Rows = grdCommission.Rows
            Do While Not rsCommission.EOF
                CNT = CNT + 1
                labID.Caption = rsCommission!ID
                grdCommission.AddItem Null2Date(rsCommission!DEYT) & Chr(9) & N2Str2Zero(rsCommission!AMOUNT) & Chr(9) & N2Str2Zero(rsCommission!TAX) & Chr(9) & Format(N2Str2Zero(rsCommission!AMOUNT) - N2Str2Zero(rsCommission!TAX), MAXIMUM_DIGIT) & Chr(9) & rsCommission!ID
                If YEAR(Null2Date(rsCommission!DEYT)) = YEAR(LOGDATE) Then
                    VYTDCommission = VYTDCommission + N2Str2Zero(rsCommission!AMOUNT)
                    VYTDCommissionTax = VYTDCommissionTax + N2Str2Zero(rsCommission!TAX)
                End If
                rsCommission.MoveNext
            Loop
            grdCommission.RemoveItem 1
        Else
            cleargrid grdCommission
        End If
        txtYTDCommission.Text = N2Str2Zero(VYTDCommission)
        'txtYTDCommissionTax.Text = N2Str2Zero(VYTDCommissionTax)
        txtPosition.Text = Null2String(rsEmpInfo!Position)
        txtName.Text = Cap1st(Null2String(rsEmpInfo!lastname)) & ", " & Cap1st(Null2String(rsEmpInfo!FIRSTNAME)) & " " & Cap1st(Null2String(rsEmpInfo!MIDDLENAME))
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False: lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL order by lastname+', '+firstname asc")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False: lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo  where EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL and lastname+', '+firstname like'" & XXX & "%' order by lastname+', '+firstname asc")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Add", "EMPLOYEE MAINTAIN COMMISSION") = False Then Exit Sub
    AddorEdit = "ADD"
    fraCommission.Visible = True
    cmdCommission.Visible = True
    fraCommission.Enabled = True
    cmdCommission.ZOrder 0
    fraCommission.ZOrder 0
    Picture1.Visible = False
    Picture2.Visible = True
    EnablePics False
    InitMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo Errorcode:
    lsAdjustment.Enabled = True
    AddorEdit = ""
    fraCommission.Visible = False
    cmdCommission.Visible = False
    fraCommission.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    cmdCommission.ZOrder 1
    fraCommission.ZOrder 1
    EnablePics True
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Delete", "EMPLOYEE MAINTAIN COMMISSION") = False Then Exit Sub
    grdCommission.Col = 4
    If grdCommission.Text <> "" Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from HRMS_Commission where id = " & grdCommission.Text
            LogAudit "X", "DELETE EMPLOYEE COMMISSION RECORD", grdCommission.Text
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE MAINTAIN COMMISSION") = False Then Exit Sub
    Dim fild                                                          As String
    grdCommission.Row = grdCommission.Row
    grdCommission.Col = 4
    fild = grdCommission.Text
    If fild <> "" Then
        lsAdjustment.Enabled = False
        AddorEdit = "EDIT"
        cmdCommission.Visible = True
        cmdCommission.ZOrder 0
        fraCommission.Visible = True
        fraCommission.ZOrder 0
        fraCommission.Enabled = True
        Picture1.Visible = False
        Picture2.Visible = True
        EnablePics False
        StoreEntry fild
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    picSearch.ZOrder 0
    On Error Resume Next
    txtSEARCH.SetFocus
    'rsRefresh
    'On Error Resume Next
    'rsCommission.Find "id = " & labID.Caption
    'Dim findStr As String
    'findStr = InputSpeechBox("Please Input Name ...", txtName.Text)
    'If findStr <> "" Then
    '   On Error Resume Next
    '   rsEmpinfo.Bookmark = rsFind(rsEmpinfo.Clone, "lastname", findStr).Bookmark
    '   If Err.Number = 3021 Then
    '      On Error GoTo ErrorCode
    '      rsEmpinfo.Bookmark = rsFind(rsEmpinfo.Clone, "firstname", findStr).Bookmark
    '   End If
    'End If
    'StoreMemvars
    'Exit Sub
    'ErrorCode:
    'If Err.Number = 3021 Then
    '   ShowCantFind findStr
    '   Resume Next
    'End If
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
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Print", "EMPLOYEE MAINTAIN COMMISSION") = False Then Exit Sub
    Screen.MousePointer = 11
    rptCommission.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptCommission.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptCommission.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptCommission.Formulas(3) = "PRINTEDBY = '" & LOGNAME & "'"
    PrintSQLReport rptCommission, HRMS_REPORT_PATH & "Commission.rpt", "{Commission.empno} = " & N2Str2Null(rsEmpInfo!EMPNO), DMIS_REPORT_Connection, 1
    LogAudit "V", "PRINT EMPLOYEE COMMISSION RECORD", ""
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode:
    Dim MM, DD, YY                                                    As String
    Dim vCUTOFF                                                       As Integer
    If cboQuensina.Text = "" Then
        ShowIsRequiredMsg "Choose a Cut_off"
        cboQuensina.SetFocus
        Exit Sub
    End If
    MM = What_month(cboMonth): YY = cboYear.Text: DD = cboDay.Text
    Diyt = DateSerial(YY, MM, DD)
    If cboQuensina.Text = "1st Cut-Off" Then vCUTOFF = 1
    If cboQuensina.Text = "2nd Cut-Off" Then vCUTOFF = 2
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_Commission " & _
                         "(EMPLEVEL, Empno, Deyt, Amount, Tax, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
                       " Values (" & EMPLIVIL & _
                         "," & N2Str2Null(rsEmpInfo!EMPNO) & _
                         ", " & N2Date2Null(Diyt) & _
                         ", " & NumericVal(txtAmount.Text) & _
                         ", " & NumericVal(txtTax.Text) & _
                         "," & vCUTOFF & _
                         "," & MM & _
                         "," & YY & ")"
        LogAudit "A", "ADD EMPLOYEE COMMISSION RECORD", rsEmpInfo!EMPNO
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "Update HRMS_Commission set" & _
                       " EMPLEVEL = " & EMPLIVIL & "," & _
                       " Empno = " & N2Str2Null(rsEmpInfo!EMPNO) & "," & _
                       " Deyt = " & N2Date2Null(Diyt) & "," & _
                       " Amount = " & NumericVal(txtAmount.Text) & "," & _
                       " Tax = " & NumericVal(txtTax.Text) & "," & _
                       " CUT_OFF = " & vCUTOFF & "," & _
                       " PAY_MONTH = " & MM & "," & _
                       " PAY_YEAR = " & YY & _
                       " where id = " & labID.Caption
        LogAudit "E", "UPDATE EMPLOYEE COMMISSION RECORD", rsEmpInfo!EMPNO
        ShowSuccessFullyUpdated
    End If
    cmdCancel.Value = True
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
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
    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    txtSEARCH.Text = ""
    rsrefresh
    InitGrid
    InitMemvars
    cmdCancel_Click
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub grdCommission_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtAmount_Change()
    txtNetAmount.Text = NumericVal(txtAmount.Text) - NumericVal(txtTax.Text)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtNetAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtTax_Change()
    txtNetAmount.Text = NumericVal(txtAmount.Text) - NumericVal(txtTax.Text)
End Sub

Private Sub txtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsEmpInfo.Bookmark = rsFind(rsEmpInfo.Clone, "empno", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lsAdjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsAdjustment
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub lsAdjustment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSEARCH.Text) = "" Then FillGrid Else FillSearchGrid (txtSEARCH.Text)
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmHRMS_Leave_Maintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Maintenance"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Leave_Maintenance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   10770
   Begin Crystal.CrystalReport rptmaintain 
      Left            =   3720
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H80000001&
      Height          =   315
      Left            =   6540
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   360
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H80000001&
      Height          =   345
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   360
      Width           =   2985
   End
   Begin VB.ComboBox cboyear 
      Enabled         =   0   'False
      Height          =   330
      Left            =   8970
      TabIndex        =   27
      Top             =   360
      Width           =   1245
   End
   Begin VB.PictureBox picSearch 
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
      Height          =   4065
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   2925
      TabIndex        =   10
      Top             =   0
      Width           =   2955
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   60
         MaxLength       =   35
         TabIndex        =   11
         Top             =   390
         Width           =   2835
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   3225
         Left            =   30
         TabIndex        =   12
         Top             =   780
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   5689
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
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
         MouseIcon       =   "Leave_Maintenance.frx":1082
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Employee Name"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empno"
            Object.Width           =   2
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   345
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   8085
         _Version        =   655364
         _ExtentX        =   14261
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   " Search"
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
      Height          =   855
      Left            =   4470
      ScaleHeight     =   855
      ScaleWidth      =   6300
      TabIndex        =   0
      Top             =   3240
      Width           =   6300
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   5580
         MouseIcon       =   "Leave_Maintenance.frx":11E4
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":1336
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   4890
         MouseIcon       =   "Leave_Maintenance.frx":169C
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":17EE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   4200
         MouseIcon       =   "Leave_Maintenance.frx":1B54
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":1CA6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3510
         MouseIcon       =   "Leave_Maintenance.frx":1FA0
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":20F2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   2820
         MouseIcon       =   "Leave_Maintenance.frx":244E
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":25A0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   2130
         MouseIcon       =   "Leave_Maintenance.frx":28F8
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":2A4A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
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
      Left            =   2970
      ScaleHeight     =   2385
      ScaleWidth      =   7755
      TabIndex        =   14
      Top             =   810
      Width           =   7785
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2040
         Width           =   1980
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1710
         Width           =   1980
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1380
         Width           =   1980
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1050
         Width           =   1980
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   720
         Width           =   1980
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   1980
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1050
         Width           =   1980
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1380
         Width           =   1980
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1710
         Width           =   1980
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2040
         Width           =   1980
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1740
         TabIndex        =   19
         Top             =   720
         Width           =   1980
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1740
         TabIndex        =   18
         Top             =   1050
         Width           =   1980
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1740
         TabIndex        =   17
         Top             =   1380
         Width           =   1980
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1740
         TabIndex        =   16
         Top             =   1710
         Width           =   1980
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1740
         TabIndex        =   15
         Top             =   2040
         Width           =   1980
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   -90
         TabIndex        =   46
         Top             =   0
         Width           =   8085
         _Version        =   655364
         _ExtentX        =   14261
         _ExtentY        =   661
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Sick Leave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   0
         TabIndex        =   45
         Top             =   720
         Width           =   1725
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Vacation Leave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   0
         TabIndex        =   44
         Top             =   1050
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Emergency Leave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   0
         TabIndex        =   43
         Top             =   1380
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Maternity Leave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   0
         TabIndex        =   42
         Top             =   1710
         Width           =   1725
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Paternity Leave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   0
         TabIndex        =   41
         Top             =   2040
         Width           =   1725
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Leave Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   0
         TabIndex        =   40
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label Label12 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Leaves Taken this Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   5760
         TabIndex        =   38
         Top             =   390
         Width           =   1980
      End
      Begin VB.Label Label7 
         BackColor       =   &H00D2BDB6&
         Caption         =   " This Year Balance Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   3750
         TabIndex        =   26
         Top             =   390
         Width           =   1980
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Last Year Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   1740
         TabIndex        =   20
         Top             =   390
         Width           =   1980
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
      Height          =   885
      Left            =   9330
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   7
      Top             =   3240
      Width           =   1440
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "Leave_Maintenance.frx":2DA9
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":2EFB
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "Leave_Maintenance.frx":3239
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":338B
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
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
      Height          =   3675
      Left            =   810
      ScaleHeight     =   3645
      ScaleWidth      =   7755
      TabIndex        =   39
      Top             =   4500
      Width           =   7785
      Begin FlexCell.Grid Grid1 
         Height          =   2175
         Left            =   720
         TabIndex        =   49
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3836
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   12632256
         Rows            =   30
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   375
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   8085
         _Version        =   655364
         _ExtentX        =   14261
         _ExtentY        =   661
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee No"
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
      Left            =   6510
      TabIndex        =   32
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
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
      Left            =   3030
      TabIndex        =   31
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Height          =   285
      Left            =   9120
      TabIndex        =   30
      Top             =   120
      Width           =   495
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   735
      Left            =   2970
      TabIndex        =   13
      Top             =   30
      Width           =   7905
      _Version        =   655364
      _ExtentX        =   13944
      _ExtentY        =   1296
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
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmHRMS_Leave_Maintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo       As ADODB.Recordset
Dim EL_NO           As Double
Dim ML_NO           As Double
Dim PL_NO           As Double
Dim VL_NO           As Double
Dim SL_NO           As Double

Private Sub cboyear_Click()
    Call StoreMemVars
End Sub

Private Sub cmdAdd_Click()
    Picture2.Visible = True
    Picture1.Visible = False
    'Enable (True)
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    'Enable (False)
    Call StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    Picture2.Visible = True
    Picture1.Visible = False
    Call DNInitTextLastYearBalance
    Text4.SetFocus
    cboyear.Locked = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    'Picture5.Visible = True
    'Picture4.Visible = False
    txtSearch.SetFocus
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    rptmaintain.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptmaintain.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    
    PrintSQLReport rptmaintain, HRMS_REPORT_PATH & "leavemaintain.rpt", "{hrms_leave.emplno} = '" & Text2.Text & "'", DMIS_REPORT_Connection, 1
    'PrintSQLReport rptmaintain, HRMS_REPORT_PATH & "leavemaintain.rpt", "", DMIS_REPORT_Connection, 1
    
    Screen.MousePointer = 0

End Sub

Function ENDOFYEAR(xdate As Date) As Boolean
    Dim X As String
    
    X = CStr(MONTH(xdate))
    
    If X = "12" Then
        ENDOFYEAR = True
    Else
        ENDOFYEAR = False
    End If
End Function

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    If Text2.Text <> "" Then
        Dim rsTemp As New ADODB.Recordset
        Dim RSTMP As New ADODB.Recordset
        Dim sqltxt, XTYPE As String
        
        'Update by: NVB -- ---
        sqltxt = "Select LEAVE_CODE from HRMS_LEAVEMASTER"
        Set RSTMP = gconDMIS.Execute(sqltxt)
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            
        End If
        
        
        RSTMP.MoveFirst
        Do While Not RSTMP.EOF
        XTYPE = Trim(RSTMP!LEAVE_CODE)
        
        Set rsTemp = New ADODB.Recordset
        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVE WHERE EMPLNO = '" & Text2.Text & "' and [type] = '" & XTYPE & "'")
        Select Case XTYPE
             Case "SL"
                If Not rsTemp.EOF And Not rsTemp.BOF Then
                    sqltxt = "UPDATE HRMS_LEAVE SET AVAILABLE = " & N2Str2Null(Text3.Text) & ","
                    sqltxt = sqltxt & "USED = " & N2Str2Null(Text13.Text) & ", DATEASOF = '" & CDate(Now) & "'"
                    sqltxt = sqltxt & " where EMPLNO = '" & Text2.Text & "' AND [TYPE] = 'SL'"
                    gconDMIS.Execute (sqltxt)
                Else
                    sqltxt = "INSERT INTO HRMS_LEAVE (EMPLNO,[type],AVAILABLE,USED,MAXSL,DATEASOF) "
                    sqltxt = sqltxt & "VALUES('" & Text2.Text & "','SL','" & Text3.Text & "','" & Text13.Text & "','" & Text3.Text & "',"
                    sqltxt = sqltxt & "'" & CDate(Now) & "')"
                    gconDMIS.Execute (sqltxt)
                              
                End If
            Case "VL"
                If Not rsTemp.EOF And Not rsTemp.BOF Then
                    sqltxt = "UPDATE HRMS_LEAVE SET AVAILABLE = " & N2Str2Null(Text9.Text) & ","
                    sqltxt = sqltxt & "USED = " & N2Str2Null(Text14.Text) & " , DATEASOF = '" & CDate(Now) & "' where"
                    sqltxt = sqltxt & " EMPLNO = '" & Text2.Text & "' AND [TYPE] = 'VL'"
                    gconDMIS.Execute (sqltxt)
                Else
                    sqltxt = "INSERT INTO HRMS_LEAVE (EMPLNO,[type],AVAILABLE,USED,MAXVL,DATEASOF) "
                    sqltxt = sqltxt & "VALUES('" & Text2.Text & "','VL','" & Text9.Text & "','" & Text14.Text & "','" & Text9.Text & "',"
                    sqltxt = sqltxt & "'" & CDate(Now) & "')"
                    gconDMIS.Execute (sqltxt)
                              
                End If
            Case "EL"
                 If Not rsTemp.EOF And Not rsTemp.BOF Then
                    sqltxt = "UPDATE HRMS_LEAVE SET AVAILABLE = " & N2Str2Null(Text10.Text) & ","
                    sqltxt = sqltxt & "USED = " & N2Str2Null(Text15.Text) & " ,DATEASOF = '" & CDate(Now) & "' where"
                    sqltxt = sqltxt & " EMPLNO = '" & Text2.Text & "' AND [TYPE] = 'EL'"
                    gconDMIS.Execute (sqltxt)
                Else
                    sqltxt = "INSERT INTO HRMS_LEAVE (EMPLNO,[type],AVAILABLE,USED,MAXEL,DATEASOF) "
                    sqltxt = sqltxt & "VALUES('" & Text2.Text & "','EL','" & Text10.Text & "','" & Text15.Text & "','" & Text10.Text & "',"
                    sqltxt = sqltxt & "'" & CDate(Now) & "')"
                    gconDMIS.Execute (sqltxt)
                              
                End If
            Case "ML"
                 If Not rsTemp.EOF And Not rsTemp.BOF Then
                    sqltxt = "UPDATE HRMS_LEAVE SET AVAILABLE = " & N2Str2Null(Text11.Text) & ","
                    sqltxt = sqltxt & "USED = " & N2Str2Null(Text16.Text) & " , DATEASOF = '" & CDate(Now) & "' where "
                    sqltxt = sqltxt & "EMPLNO = '" & Text2.Text & "' AND [TYPE] = 'ML'"
                    gconDMIS.Execute (sqltxt)
                Else
                    sqltxt = "INSERT INTO HRMS_LEAVE (EMPLNO,[type],AVAILABLE,USED,MAXML,DATEASOF) "
                    sqltxt = sqltxt & "VALUES('" & Text2.Text & "','ML','" & Text11.Text & "','" & Text16.Text & "','" & Text11.Text & "',"
                    sqltxt = sqltxt & "'" & CDate(Now) & "')"
                    gconDMIS.Execute (sqltxt)
                              
                End If
            Case "PL"
                  If Not rsTemp.EOF And Not rsTemp.BOF Then
                    sqltxt = "UPDATE HRMS_LEAVE SET AVAILABLE = " & N2Str2Null(Text12.Text) & ","
                    sqltxt = sqltxt & "USED = " & N2Str2Null(Text17.Text) & ",DATEASOF = '" & Now & "' where"
                    sqltxt = sqltxt & " EMPLNO = '" & Text2.Text & "' AND [TYPE] = 'PL'"
                    gconDMIS.Execute (sqltxt)
                Else
                    sqltxt = "INSERT INTO HRMS_LEAVE (EMPLNO,[type],AVAILABLE,USED,MAXPL,DATEASOF) "
                    sqltxt = sqltxt & "VALUES('" & Text2.Text & "','PL','" & Text12.Text & "','" & Text17.Text & "','" & Text12.Text & "',"
                    sqltxt = sqltxt & "'" & Now & "')"
                    gconDMIS.Execute (sqltxt)
                              
                End If
            
            End Select
    '-------
    rsTemp.Close
    RSTMP.MoveNext
    Loop
    End If
    
    Set rsTemp = Nothing
    cmdCancel.Value = True
    
    
    Call StoreMemVars
Errorcode:
    'MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Call FillCombo
    Call rsrefresh
    Call FillGrid
    Picture5.Visible = False
    Call InitMemvars
    Call StoreMemVars
    Screen.MousePointer = 0
    Call SendToBack
    Call InitTextLastYearBalance
    Call YEAR_END
End Sub

Private Sub SendToBack()
    cmdEdit.Enabled = False
    'cmdDelete.Enabled = False
    'cmdPrevious.Enabled = False
    'cmdNext.Enabled = False
End Sub

Private Sub YEAR_END()
    If ENDOFYEAR(CDate(Date)) = True Then
        Call PROCESSAVAILABLELEAVE
        'Call RESTORELYLEAVE
    Else
        Exit Sub
    End If
End Sub

Private Sub PROCESSAVAILABLELEAVE()
    Dim sqltxt As String
    Dim rsEND As New ADODB.Recordset
    Dim xAV As Integer
    Dim xemp, XTYPE As String
    Dim xdate As String
    Dim SQL As String
    Dim DEC_END As Date
    
    On Error GoTo Errorcode
    
    DEC_END = CDate("12/31/" & YEAR(Date)) + 1
    
    sqltxt = "SELECT available,emplno,[type],dateasof FROM HRMS_LEAVE"
    Set rsEND = gconDMIS.Execute(sqltxt)
    If Not (rsEND.BOF And rsEND.EOF) Then
        
    End If
    
    rsEND.MoveFirst
    Do While Not rsEND.EOF
        xAV = Trim(rsEND!Available)
        xemp = Trim(rsEND!EMPLNO)
        XTYPE = Trim(rsEND![Type])
        xdate = CStr(YEAR(Date))
        
        SQL = "UPDATE HRMS_LEAVE SET LASTYEARLEAVE = '" & xAV & "',LASTYEAR = '" & xdate & "',"
        SQL = SQL & "DATEASOF = '" & DEC_END & "'"
        SQL = SQL & " WHERE EMPLNO = '" & xemp & "' AND [TYPE] = '" & XTYPE & "'"
        gconDMIS.Execute (SQL)
                
                
    rsEND.MoveNext
    Loop
    
Errorcode:
    Exit Sub

End Sub

Private Sub RESTOREYLEAVE()
    Dim X1, X2, X3, X4, X5 As Integer
    Dim rsFIND As New ADODB.Recordset
    Dim sqltxt As String
    
    X1 = 0: X2 = 0: X3 = 0: X4 = 0: X5 = 0
    
    sqltxt = "Select * from hrms_leave where emplno = '" & Trim(Text2.Text) & "'"
    Set rsFIND = gconDMIS.Execute(sqltxt)
    If Not (rsFIND.BOF And rsFIND.EOF) Then
    
        X1 = GETLASTYEARLEAVE(Text2.Text, "SL")
        X2 = GETLASTYEARLEAVE(Text2.Text, "VL")
        X3 = GETLASTYEARLEAVE(Text2.Text, "EL")
        X4 = GETLASTYEARLEAVE(Text2.Text, "ML")
        X5 = GETLASTYEARLEAVE(Text2.Text, "PL")
        
    Else
        X1 = GETDAYSTYPE("SL")
        X2 = GETDAYSTYPE("VL")
        X2 = GETDAYSTYPE("EL")
        X2 = GETDAYSTYPE("ML")
        X2 = GETDAYSTYPE("PL")
    End If
    
    Text4.Text = X1:  Text5.Text = X2:  Text6.Text = X3:  Text7.Text = X4:
    Text8.Text = X5:
End Sub

Function GETLASTYEAR(EMPNO, XTYPE) As Integer
    Dim sqltxt As String
    Dim rsGETLEAVE As New ADODB.Recordset
    
    sqltxt = "Select LASTYEAR from hrms_leave where emplno = '" & EMPNO & "'"
    sqltxt = sqltxt & " and [type] = '" & XTYPE & "'"
    Set rsGETLEAVE = gconDMIS.Execute(sqltxt)
    
    If Not (rsGETLEAVE.BOF And rsGETLEAVE.EOF) Then
        If IsNull(Trim(rsGETLEAVE!LASTYEAR)) = True Then
            GETLASTYEAR = CInt(YEAR(Now))
        Else
            GETLASTYEAR = (Trim(rsGETLEAVE!LASTYEAR))
        End If
    End If
End Function

Function GETLASTYEARLEAVE(EMPNO, XTYPE) As Integer
    Dim sqltxt As String
    Dim rsGETLEAVE As New ADODB.Recordset
    
    sqltxt = "Select LASTYEARLEAVE from hrms_leave where emplno = '" & EMPNO & "'"
    sqltxt = sqltxt & " and [type] = '" & XTYPE & "'"
    Set rsGETLEAVE = gconDMIS.Execute(sqltxt)
    
    If Not (rsGETLEAVE.BOF And rsGETLEAVE.EOF) Then
        GETLASTYEARLEAVE = Trim(rsGETLEAVE!LASTYEARLEAVE)
    End If
End Function

Sub rsrefresh()
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE <> 'I' ORDER BY LASTNAME + ', ' + FIRSTNAME")
End Sub

Sub FillGrid()
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEmpInfo
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call YEAR_END
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Dim COMPAIRE As String
    Dim rsFIND As New ADODB.Recordset
    
    Text1.Text = ITEM.Text
    Text2.Text = ITEM.ListSubItems(1).Text
    Call SendToFront
        
    
    COMPAIRE = GETLASTYEAR(Text2.Text, "SL") + 1
    If COMPAIRE = CStr(YEAR(Date)) Then
        If GETLASTYEARLEAVE(Text2.Text, "SL") = GETAVTYPE(Trim(Text2.Text), "SL") Then
            Call RESTOREYLEAVE
            Call InitLeavesTaken
            Call cmdSave_Click
        Else
            Call StoreMemVars
            Call InitTextLastYearBalance
        End If
    Else
        Call StoreMemVars
        Call InitTextLastYearBalance
    End If
End Sub

Private Sub SendToFront()
    cmdEdit.Enabled = True
    'cmdDelete.Enabled = False
    'cmdPrevious.Enabled = True
    'cmdNext.Enabled = True
End Sub

Sub StoreMemVars()
'    On Error Resume Next
'    If Text2.Text <> "" Then
'        Dim rsTemp As ADODB.Recordset
'        Set rsTemp = New ADODB.Recordset
'        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVE WHERE EMPNO = '" & Text2.Text & "' AND YEAR_BALANCE = '" & NumericVal(cboYear.Text) - 1 & "'")
'        If Not rsTemp.EOF And Not rsTemp.BOF Then
'            Text4.Text = N2Str2Zero(rsTemp!SL)
'            Text5.Text = N2Str2Zero(rsTemp!VL)
'            Text6.Text = N2Str2Zero(rsTemp!EL)
'            Text7.Text = N2Str2Zero(rsTemp!ML)
'            Text8.Text = N2Str2Zero(rsTemp!PL)
'        Else
'            InitTextLastYearBalance
'        End If
'        Call ComputeLeaveTaken(Text2.Text, cboYear.Text)
'        Text13.Text = SL_NO
'        Text14.Text = VL_NO
'        Text15.Text = EL_NO
'        Text16.Text = ML_NO
'        Text17.Text = PL_NO
'    Else
'        InitMemvars
'    End If
'    Set rsTemp = Nothing

'Update BY: NVB
    If Text2.Text <> "" Then
        Dim MAXSL, MAXPL, MAXVL, MAXEL, MAXML As Integer
        Dim CHECK As Boolean
        
        CHECK = TRACEIF(Trim(Text2.Text)) ', GETLASTYEAR(Trim(Text2.Text), "SL"))
        
        If CHECK = True Then
        
            MAXSL = GETDAYSTYPE("SL")
            MAXPL = GETDAYSTYPE("PL")
            MAXVL = GETDAYSTYPE("VL")
            MAXEL = GETDAYSTYPE("EL")
            MAXML = GETDAYSTYPE("ML")
           
            Text3.Text = MAXSL: Text9.Text = MAXVL: Text10.Text = MAXEL
            Text11.Text = MAXML: Text12.Text = MAXPL
            
            Text13.Text = 0: Text14.Text = 0: Text15.Text = 0: Text16.Text = 0
            Text17.Text = 0
            
        ElseIf CHECK = False Then
            
            MAXSL = GETAVTYPE(Trim(Text2.Text), "SL") ', cboYear)
            MAXPL = GETAVTYPE(Trim(Text2.Text), "PL") ', cboYear)
            MAXVL = GETAVTYPE(Trim(Text2.Text), "VL") ', cboYear)
            MAXEL = GETAVTYPE(Trim(Text2.Text), "EL") ', cboYear)
            MAXML = GETAVTYPE(Trim(Text2.Text), "ML") ', cboYear)

            Text3.Text = MAXSL: Text9.Text = MAXVL: Text10.Text = MAXEL
            Text11.Text = MAXML: Text12.Text = MAXPL

            Text13.Text = GetUsed(Trim(Text2.Text), "SL") ', cboYear)
            Text14.Text = GetUsed(Trim(Text2.Text), "VL") ', cboYear)
            Text15.Text = GetUsed(Trim(Text2.Text), "EL") ', cboYear)
            Text16.Text = GetUsed(Trim(Text2.Text), "ML") ', cboYear)
            Text17.Text = GetUsed(Trim(Text2.Text), "PL") ', cboYear)
        
        End If
    End If
End Sub

Function GETAVTYPE(XEMPNO As String, TYPE_LEAVE As String) As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim sqltxt As String
        
    sqltxt = "Select available from HRMS_LEAVE where EMPLNO = '" & XEMPNO & "'"
    sqltxt = sqltxt & " and [type] = '" & TYPE_LEAVE & "'" ' and LASTYEAR = '" & XYEAR & "'"
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GETAVTYPE = Trim(RSTMP!Available)
    
    Else
        GETAVTYPE = GETDAYSTYPE(TYPE_LEAVE)
    End If
    
    Set RSTMP = Nothing
End Function

Function GetUsed(XEMPNO As String, TYPE_LEAVE As String) As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim sqltxt As String
        
    sqltxt = "Select used from HRMS_LEAVE where EMPLNO = '" & XEMPNO & "'"
    sqltxt = sqltxt & " and [type] = '" & TYPE_LEAVE & "'" ' and LASTYEAR = '" & (XYEAR) & "'"
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetUsed = Trim(RSTMP!used)
   End If
    
    Set RSTMP = Nothing
End Function

Function TRACEIF(XEMPNO As String) As Boolean
    Dim RSTMP As New ADODB.Recordset
    Dim sqltxt As String
    
    sqltxt = "Select * from HRMS_LEAVE where EMPLNO = '" & XEMPNO & "'" '  and LASTYEAR = '" & XYEAR & "'"
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        TRACEIF = False
    Else
        TRACEIF = True
    End If
    
    Set RSTMP = Nothing
End Function

Function GETDAYSTYPE(TYPE_LEAVE As String) As Integer
    Dim RSTMP As New ADODB.Recordset
    Dim sqltxt As String
    
    sqltxt = "Select days_no from hrms_leavemaster where LEAVE_CODE = '" & TYPE_LEAVE & "'"
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GETDAYSTYPE = NumericVal(RSTMP!DAYS_NO)
    End If
    
    Set RSTMP = Nothing
End Function

Sub InitMemvars()
    InitTextLastYearBalance
    InitTextThisYearBalance
    InitLeavesTaken
End Sub

Sub FillCombo()
    FillcboYear cboyear
End Sub

Sub InitTextLastYearBalance()
    Text4.Text = 0: Text4.Locked = True
    Text5.Text = 0: Text5.Locked = True
    Text6.Text = 0: Text6.Locked = True
    Text7.Text = 0: Text7.Locked = True
    Text8.Text = 0: Text8.Locked = True
End Sub

Sub DNInitTextLastYearBalance()
    Text4.Locked = False
    Text5.Locked = False
    Text6.Locked = False
    Text7.Locked = False
    Text8.Locked = False
End Sub

Sub InitTextThisYearBalance()
    Text3.Text = 0
    Text9.Text = 0
    Text10.Text = 0
    Text11.Text = 0
    Text12.Text = 0
End Sub

Sub InitLeavesTaken()
    Text13.Text = 0
    Text14.Text = 0
    Text15.Text = 0
    Text16.Text = 0
    Text17.Text = 0
End Sub

Sub ComputeLeaveTaken(EMPNO As String, LEAVE_YEAR As Integer)
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVEDET WHERE EMPNO = '" & EMPNO & "' AND YEAR(DATEFROM) = '" & LEAVE_YEAR & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            If Null2String(rsTemp!LEAVETYPE) = "SL" Then
                SL_NO = SL_NO + N2Str2Zero(rsTemp!DAYS_NO)
            ElseIf Null2String(rsTemp!LEAVETYPE) = "EL" Then
                EL_NO = EL_NO + N2Str2Zero(rsTemp!DAYS_NO)
            ElseIf Null2String(rsTemp!LEAVETYPE) = "VL" Then
                VL_NO = VL_NO + N2Str2Zero(rsTemp!DAYS_NO)
            ElseIf Null2String(rsTemp!LEAVETYPE) = "ML" Then
                ML_NO = ML_NO + N2Str2Zero(rsTemp!DAYS_NO)
            ElseIf Null2String(rsTemp!LEAVETYPE) = "PL" Then
                PL_NO = PL_NO + N2Str2Zero(rsTemp!DAYS_NO)
            End If
            rsTemp.MoveNext
        Wend
    Else
        InitLeavesTaken
    End If
End Sub

'Private Sub Text4_KeyPress(KeyAscii As Integer)
'     Select Case KeyAscii
'
'     Case 48 To 57
'     Case 8: StoreMemVars
'     Case Else: KeyAscii = 0 'do nothing
'     End Select
'End Sub

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Private Sub Search()
    Dim RSTMP As New ADODB.Recordset
    Dim sqltxt As String
    
    sqltxt = "SELECT LASTNAME + ', ' + FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE <> 'I'"
    sqltxt = sqltxt & " AND LASTNAME LIKE '" & txtSearch & "%'"
    lsAdjustment.ListItems.Clear
    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
       Listview_Loadval Me.lsAdjustment.ListItems, RSTMP
    End If
    Set RSTMP = Nothing
End Sub

Private Sub Text4_Change()
    Dim X As Integer
    Dim Y As Integer
    Dim z As Integer
    
    X = NumericVal(Text4.Text)
    Y = NumericVal(Text3.Text)
    z = GETAVTYPE(Trim(Text2.Text), "SL") ', cboyear)
    
    Y = X + z
    Text3.Text = Y
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 13
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub Text5_Change()
    Dim X As Integer
    Dim Y As Integer
    Dim z As Integer
    
    X = NumericVal(Text5.Text)
    Y = NumericVal(Text9.Text)
    z = GETAVTYPE(Trim(Text2.Text), "VL") ', cboyear)
    
    Y = X + z
    Text9.Text = Y
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 13
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub Text6_Change()
    Dim X As Integer
    Dim Y As Integer
    Dim z As Integer
    
    X = NumericVal(Text6.Text)
    Y = NumericVal(Text10.Text)
    z = GETAVTYPE(Trim(Text2.Text), "EL") ', cboyear)
    
    Y = X + z
    Text10.Text = Y
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 13
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub Text7_Change()
    Dim X As Integer
    Dim Y As Integer
    Dim z As Integer
    
    X = NumericVal(Text7.Text)
    Y = NumericVal(Text11.Text)
    z = GETAVTYPE(Trim(Text2.Text), "ML") ', cboyear)
    
    Y = X + z
    Text11.Text = Y
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 13
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub Text8_Change()
    Dim X As Integer
    Dim Y As Integer
    Dim z As Integer
    
    X = NumericVal(Text8.Text)
    Y = NumericVal(Text12.Text)
    z = GETAVTYPE(Trim(Text2.Text), "PL") ', cboyear)
    
    Y = X + z
    Text12.Text = Y
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 13
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtsearch_Change()
    If txtSearch.Text = "" Then
        Call rsrefresh
        Call FillGrid
    Else
        Call Search
    End If
End Sub

Private Sub TXTSEARCH_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsAdjustment.ListItems.count > 0 And lsAdjustment.Enabled = True Then lsAdjustment.SetFocus
    End If
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMSBeginingBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beginning Balance"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_BeginingBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   10950
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   2730
      TabIndex        =   34
      Top             =   30
      Width           =   8025
      Begin VB.ComboBox cboYEAR 
         Height          =   360
         Left            =   6660
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lblEmpName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   2460
         TabIndex        =   38
         Top             =   240
         Width           =   4005
      End
      Begin VB.Label lblEmpNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   1170
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   36
         Top             =   330
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5130
      ScaleHeight     =   855
      ScaleWidth      =   5760
      TabIndex        =   14
      Top             =   3840
      Width           =   5760
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
         Left            =   4920
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_BeginingBalance.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   4230
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_BeginingBalance.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
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
         Left            =   3540
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_BeginingBalance.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   2850
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":1377
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_BeginingBalance.frx":14C9
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
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
         Left            =   2160
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":1825
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_BeginingBalance.frx":1977
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   1470
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":1C8A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_BeginingBalance.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame fmeInfo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3045
      Left            =   2730
      TabIndex        =   8
      Top             =   630
      Width           =   8025
      Begin VB.TextBox txtERPHIC 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1680
         TabIndex        =   28
         Top             =   2490
         Width           =   2115
      End
      Begin VB.TextBox txtERPAGIBIG 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1680
         TabIndex        =   26
         Top             =   2100
         Width           =   2115
      End
      Begin VB.TextBox txtERSSS 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1680
         TabIndex        =   24
         Top             =   1680
         Width           =   2115
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1890
         TabIndex        =   0
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5700
         TabIndex        =   1
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txtEESSS 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5640
         TabIndex        =   2
         Top             =   1650
         Width           =   2115
      End
      Begin VB.TextBox txtEEPHIC 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5670
         TabIndex        =   4
         Top             =   2460
         Width           =   2115
      End
      Begin VB.TextBox txtEEPAGIBIG 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5670
         TabIndex        =   3
         Top             =   2070
         Width           =   2115
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   315
         Left            =   -30
         TabIndex        =   32
         Top             =   1200
         Width           =   3975
         _Version        =   655364
         _ExtentX        =   7011
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "EMPLOYER CONTRIBUTION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   3930
         TabIndex        =   31
         Top             =   1200
         Width           =   4065
         _Version        =   655364
         _ExtentX        =   7170
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "EMPLOYEE CONTRIBUTION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   315
         Left            =   0
         TabIndex        =   30
         Top             =   120
         Width           =   7995
         _Version        =   655364
         _ExtentX        =   14102
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "BEGINNING BALANCE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Philhealth"
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
         Height          =   240
         Index           =   9
         Left            =   210
         TabIndex        =   29
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Pag Ibig"
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
         Height          =   240
         Index           =   3
         Left            =   210
         TabIndex        =   27
         Top             =   2130
         Width           =   795
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS"
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
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   1740
         Width           =   405
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Withheld Tax"
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
         Height          =   240
         Index           =   4
         Left            =   4260
         TabIndex        =   13
         Top             =   630
         Width           =   1290
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Taxable"
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
         Height          =   240
         Index           =   5
         Left            =   570
         TabIndex        =   12
         Top             =   660
         Width           =   1185
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS"
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
         Height          =   240
         Index           =   6
         Left            =   5160
         TabIndex        =   11
         Top             =   1740
         Width           =   405
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Philhealth"
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
         Height          =   240
         Index           =   8
         Left            =   4470
         TabIndex        =   10
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Pag Ibig"
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
         Height          =   240
         Index           =   10
         Left            =   4770
         TabIndex        =   9
         Top             =   2160
         Width           =   795
      End
   End
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   4635
      Left            =   90
      ScaleHeight     =   4635
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   30
      Width           =   2595
      Begin VB.TextBox txtSearch 
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
         TabIndex        =   6
         Top             =   120
         Width           =   2475
      End
      Begin MSComctlLib.ListView lsvEmp 
         Height          =   3975
         Left            =   60
         TabIndex        =   7
         Top             =   600
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   7011
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
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":20D6
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
         Picture         =   "frmHRMS_BeginingBalance.frx":2238
      End
   End
   Begin Crystal.CrystalReport rptBB 
      Left            =   5400
      Top             =   4020
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
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9285
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   21
      Top             =   3840
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
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":15FA5
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_BeginingBalance.frx":160F7
         Style           =   1  'Graphical
         TabIndex        =   23
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
         MouseIcon       =   "frmHRMS_BeginingBalance.frx":16435
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_BeginingBalance.frx":16587
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label lblID 
      Caption         =   "ID"
      Height          =   285
      Left            =   2760
      TabIndex        =   33
      Top             =   4230
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmHRMSBeginingBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub InitMemvars()
    txtEEPHIC = ""
    txtEEPAGIBIG = ""
    txtEESSS = ""
    txtERPAGIBIG = ""
    txtERPHIC = ""
    txtERSSS = ""
    txtNet = ""
    txtTax = ""
End Sub

Sub EnabledPics(COND As Boolean)
    picSearch.Enabled = COND
    Picture1.Visible = COND
    Picture2.Visible = Not COND
    fmeInfo.Enabled = Not COND
End Sub

Sub StoreMemVars(VEMPNO As String)
    Dim RSTMP                                                         As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_BEGBALANCE WHERE EMPNO = '" & VEMPNO & "' AND BEGBALYEAR = '" & cboYear & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtEEPHIC = Null2String(RSTMP!EEPHIC)
        txtEEPAGIBIG = Null2String(RSTMP!EEPAGIBIG)
        txtEESSS = Null2String(RSTMP!EESSS)
        txtERPAGIBIG = Null2String(RSTMP!ERPAGIBIG)
        txtERPHIC = Null2String(RSTMP!ERPHIC)
        txtERSSS = Null2String(RSTMP!ERSSS)
        txtNet = Null2String(RSTMP!net)
        txtTax = Null2String(RSTMP!TAX)
    Else
        InitMemvars
    End If
    Set RSTMP = Nothing
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsvEmp.Sorted = False
    lsvEmp.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME, EMPNO FROM HRMS_EMPINFO ORDER BY LASTNAME + ', ' + FIRSTNAME ASC")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsvEmp.ListItems, rsEMPINFO2
        lsvEmp.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)

    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsvEmp.Sorted = False
    lsvEmp.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE LASTNAME + ', ' + FIRSTNAME LIKE '" & XXX & "%' ORDER BY LASTNAME + ', ' + FIRSTNAME ASC")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsvEmp.ListItems, rsEMPINFO2
        lsvEmp.Refresh
    End If
End Sub

Private Sub cboYEAR_Change()
    If lblEmpNo.Caption <> "" Then
        StoreMemVars (lblEmpNo.Caption)
    End If
End Sub

Private Sub cboyear_Click()
    If lblEmpNo.Caption <> "" Then
        StoreMemVars (lblEmpNo.Caption)
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "ACESS_ADD", "EMPLOYEE BEGINNING BALANCE") = False Then Exit Sub
    If Not lblEmpNo.Caption = "" Then
        If lblADD_EDIT.Caption = "ADD" Then
            EnabledPics (False)
            InitMemvars
        End If
    Else
        MsgBox "Choose an Employee First", vbInformation, "Beginning Balance"
        txtsearch.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    EnabledPics True
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "ACESS_EDIT", "EMPLOYEE BEGINNING BALANCE") = False Then Exit Sub
    EnabledPics False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtsearch.SetFocus
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Print", "EMPLOYEE BEGINNING BALANCE") = False Then Exit Sub
    Screen.MousePointer = 11
    rptBB.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptBB.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptBB.Formulas(2) = "COMPANYTIN = '" & COMPANY_TIN & "'"
    rptBB.Formulas(3) = "PRINTEDBY = '" & LOGNAME & "'"

    PrintSQLReport rptBB, HRMS_REPORT_PATH & "Beginning Balance.rpt", "{HRMS_BegBalance.empno} = '" & lblEmpNo.Caption & "'", DMIS_REPORT_Connection, 1
    LogAudit "V", "PRINT EMPLOYEE BEGGINING BALANCE", lblEmpNo.Caption
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    Dim VtxtEEPHIC                                              As String
    Dim VtxtEEPAGIBIG                                           As String
    Dim VtxtEESSS                                               As String
    Dim VtxtERPAGIBIG                                           As String
    Dim VtxtERPHIC                                              As String
    Dim VtxtERSSS                                               As String
    Dim VtxtNet                                                 As String
    Dim VtxtTax                                                 As String
    Dim VcboYEAR                                                As String

    VtxtEEPHIC = N2Str2Null(txtEEPHIC)
    VtxtEEPAGIBIG = N2Str2Null(txtEEPAGIBIG)
    VtxtEESSS = N2Str2Null(txtEESSS)
    VtxtERPAGIBIG = N2Str2Null(txtERPAGIBIG)
    VtxtERPHIC = N2Str2Null(txtERPHIC)
    VtxtERSSS = N2Str2Null(txtERSSS)
    VtxtNet = N2Str2Null(txtNet)
    VtxtTax = N2Str2Null(txtTax)
    VcboYEAR = N2Str2Null(cboYear)
    
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_BEGBALANCE WHERE EMPNO = '" & lblEmpNo & "' AND BEGBALYEAR = '" & cboYear & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        gconDMIS.Execute ("UPDATE HRMS_BEGBALANCE SET " & _
                         "EEPHIC = " & VtxtEEPHIC & _
                         ", EEPAGIBIG = " & VtxtEEPAGIBIG & _
                         ", EESSS = " & VtxtEESSS & _
                         ", ERPAGIBIG = " & VtxtERPAGIBIG & _
                         ", ERPHIC = " & VtxtERPHIC & _
                         ", ERSSS = " & VtxtERSSS & _
                         ", NET = " & VtxtNet & _
                         ", TAX = " & VtxtTax & _
                         " WHERE EMPNO = '" & lblEmpNo.Caption & "' AND BEGBALYEAR = '" & cboYear & "'")
    Else
        gconDMIS.Execute ("INSERT INTO HRMS_BEGBALANCE( EMPLEVEL , EMPNO, EEPHIC, EEPAGIBIG, EESSS, ERPAGIBIG, ERPHIC, ERSSS, NET, TAX, BEGBALYEAR)" & _
                  " VALUES(" & _
                  "'" & GetEmployeeLevel(lblEmpNo.Caption) & "'" & _
                  ", '" & lblEmpNo.Caption & "'" & _
                  ", " & VtxtEEPHIC & _
                  ", " & VtxtEEPAGIBIG & _
                  ", " & VtxtEESSS & _
                  ", " & VtxtERPAGIBIG & _
                  ", " & VtxtERPHIC & _
                  ", " & VtxtERSSS & _
                  ", " & VtxtNet & _
                  ", " & VtxtTax & _
                  ", " & VcboYEAR & ")")
    End If
    LogAudit "E", "UPDATE EMPLOYEE BEGGINING BALANCE", lblEmpNo.Caption
    ShowSuccessFullyUpdated
    cmdCancel.Value = True
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    DrawXPCtl Me
    'FillcboYear cboYear
    fillcombo_up cboYear
    FillGrid
End Sub

Private Sub lsvEmp_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Dim Index                                                         As Double

    If Not lsvEmp.ListItems.count = 0 Then
        With lsvEmp
            Index = .SelectedItem.Index
            lblEmpName.Caption = .ListItems(Index).Text
            lblEmpNo.Caption = .ListItems(Index).SubItems(1)
            
            InitMemvars
            StoreMemVars (lblEmpNo.Caption)
        End With
    End If
End Sub

Private Sub txtsearch_Change()
    If Trim(txtsearch.Text) = "" Then FillGrid Else FillSearchGrid txtsearch.Text
End Sub

Function GetEmployeeLevel(EMPNO As String) As String
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT EMPLEVEL FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    
    GetEmployeeLevel = "E"
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GetEmployeeLevel = Null2String(rsTemp!EMPLEVEL)
    End If
End Function


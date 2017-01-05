VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmHRMSSalaryAdvance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salary Advances Entry"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8325
   ForeColor       =   &H00D8E9EC&
   Icon            =   "SalaryAdvance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8325
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2655
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   33
      Top             =   4905
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
         MouseIcon       =   "SalaryAdvance.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Exit Window"
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
         MouseIcon       =   "SalaryAdvance.frx":07C2
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   39
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
         MouseIcon       =   "SalaryAdvance.frx":0C3F
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":0D91
         Style           =   1  'Graphical
         TabIndex        =   38
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
         MouseIcon       =   "SalaryAdvance.frx":10ED
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":123F
         Style           =   1  'Graphical
         TabIndex        =   37
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
         MouseIcon       =   "SalaryAdvance.frx":1552
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":16A4
         Style           =   1  'Graphical
         TabIndex        =   36
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
         MouseIcon       =   "SalaryAdvance.frx":199E
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":1AF0
         Style           =   1  'Graphical
         TabIndex        =   35
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
         MouseIcon       =   "SalaryAdvance.frx":1E48
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":1F9A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Move to Previous Record"
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
         MouseIcon       =   "SalaryAdvance.frx":22F9
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":244B
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   3345
      Left            =   2640
      ScaleHeight     =   3345
      ScaleWidth      =   5655
      TabIndex        =   27
      Top             =   1515
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid grdSalaryAdvance 
         Height          =   3195
         Left            =   60
         TabIndex        =   28
         Top             =   60
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   5636
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   0
         BackColorSel    =   -2147483633
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         Appearance      =   0
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
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   2640
      ScaleHeight     =   1365
      ScaleWidth      =   5655
      TabIndex        =   19
      Top             =   135
      Width           =   5655
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
         Left            =   900
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   480
         Width           =   2895
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
         Left            =   60
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   60
         Width           =   3735
      End
      Begin VB.TextBox txtYTDSalaryAdvance 
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
         Left            =   3840
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   480
         Width           =   1725
      End
      Begin VB.TextBox txtSABalance 
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
         Left            =   3840
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   900
         Width           =   1725
      End
      Begin Crystal.CrystalReport rptSalaryAdvance 
         Left            =   480
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
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
         Left            =   60
         TabIndex        =   26
         Top             =   510
         Width           =   825
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Cash Advance"
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
         Left            =   3840
         TabIndex        =   25
         Top             =   180
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Advance Balance"
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
         Left            =   1530
         TabIndex        =   24
         Top             =   960
         Width           =   2325
      End
   End
   Begin VB.Frame fraSalaryAdvance 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      Caption         =   "Add Cash Advance"
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
      Height          =   2745
      Left            =   2790
      TabIndex        =   7
      Top             =   1785
      Visible         =   0   'False
      Width           =   5265
      Begin VB.ComboBox cboDay0 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   570
         Width           =   945
      End
      Begin VB.ComboBox cboMonth0 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   570
         Width           =   1545
      End
      Begin VB.ComboBox cboYear0 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4260
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   570
         Width           =   825
      End
      Begin VB.ComboBox cboDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1350
         Width           =   945
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1350
         Width           =   1545
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4260
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1350
         Width           =   825
      End
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   555
         Left            =   1590
         TabIndex        =   6
         Top             =   2070
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   979
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5220
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Deduction in Payroll"
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
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1050
         Width           =   2955
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Height          =   285
         Left            =   60
         TabIndex        =   17
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Height          =   285
         Left            =   1530
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Height          =   285
         Left            =   3750
         TabIndex        =   15
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Height          =   285
         Left            =   60
         TabIndex        =   14
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Height          =   285
         Left            =   1530
         TabIndex        =   13
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Height          =   285
         Left            =   3750
         TabIndex        =   12
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Cash Advance"
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
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Advance Amount"
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
         Left            =   1530
         TabIndex        =   9
         Top             =   1800
         Width           =   2025
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         ForeColor       =   &H8000000D&
         Height          =   15
         Left            =   30
         TabIndex        =   8
         Top             =   330
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin wizButton.cmd cmdSalaryAdvance 
      Height          =   2865
      Left            =   2730
      TabIndex        =   10
      Top             =   1725
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   5054
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
      MICON           =   "SalaryAdvance.frx":27B1
   End
   Begin VB.PictureBox Picture11 
      Height          =   5625
      Left            =   60
      Picture         =   "SalaryAdvance.frx":27CD
      ScaleHeight     =   5565
      ScaleWidth      =   2445
      TabIndex        =   29
      Top             =   135
      Width           =   2505
   End
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   5625
      Left            =   60
      ScaleHeight     =   5625
      ScaleWidth      =   2505
      TabIndex        =   30
      Top             =   135
      Width           =   2505
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
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   45
         Width           =   2415
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   5085
         Left            =   0
         TabIndex        =   32
         Top             =   450
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   8969
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
         Appearance      =   1
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
         MouseIcon       =   "SalaryAdvance.frx":1652A
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
         Picture         =   "SalaryAdvance.frx":1668C
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6795
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   42
      Top             =   4890
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
         MouseIcon       =   "SalaryAdvance.frx":2A3F9
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":2A54B
         Style           =   1  'Graphical
         TabIndex        =   44
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
         MouseIcon       =   "SalaryAdvance.frx":2A889
         MousePointer    =   99  'Custom
         Picture         =   "SalaryAdvance.frx":2A9DB
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSSalaryAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSEMPINFO, rsSalaryAdvance, rsSalaryGrade As ADODB.Recordset
Attribute rsSalaryAdvance.VB_VarUserMemId = 1073938432
Attribute rsSalaryGrade.VB_VarUserMemId = 1073938432
Dim AddorEdit, Diyt, Diyt2                   As String
Attribute AddorEdit.VB_VarUserMemId = 1073938435
Attribute Diyt.VB_VarUserMemId = 1073938435
Attribute Diyt2.VB_VarUserMemId = 1073938435
Dim Obertaym, Halidey                        As Double
Attribute Obertaym.VB_VarUserMemId = 1073938438
Attribute Halidey.VB_VarUserMemId = 1073938438

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Add", "SALARY ADVANCES") = False Then Exit Sub
    AddorEdit = "ADD"
    fraSalaryAdvance.Caption = "Add Cash Advance"
    fraSalaryAdvance.Visible = True
    cmdSalaryAdvance.Visible = True
    fraSalaryAdvance.Enabled = True
    cmdSalaryAdvance.ZOrder 0
    fraSalaryAdvance.ZOrder 0
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    fraSalaryAdvance.Caption = ""
    fraSalaryAdvance.Visible = False
    cmdSalaryAdvance.Visible = False
    fraSalaryAdvance.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    cmdSalaryAdvance.ZOrder 1
    fraSalaryAdvance.ZOrder 1
    storeMemvars
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Delete", "SALARY ADVACES") = False Then Exit Sub
    grdSalaryAdvance.Col = 3
    If grdSalaryAdvance.Text <> "" Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "DELETE FROM HRMS_SALARYADVANCE WHERE ID = " & grdSalaryAdvance.Text
            LogAudit "X", "SALARY ADVANCE", EMPLOYEE_NO
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    storeMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Edit", "SALARY ADVACES") = False Then Exit Sub
    Dim fild                                 As String
    grdSalaryAdvance.Row = grdSalaryAdvance.Row
    grdSalaryAdvance.Col = 3
    fild = grdSalaryAdvance.Text
    If fild <> "" Then
        AddorEdit = "EDIT"
        fraSalaryAdvance.Caption = "Edit Cash Advance"
        cmdSalaryAdvance.Visible = True
        cmdSalaryAdvance.ZOrder 0
        fraSalaryAdvance.Visible = True
        fraSalaryAdvance.ZOrder 0
        fraSalaryAdvance.Enabled = True
        Picture1.Visible = False
        Picture2.Visible = True
        StoreEntry (fild)
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Function StoreEntry(ByVal ID As Variant)
    Dim MM                          As String
    Dim DD                          As String
    Dim YY                          As String
    Dim TheDeyt                     As String
    Dim TheDeyt2                    As String
    Set rsSalaryAdvance = New ADODB.Recordset
    rsSalaryAdvance.Open "SELECT * FROM HRMS_SALARYADVANCE WHERE ID = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalaryAdvance.EOF And Not rsSalaryAdvance.BOF Then
        labID.Caption = rsSalaryAdvance!ID
        TheDeyt = Null2Date(rsSalaryAdvance!DEYT)
        DD = Day(TheDeyt)
        MM = The_month(Month(TheDeyt))
        YY = Year(TheDeyt)
        cboDay.Text = DD
        cboMonth.Text = MM
        cboYear.Text = YY
        txtAmount.Text = N2Str2Zero(rsSalaryAdvance!AMOUNT)
    End If
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    picSearch.ZOrder 0
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    RSEMPINFO.MoveNext
    If RSEMPINFO.EOF Then
        RSEMPINFO.MoveLast
        ShowLastRecordMsg
    End If
    storeMemvars
End Sub

Private Sub cmdPrevious_Click()
    RSEMPINFO.MovePrevious
    If RSEMPINFO.BOF Then
        RSEMPINFO.MoveFirst
        ShowFirstRecordMsg
    End If
    storeMemvars
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Print", "SALARY ADVACES") = False Then Exit Sub
    Screen.MousePointer = 11
    rptSalaryAdvance.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptSalaryAdvance.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptSalaryAdvance.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    PrintSQLReport rptSalaryAdvance, HRMS_REPORT_PATH & "SALARYADVANCE.RPT", "{SALARYADVANCE.EMPNO} = " & N2Str2Null(RSEMPINFO!EMPNO), DMIS_REPORT_Connection, 1
    Call LogAudit("V", "SALARY ADVANCE", EMPLOYEE_NO)
    Screen.MousePointer = 0
Exit Sub
Errorcode:
ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim MM                          As String
    Dim DD                          As String
    Dim YY                          As String
    Dim MM0                         As String
    Dim DD0                         As String
    Dim YY0                         As String
    Dim SalaryAdvanceCode           As String
    MM = What_month(cboMonth)
    YY = cboYear.Text
    DD = cboDay.Text
    MM0 = What_month(cboMonth0)
    YY0 = cboYear0.Text
    DD0 = cboDay0.Text
    Diyt = DateSerial(YY, MM, DD)
    Diyt2 = DateSerial(YY0, MM0, DD0)
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "INSERT INTO HRMS_SALARYADVANCE " & _
                         "(EMPNO,DATEOFSALARYADVANCE,DEYT,AMOUNT) " & _
                         "VALUES (" & N2Str2Null(RSEMPINFO!EMPNO) & ", " & N2Date2Null(Diyt2) & ", " & N2Date2Null(Diyt) & _
                         ", " & NumericVal(txtAmount.Text) & ")"
    Else
        gconDMIS.Execute "UPDATE HRMS_SALARYADVANCE SET" & _
                       " EMPNO = " & N2Str2Null(RSEMPINFO!EMPNO) & "," & _
                       " DATEOFSALARYADVANCE = " & N2Date2Null(Diyt2) & "," & _
                       " DEYT = " & N2Date2Null(Diyt) & "," & _
                       " AMOUNT = " & NumericVal(txtAmount.Text) & _
                       " WHERE ID = " & labID.Caption
    End If
    cmdCancel.Value = True
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
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
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsrefresh
    txtSearch.Text = ""
    InitGrid
    InitMemVars
    cmdCancel_Click
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set RSEMPINFO = New ADODB.Recordset
        RSEMPINFO.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & EmpInfoEmpno.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf HEADEMPINFOSHOW = True Then
        Set RSEMPINFO = New ADODB.Recordset
        RSEMPINFO.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & frmHRMSEmpInfo.labID.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set RSEMPINFO = New ADODB.Recordset
        RSEMPINFO.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'E' AND RESIGNED IS NULL ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
End Sub

Sub InitGrid()
    With grdSalaryAdvance
        .Rows = 2
        .ColWidth(0) = 1500
        .ColWidth(1) = 2200
        .ColWidth(2) = 1500
        .ColWidth(3) = 1
        .Row = 0
        .Col = 0
        .Text = "Date of SA"
        .Col = 1
        .Text = "Date of Deduction"
        .Col = 2
        .Text = "Amount"
        .Col = 3
        .Text = "ID"
    End With
End Sub

Sub InitMemVars()
    'fillcboDay cboDay
    'fillcbomonth cboMonth
    'FillcboYear cboYear
    'fillcboDay cboDay0
    'fillcbomonth cboMonth0
    'FillcboYear cboYear0
    'cboYear0.Text = Year(LOGDATE)
    'cboMonth0.Text = The_month(Month(LOGDATE))
    'cboDay0.Text = Day(LOGDATE)
    'cboYear.Text = Year(LOGDATE)
    'cboMonth.Text = The_month(Month(LOGDATE))
    'If Day(LOGDATE) > 15 Then
    '    cboDay.Text = Day(lastDay(LOGDATE))
    'Else
    '    cboDay.Text = 15
    'End If
    'txtAmount.Text = 0
    
End Sub

Sub storeMemvars()
    On Error GoTo Errorcode
    Dim CNT                                  As Integer
    Dim VYTDSalaryAdvance                    As Double
    Dim rsPAYROLL                            As ADODB.Recordset
    Dim TotSA                                As Double
    Set rsPAYROLL = New ADODB.Recordset
    rsPAYROLL.Open "SELECT SUM(SALARYADVANCE) AS TOTALSADED FROM HRMS_PAYROLL WHERE EMPNO = " & N2Str2Null(RSEMPINFO!EMPNO), gconDMIS
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        TotSA = N2Str2Zero(rsPAYROLL!TOTALSADED)
    End If
    If Not RSEMPINFO.EOF And Not RSEMPINFO.BOF Then
        Set rsSalaryAdvance = New ADODB.Recordset
        rsSalaryAdvance.Open "SELECT * FROM HRMS_SALARYADVANCE WHERE EMPNO = " & N2Str2Null(RSEMPINFO!EMPNO) & " ORDER BY DEYT DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        CNT = 0
        VYTDSalaryAdvance = 0
        If Not rsSalaryAdvance.EOF And Not rsSalaryAdvance.BOF Then
            rsSalaryAdvance.MoveFirst
            cleargrid grdSalaryAdvance
            grdSalaryAdvance.Rows = grdSalaryAdvance.Rows
            Do While Not rsSalaryAdvance.EOF
                CNT = CNT + 1
                labID.Caption = rsSalaryAdvance!ID
                grdSalaryAdvance.AddItem Null2Date(rsSalaryAdvance!dateofSalaryAdvance) & Chr(9) & Null2Date(rsSalaryAdvance!DEYT) & Chr(9) & N2Str2Zero(rsSalaryAdvance!AMOUNT) & Chr(9) & rsSalaryAdvance!ID
                If Year(Null2Date(rsSalaryAdvance!DEYT)) = Year(LOGDATE) Then
                    VYTDSalaryAdvance = VYTDSalaryAdvance + N2Str2Zero(rsSalaryAdvance!AMOUNT)
                End If
                rsSalaryAdvance.MoveNext
            Loop
            grdSalaryAdvance.RemoveItem 1
        Else
            cleargrid grdSalaryAdvance
        End If
        txtYTDSalaryAdvance.Text = Format(N2Str2Zero(VYTDSalaryAdvance), MAXIMUM_DIGIT)
        txtPosition.Text = Null2String(RSEMPINFO!Position)
        txtName.Text = Cap1st(Null2String(RSEMPINFO!lastname)) & ", " & Cap1st(Null2String(RSEMPINFO!FIRSTNAME)) & " " & Cap1st(Null2String(RSEMPINFO!MIDDLENAME))
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
    txtSABalance.Text = VYTDSalaryAdvance - TotSA
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub grdSalaryAdvance_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub lsAdjustment_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RSEMPINFO.Bookmark = rsFind(RSEMPINFO.Clone, "EMPNO", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
    storeMemvars
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

Private Sub TXTSEARCH_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                           As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = 'E' AND RESIGNED IS NULL ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
XXX = Repleys(XXX)
    Dim rsEMPINFO2                           As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = 'E' AND RESIGNED IS NULL AND LASTNAME+', '+FIRSTNAME LIKE'" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub


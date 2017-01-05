VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmPMIS_Physical_AddPhyCntTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Physical Count Ticket"
   ClientHeight    =   6780
   ClientLeft      =   1125
   ClientTop       =   435
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AddPhyCntTicket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11520
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7080
      ScaleHeight     =   855
      ScaleWidth      =   5325
      TabIndex        =   30
      Top             =   5790
      Width           =   5325
      Begin VB.PictureBox picMatAdjust 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   720
         ScaleHeight     =   870
         ScaleWidth      =   3630
         TabIndex        =   31
         Top             =   0
         Width           =   3630
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
            Left            =   2880
            MouseIcon       =   "AddPhyCntTicket.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "AddPhyCntTicket.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Exit Window"
            Top             =   0
            Width           =   735
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
            Left            =   2160
            MouseIcon       =   "AddPhyCntTicket.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "AddPhyCntTicket.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Print Record"
            Top             =   0
            Width           =   735
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
            Left            =   1440
            MouseIcon       =   "AddPhyCntTicket.frx":123A
            MousePointer    =   99  'Custom
            Picture         =   "AddPhyCntTicket.frx":138C
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Delete Selected Record"
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "Edit"
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
            MouseIcon       =   "AddPhyCntTicket.frx":16B7
            MousePointer    =   99  'Custom
            Picture         =   "AddPhyCntTicket.frx":1809
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Edit Selected Record"
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
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
            Left            =   0
            MouseIcon       =   "AddPhyCntTicket.frx":1C61
            MousePointer    =   99  'Custom
            Picture         =   "AddPhyCntTicket.frx":1DB3
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Add Record"
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
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
         Left            =   0
         MouseIcon       =   "AddPhyCntTicket.frx":20C6
         MousePointer    =   99  'Custom
         Picture         =   "AddPhyCntTicket.frx":2218
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Refresh Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Location"
      Height          =   240
      Left            =   3060
      TabIndex        =   29
      Top             =   90
      Width           =   1425
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Part No"
      Height          =   240
      Left            =   1590
      TabIndex        =   27
      Top             =   90
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tag No"
      Height          =   240
      Left            =   120
      TabIndex        =   26
      Top             =   90
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.TextBox txtSearchPartNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4590
      TabIndex        =   0
      Top             =   60
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   420
      Top             =   4770
   End
   Begin VB.PictureBox picPhyCnt2 
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
      Height          =   3495
      Left            =   2130
      ScaleHeight     =   3465
      ScaleWidth      =   7065
      TabIndex        =   11
      Top             =   1260
      Width           =   7095
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         Height          =   285
         Left            =   6780
         TabIndex        =   39
         Top             =   30
         Width           =   255
      End
      Begin VB.ComboBox cboPartNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         TabIndex        =   28
         Top             =   540
         Width           =   2055
      End
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
         Left            =   6120
         MouseIcon       =   "AddPhyCntTicket.frx":2670
         MousePointer    =   99  'Custom
         Picture         =   "AddPhyCntTicket.frx":27C2
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cancel Entry"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtMAC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "Text1"
         ToolTipText     =   "Type the average cost. Do not include comma or peso sign (e.g. 500)"
         Top             =   2070
         Width           =   1965
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "Text"
         Top             =   2100
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtGroup_No 
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
         Height          =   345
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "Type the employee number who performed the physical counting."
         Top             =   1710
         Width           =   1755
      End
      Begin VB.ComboBox cboAmark 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4950
         TabIndex        =   5
         Text            =   "cboAmark"
         ToolTipText     =   "Select from the list."
         Top             =   1320
         Width           =   1980
      End
      Begin VB.TextBox txtLocation 
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
         Height          =   345
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Type the location (e.g. 3RD FLOOR)"
         Top             =   1320
         Width           =   1755
      End
      Begin VB.TextBox txtQCount 
         Alignment       =   2  'Center
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
         Height          =   345
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "Text"
         ToolTipText     =   "Type the quantity counted (e.g. 50, 55)"
         Top             =   930
         Width           =   1755
      End
      Begin VB.TextBox txtAdate 
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
         Height          =   345
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Type the date in mm/dd/yyyy format (e.g. 7/5/2004)"
         Top             =   930
         Width           =   1965
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   4950
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "Text1"
         ToolTipText     =   "Input the status of the ticket (e.g. U for Unposted)"
         Top             =   1680
         Width           =   1965
      End
      Begin VB.TextBox txtTagNo 
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
         Height          =   345
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Type the Tag Number (e.g. 10, 20)"
         Top             =   540
         Width           =   1755
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
         Left            =   5400
         MouseIcon       =   "AddPhyCntTicket.frx":2B00
         MousePointer    =   99  'Custom
         Picture         =   "AddPhyCntTicket.frx":2C52
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Save Entry"
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label labdetid 
         Caption         =   "0"
         Height          =   285
         Left            =   660
         TabIndex        =   41
         Top             =   3480
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label labeeror 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Proof in Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   90
         TabIndex        =   40
         Top             =   2580
         Width           =   5265
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   7065
         _Version        =   655364
         _ExtentX        =   12462
         _ExtentY        =   661
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label labPartDesc 
         BackColor       =   &H8000000D&
         Caption         =   "Part Number"
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
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         Top             =   570
         Width           =   645
      End
      Begin VB.Label labPhyCntStatus 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Proof in Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   60
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Average Cost"
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
         Height          =   255
         Left            =   3540
         TabIndex        =   21
         Top             =   2100
         Width           =   1545
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Computer QTY"
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2130
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Checked By"
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
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1740
         Width           =   1545
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1350
         Width           =   1545
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Certified"
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
         Height          =   255
         Left            =   3540
         TabIndex        =   17
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "QTY Counted"
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
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Counted"
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
         Height          =   255
         Left            =   3540
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Height          =   255
         Left            =   3540
         TabIndex        =   14
         Top             =   1740
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
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
         Height          =   255
         Left            =   3510
         TabIndex        =   13
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tag No."
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
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   570
         Width           =   1545
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdPhyCnt 
      Height          =   5145
      Left            =   60
      TabIndex        =   10
      Top             =   525
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   9075
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      TextStyleFixed  =   3
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "AddPhyCntTicket.frx":2FA2
   End
End
Attribute VB_Name = "frmPMIS_Physical_AddPhyCntTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPHYCNT                                                          As ADODB.Recordset
Dim ADDOREDIT                                                         As String
Dim Kawnter_Ini                                                       As Integer

Sub rsRefresh()
    Set rsPHYCNT = New ADODB.Recordset
    rsPHYCNT.Open "Select * from PHYCNT  order by tagno asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitGrid()
    With grdPhyCnt
        .Row = 0
        .FormatString = "TAG NO.      | Part Number      | Part Description     | Location         | Computer QTY | QTY Counted  | " & _
                        "Variance | Average Cost | Total Cost | Date Acknowledged | Checked By | Print Status | Status | Last Update | Time Update | User Code |id"
    End With
End Sub

Sub FillGrid()
    Dim kcnt                                                          As Integer
    kcnt = 0
    Kawnter_Ini = 0
    If Not rsPHYCNT.EOF And Not rsPHYCNT.BOF Then
        Screen.MousePointer = 11
        rsPHYCNT.MoveFirst
        Do While Not rsPHYCNT.EOF
            kcnt = kcnt + 1
            Kawnter_Ini = Kawnter_Ini + 1
            grdPhyCnt.AddItem Format(Null2String(rsPHYCNT!tagno), "0000000000") & Chr(9) & _
                              Null2String(rsPHYCNT!PARTNO) & Chr(9) & _
                              Null2String(rsPHYCNT!PARTDESC) & Chr(9) & _
                              Null2String(rsPHYCNT!Location) & Chr(9) & _
                              Null2String(rsPHYCNT!ONHAND) & Chr(9) & _
                              Null2String(rsPHYCNT!Qcount) & Chr(9) & _
                              Null2String(rsPHYCNT!VARIANCE) & Chr(9) & _
                              Null2String(rsPHYCNT!Mac) & Chr(9) & _
                              Null2String(rsPHYCNT!totalmac) & Chr(9) & _
                              Null2String(rsPHYCNT!ADate) & Chr(9) & _
                              Null2String(rsPHYCNT!Amark) & Chr(9) & _
                              Null2String(rsPHYCNT!Print_Stat) & Chr(9) & _
                              Null2String(rsPHYCNT!STATUS) & Chr(9) & _
                              Null2String(rsPHYCNT!LASTUPDATE) & Chr(9) & _
                              Null2String(rsPHYCNT![Time]) & Chr(9) & _
                              Null2String(rsPHYCNT!USERCODE) & Chr(9) & Null2String(rsPHYCNT!ID)
            rsPHYCNT.MoveNext
        Loop
        If kcnt <> 0 Then grdPhyCnt.RemoveItem 1
        Screen.MousePointer = 0
    End If
    grdPhyCnt.TopRow = grdPhyCnt.Rows - 1
End Sub

Sub initMemvars()
labeeror = ""
txtTagNo.Text = ""
    cboPartNo.Text = ""
    txtQCount.Text = 0
    txtAdate.Text = LOGDATE
    txtLocation.Text = ""
    cboAmark.Clear
    cboAmark.AddItem ""
    cboAmark.AddItem "Y"
    cboAmark.AddItem "N"
    cboAmark.Text = "Y"
    txtGroup_No.Text = ""
    txtStatus.Text = ""
    txtOnHand.Text = 0
    txtMAC.Text = 0
    
End Sub

Sub StoreMemvars()
    grdPhyCnt.Row = grdPhyCnt.Row
    grdPhyCnt.Col = 0
    Set rsPHYCNT = New ADODB.Recordset
    
    rsPHYCNT.Open "Select * from PHYCNT where ID = " & grdPhyCnt.TextMatrix(grdPhyCnt.Row, 16), gconINVENTORY, adOpenForwardOnly, adLockReadOnly
    If Not rsPHYCNT.EOF And Not rsPHYCNT.BOF Then
        txtTagNo.Text = Null2String(Format(rsPHYCNT!tagno, "0000000000"))
        cboPartNo.Text = Null2String(rsPHYCNT!PARTNO)
        labPartDesc.Caption = Null2String(rsPHYCNT!PARTDESC)
        txtQCount.Text = N2Str2Zero(rsPHYCNT!Qcount)
        txtAdate.Text = Null2Date(rsPHYCNT!ADate)
        cboAmark.Text = Null2String(rsPHYCNT!Amark)
        txtGroup_No.Text = Null2String(rsPHYCNT!Group_No)
        txtStatus.Text = Null2String(rsPHYCNT!STATUS)
        txtOnHand.Text = N2Str2Zero(rsPHYCNT!ONHAND)
        txtMAC.Text = N2Str2Zero(rsPHYCNT!Mac)
        txtLocation.Text = Null2String(rsPHYCNT!Location)
        labDetID = Null2String(rsPHYCNT!ID)
    End If
End Sub

Private Sub cboPartNo_Click()
If ADDOREDIT = "" Then Exit Sub
    labeeror = ""
    If cboPartNo.Text = "" Then Exit Sub
    
    Dim rsCUTOFF                                                      As ADODB.Recordset
    On Error Resume Next
     
    Set rsCUTOFF = New ADODB.Recordset
    rsCUTOFF.Open "Select onhand,PARTNO,PARTDESC,mac,location from CUTOFF where PARTNO=" & N2Str2Null(Repleys(cboPartNo)), gconINVENTORY
    If Not rsCUTOFF.EOF And Not rsCUTOFF.BOF Then
        txtOnHand.Text = N2Str2Zero(rsCUTOFF!ONHAND)
        txtMAC.Text = N2Str2Zero(rsCUTOFF!Mac)
        txtLocation.Text = Null2String(rsCUTOFF!Location)
        labPartDesc.Caption = Null2String(rsCUTOFF!PARTDESC)
        cmdSave.Enabled = True
        'DoEvents
    
    Else
        labeeror = "Error: This Part number " & cboPartNo.Text & " doesn't exist in Cut Off Master File."
        cmdSave.Enabled = False
'        On Error Resume Next
'        cboPartNo.SetFocus
    End If
End Sub

Private Sub cmdAdd_Click()
    ADDOREDIT = "ADD"
    picPhyCnt2.ZOrder 0: picPhyCnt2.Visible = True
    txtTagNo.Enabled = True
    Picture1.Enabled = False
    initMemvars
    
    On Error Resume Next
    txtTagNo.Enabled = True
    txtTagNo.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    ADDOREDIT = ""
    initMemvars
    picPhyCnt2.ZOrder 1
    picPhyCnt2.Visible = False
    Picture1.Enabled = True
End Sub

Private Sub cmdChange_Click()
 grdPhyCnt_DblClick
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ERRORCODE:

    If grdPhyCnt.Text <> "No Entry" Or grdPhyCnt.Text <> "TAG NO." Then
        If ShowConfirmDelete = True Then
            grdPhyCnt.Col = 0
            '**************
            'updating code:     jaa - 11102008
            'SQL_STATEMENT = "Delete * from PHYCNT Where tagno = '" & NumericVal(grdPhyCnt.Text) & "'"
            SQL_STATEMENT = "Delete * from PHYCNT Where tagno = " & NumericVal(grdPhyCnt.Text)
            gconINVENTORY.Execute SQL_STATEMENT
            NEW_LogAudit "X", "PHYSICAL COUNT", SQL_STATEMENT, "", "", grdPhyCnt.Text, "", ""
            
            gconINVENTORY.Execute "update tags set" & _
                                " PARTNO = NULL" & _
                                " where tag = " & NumericVal(grdPhyCnt.Text)
                                '" where tag = '" & grdPhyCnt.Text & "'"
            grdPhyCnt.Col = 1
            '**************
            gconINVENTORY.Execute "update CUTOFF set" & _
                                " TAGNO = NULL" & _
                                " where PARTNO = '" & grdPhyCnt.Text & "'"
            cleargrid grdPhyCnt
            rsRefresh
            InitGrid
            FillGrid
        End If
    End If

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub cmdF6_Click()
    Unload Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    FlexGrid_To_Excel grdPhyCnt, grdPhyCnt.Rows, grdPhyCnt.Cols, 5, "TAG DATA ENTRY"
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    
    
    'Update by :NVB
    'Date updated: 1/27/2009
    'Description: User validation input
    
    Dim rsPHYCNTDUP                                    As ADODB.Recordset
    Dim rsTAGS                                         As ADODB.Recordset
    Dim rsPHYCNTDUP2 As New ADODB.Recordset
    Dim rsPartsCount As New ADODB.Recordset
        
    Set rsPHYCNTDUP = New ADODB.Recordset
    Set rsPHYCNTDUP2 = New ADODB.Recordset
    Set rsPartsCount = New ADODB.Recordset
    
    '***************
    'Update by :JBF
    'Date updated: 1/27/2009
    'Description: User validation input
    On Error GoTo ERRORCODE
    
    
    If txtTagNo.Text = "" Then
        MsgSpeechBox "Error: Tag Number must not be empty"
        On Error Resume Next
        txtTagNo.SetFocus
        Exit Sub
    End If
    
    '**************
    
    'Three times of recordset
    
    rsPHYCNTDUP.Open "Select tagno,PARTNO from PHYCNT where tagno = " & NumericVal(txtTagNo.Text) & " and PARTNO <> " & N2Str2Null(cboPartNo.Text), gconINVENTORY, adOpenForwardOnly, adLockReadOnly
    rsPHYCNTDUP2.Open "Select tagno,PARTNO from PHYCNT where tagno = " & NumericVal(txtTagNo.Text) & " and PARTNO = " & N2Str2Null(cboPartNo.Text), gconINVENTORY, adOpenForwardOnly, adLockReadOnly
    rsPartsCount.Open "Select PartNo,tagno from PHYCNT where PartNo = " & N2Str2Null(cboPartNo.Text), gconINVENTORY, adOpenForwardOnly, adLockReadOnly
        
    If ADDOREDIT = "ADD" Then 'this validation is for adding data
        
        
            'rsPHYCNTDUP.Open "Select tagno,PARTNO from PHYCNT where tagno = '" & NumericVal(txtTagNo.Text) & "' and PARTNO <> " & N2Str2Null(cboPartNo.Text), gconINVENTORY, adOpenForwardOnly, adLockReadOnly
            
        'validate if there is the same part number in the transaction
        'in every tag number there just only one part number
        If Not rsPartsCount.EOF And Not rsPartsCount.BOF Then
            MsgSpeechBox "Error: Part Number " & cboPartNo.Text & " was already used" & vbCrLf & _
                         "by tag number " & Null2String(rsPartsCount!tagno)
                On Error Resume Next
                 cboPartNo.SetFocus
                Exit Sub
        End If
        
        'Validation: (To prevent the duplication of tag number
        'There is a Dupliction when you add with the same tag number and part number
        
        If Not rsPHYCNTDUP2.EOF And Not rsPHYCNTDUP2.BOF Then
                MsgSpeechBox "Error: Tag Number " & txtTagNo.Text & " was already used" & vbCrLf & _
                         "by Part number " & Null2String(rsPHYCNTDUP2!PARTNO)
                On Error Resume Next
                txtTagNo.Text = "": txtTagNo.SetFocus
                Exit Sub
        
        End If
        
        If Not rsPHYCNTDUP.EOF And Not rsPHYCNTDUP.BOF Then
           
            MsgSpeechBox "Error: Tag Number " & txtTagNo.Text & " was already used" & vbCrLf & _
                         "by Part number " & Null2String(rsPHYCNTDUP!PARTNO)
                On Error Resume Next
                txtTagNo.SetFocus
                Exit Sub
        End If
        Else
            Set rsTAGS = New ADODB.Recordset
            'rsTAGS.Open "select tag from tags where tag = " & N2Str2Null(txtTagNo.Text), gconINVENTORY
            rsTAGS.Open "select tag from tags where tag = " & NumericVal(txtTagNo.Text), gconINVENTORY
            
            'validate if there is the same part number in the transaction
            'in every tag number there just only one part number
            If Not rsPartsCount.EOF And Not rsPartsCount.BOF Then
            MsgSpeechBox "Error: Part Number " & cboPartNo.Text & " was already used" & vbCrLf & _
                         "by tag number " & Null2String(rsPartsCount!tagno)
                On Error Resume Next
                 cboPartNo.SetFocus
                Exit Sub
             End If
             
            'Validation: (To prevent the duplication of tag number
            'There is a Dupliction when you add with the same tag number and part number
            If Not rsPHYCNTDUP2.EOF And Not rsPHYCNTDUP2.BOF Then
                MsgSpeechBox "Error: Tag Number " & txtTagNo.Text & " was already used" & vbCrLf & _
                         "by Part number " & Null2String(rsPHYCNTDUP2!PARTNO)
                On Error Resume Next
                txtTagNo.Text = "": txtTagNo.SetFocus
                Exit Sub
            End If
            If rsTAGS.EOF And rsTAGS.BOF Then
                MsgSpeechBox "This Tag number " & txtTagNo.Text & " is not being used or" & vbCrLf & _
                             "is not available in Tags Master File..."
                On Error Resume Next
                txtTagNo.SetFocus
                Exit Sub
            End If
       End If


    

    Dim vtxtTagNo, vtxtPARTNO, vtxtAdate               As String
    Dim VTXTLocation, vcboAmark, vtxtGroup_No, vtxtStatus As String
    Dim vtxtQCount, VTXTOnHand                         As Long
    Dim vtxtMAC                                        As Double
    Dim vtxtPARTDESC                                   As String
    Dim vVariance, vTotalMac                           As Double
    Dim vPrint_Stat, Vusercode                         As String
    Dim vDate, vTime                                   As String
    Dim vNewPARTNO As String

    Dim LastID                                         As Integer
    '************
    'updating code:     jaa - 11102008
    'vtxtTagNo = N2Str2Null(txtTagNo.Text)
    vtxtTagNo = NumericVal(txtTagNo.Text)
    '************
    vtxtPARTNO = N2Str2Null(cboPartNo.Text)
    vtxtQCount = NumericVal(txtQCount.Text)
    vtxtAdate = N2Date2Null(txtAdate.Text)
    VTXTLocation = N2Str2Null(txtLocation.Text)
    vcboAmark = N2Str2Null(cboAmark.Text)
    vtxtGroup_No = N2Str2Null(txtGroup_No.Text)
    vtxtStatus = N2Str2Null(txtStatus.Text)
    VTXTOnHand = NumericVal(txtOnHand.Text)
    vtxtMAC = NumericVal(txtMAC.Text)

    vVariance = vtxtQCount - VTXTOnHand
    vTotalMac = vVariance * vtxtMAC
    vPrint_Stat = "'N'"
    vDate = "'" & Date & "'"
    vTime = "'" & Time & "'"
    Vusercode = "'" & Left(LOGCODE, 2) & "'"
    vtxtPARTDESC = N2Str2Null(labPartDesc.Caption)

    vNewPARTNO = N2Str2Null(cboPartNo.Text)

    If ADDOREDIT = "ADD" Then

        Set rsPHYCNTDUP = New ADODB.Recordset
        rsPHYCNTDUP.Open "Select id from PHYCNT  order by id asc", gconINVENTORY, adOpenKeyset
        If Not rsPHYCNTDUP.EOF And Not rsPHYCNTDUP.BOF Then
            rsPHYCNTDUP.MoveLast
            LastID = N2Str2Zero(rsPHYCNTDUP!ID) + 1
        End If
        SQL_STATEMENT = "insert into phycnt " & _
                        "(id,tagno,PARTNO,PARTDESC,qcount,adate,location,amark,group_no,status,onhand,mac" & _
                        ",variance,totalmac,print_stat,lastupdate,[time],usercode,newPARTNO)" & _
                      " values (" & LastID & ", " & vtxtTagNo & ", " & vtxtPARTNO & ", " & vtxtPARTDESC & ", " & vtxtQCount & ", " & vtxtAdate & ", " & VTXTLocation & ", " & vcboAmark & ", " & vtxtGroup_No & ", " & vtxtStatus & ", " & VTXTOnHand & ", " & vtxtMAC & _
                        ", " & vVariance & ", " & vTotalMac & ", " & vPrint_Stat & ", " & vDate & ", " & vTime & ", " & Vusercode & ", " & vNewPARTNO & ")"
        gconINVENTORY.Execute SQL_STATEMENT
        NEW_LogAudit "A", "PHYSICAL COUNT", SQL_STATEMENT, N2Str2Null(LastID), "", N2Str2Null(vtxtTagNo), "", ""

        gconINVENTORY.Execute "update CUTOFF set" & _
                            " TAGNO = " & vtxtTagNo & _
                            " where PARTNO = " & vtxtPARTNO
        gconINVENTORY.Execute "update tags set" & _
                            " PARTNO = " & vtxtPARTNO & _
                            " where tag = " & vtxtTagNo

    Else
        SQL_STATEMENT = "update phycnt set" & _
                      " PARTNO = " & vtxtPARTNO & "," & _
                      " PARTDESC = " & vtxtPARTDESC & "," & _
                      " qcount = " & vtxtQCount & "," & _
                      " adate = " & vtxtAdate & "," & _
                      " location = " & VTXTLocation & "," & _
                      " amark = " & vcboAmark & "," & _
                      " group_no = " & vtxtGroup_No & "," & _
                      " status = " & vtxtStatus & "," & _
                      " onhand = " & VTXTOnHand & "," & _
                      " mac = " & vtxtMAC & "," & _
                      " variance = " & vVariance & "," & _
                      " totalmac = " & vTotalMac & "," & _
                      " print_stat = " & vPrint_Stat & "," & _
                      " lastupdate = " & vDate & "," & _
                      " [time] = " & vTime & "," & _
                      " usercode = " & Vusercode & "," & _
                      " newPARTNO = " & vNewPARTNO & _
                      " where id= " & labDetID
                      
        gconINVENTORY.Execute SQL_STATEMENT
        NEW_LogAudit "E", "PHYSICAL COUNT", SQL_STATEMENT, N2Str2Null(LastID), "", N2Str2Null(vtxtTagNo), "", ""

        gconINVENTORY.Execute "update tags set" & _
                            " PARTNO = " & vtxtPARTNO & _
                            " where tag = " & vtxtTagNo
                            
        gconINVENTORY.Execute "update CUTOFF set" & _
                            " TAGNO = " & vtxtTagNo & _
                            " where PARTNO = " & vtxtPARTNO
    End If
    cleargrid grdPhyCnt
    rsRefresh
    InitGrid
    FillGrid
    initMemvars
    
    If ADDOREDIT = "EDIT" Then
        cmdCancel.Value = True
    Else
    On Error Resume Next
        txtTagNo.SetFocus
    End If
    
    Exit Sub

ERRORCODE:
    MsgBox "Tag Number " & txtTagNo.Text & " is already used by another user ", vbInformation, "Please Change your  Tag Number"
    Exit Sub
    
End Sub

Private Sub Command1_Click()
InitGrid:     FillGrid
End Sub

Private Sub Command2_Click()
cmdCancel_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ADDOREDIT = ""
            initMemvars
            txtSearchPartNo.SetFocus
            picPhyCnt2.ZOrder 1
        Case vbKeyF2
            ADDOREDIT = "ADD": picPhyCnt2.ZOrder 0: initMemvars
            txtTagNo.Enabled = True: On Error Resume Next: txtTagNo.SetFocus
        Case vbKeyF3
            grdPhyCnt.Col = 0
            If grdPhyCnt.Text = "No Entry" Or grdPhyCnt.Text = "TAG NO." Then
                MsgSpeechBox "Nothing to Edit!": Exit Sub
            End If
            ADDOREDIT = "EDIT": txtTagNo.Enabled = False: picPhyCnt2.ZOrder 0: initMemvars: StoreMemvars
        Case vbKeyF4: cmdDelete_Click
        Case vbKeyF5: Unload Me
        Case vbKeyF8
            txtSearchPartNo.Enabled = True
            txtSearchPartNo.SetFocus
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1: cleargrid grdPhyCnt: rsRefresh: initMemvars: InitGrid: FillGrid
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtSearchPartNo.Text = "":
    picPhyCnt2.ZOrder 1: Screen.MousePointer = 0
    Combo_Loadval cboPartNo, gconINVENTORY.Execute("Select PARTNO from CUTOFF order by partno")

End Sub

Private Sub grdPhyCnt_DblClick()
    grdPhyCnt.Col = 0
    Picture1.Enabled = False
    
    If grdPhyCnt.Text = "No Entry" Or grdPhyCnt.Text = "TAG NO." Then
        MsgSpeechBox "Nothing to Edit!"
        Exit Sub
    End If
    picPhyCnt2.ZOrder 0
    picPhyCnt2.Visible = True
    initMemvars
    StoreMemvars
    ADDOREDIT = "EDIT"
    txtTagNo.Enabled = False
End Sub

Private Sub Option1_Click()
    On Error Resume Next
    txtSearchPartNo.SetFocus
End Sub

Private Sub Option2_Click()
    On Error Resume Next
    txtSearchPartNo.SetFocus
End Sub

Private Sub Option3_Click()
    On Error Resume Next
    txtSearchPartNo.SetFocus
End Sub

Private Sub Timer1_Timer()
    Exit Sub
    If NumericVal(txtQCount.Text) <> 0 Then
        If NumericVal(txtQCount.Text) = NumericVal(txtOnHand.Text) Then
            labPhyCntStatus.Caption = "Proof in Balance"
        ElseIf NumericVal(txtQCount.Text) > NumericVal(txtOnHand.Text) Then
            labPhyCntStatus.Caption = "Positive Variance"
        Else
            labPhyCntStatus.Caption = "Negative Variance"
        End If
    Else
        labPhyCntStatus.Caption = ""
    End If
    If labPhyCntStatus.Visible = False Then
        labPhyCntStatus.Visible = True
    Else
        labPhyCntStatus.Visible = False
    End If
End Sub

Private Sub txtAdate_GotFocus()
    txtAdate.Text = Format(txtAdate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtAdate_LostFocus()
    txtAdate.Text = Format(txtAdate.Text, "MM/DD/YYYY")
End Sub

Private Sub txtLocation_Change()
    If Trim(txtLocation.Text) = "" Then
        labeeror = "Warning: Location is empty, please enter the correct location before saving this ticket."
    Else
        labeeror = ""
    End If
End Sub

Private Sub txtOnHand_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub cboPartNo_Change()
    cboPartNo_Click
End Sub

Private Sub txtQCount_Change()
    If NumericVal(txtQCount.Text) <> 0 Then
    Else
        labPhyCntStatus.Caption = ""
    End If
End Sub

Private Sub txtQCount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtQCount_LostFocus()
    If NumericVal(txtQCount.Text) < 0 Then MsgBoxXP "Quantity Counted must not be less than zero!", "Invalid QTY Counted", XP_OKOnly, msg_Exclamation
End Sub

Private Sub txtSearchPartNo_Change()
    txtSearchPARTNO_KeyPress 13
End Sub

Private Sub txtSearchPARTNO_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
    If KeyAscii = 13 Then

        Set rsPHYCNT = New ADODB.Recordset
        If Option2.Value = True Then
            rsPHYCNT.Open "Select TOP 50 * from PHYCNT where PARTNO like '" & Repleys(txtSearchPartNo) & "%' order by PARTNO asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
        ElseIf Option1.Value = True Then
            rsPHYCNT.Open "Select TOP 50 * from PHYCNT where TagNo like '" & Repleys(txtSearchPartNo) & "%' order by tagno asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
        Else
            rsPHYCNT.Open "Select TOP 50 * from PHYCNT where Location like '" & Repleys(txtSearchPartNo) & "%' order by Location asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
        End If
        Dim kcnt                                                      As Integer
        kcnt = 0
        cleargrid grdPhyCnt
        
        
        
        'If Not rsPHYCNT.EOF And Not rsPHYCNT.BOF Then

            'Screen.MousePointer = 11
            'rsPHYCNT.MoveFirst
            'Do While Not rsPHYCNT.EOF
            '    kcnt = kcnt + 1
             '   grdPhyCnt.AddItem Format(Null2String(rsPHYCNT!tagno), "0000000000") & Chr(9) & _
                                  'Null2String(rsPHYCNT!PARTNO) & Chr(9) & _
                                  'Null2String(rsPHYCNT!PARTDESC) & Chr(9) & _
                                  'Null2String(rsPHYCNT!Location) & Chr(9) & _
                                  'Null2String(rsPHYCNT!ONHAND) & Chr(9) & _
                                  'Null2String(rsPHYCNT!Qcount) & Chr(9) & _
                                  'Null2String(rsPHYCNT!VARIANCE) & Chr(9) & _
                                  'Null2String(rsPHYCNT!Mac) & Chr(9) & _
                                  'Null2String(rsPHYCNT!totalmac) & Chr(9) & _
                                  'Null2String(rsPHYCNT!ADate) & Chr(9) & _
                                  'Null2String(rsPHYCNT!Amark) & Chr(9) & _
                                  'Null2String(rsPHYCNT!Print_Stat) & Chr(9) & _
                                  'Null2String(rsPHYCNT!STATUS) & Chr(9) & _
                                 'Null2String(rsPHYCNT!LASTUPDATE) & Chr(9) & _
                                 'Null2String(rsPHYCNT![Time]) & Chr(9) & _
                                 'Null2String(rsPHYCNT!USERCODE)
      
                         
                'rsPHYCNT.MoveNext
            'Loop
            
            'updating code:     jbf - 01262009
          
            
            If Not rsPHYCNT.EOF And Not rsPHYCNT.BOF Then

            Screen.MousePointer = 11
            rsPHYCNT.MoveFirst
            Do While Not rsPHYCNT.EOF
                kcnt = kcnt + 1
                grdPhyCnt.AddItem Format(Null2String(rsPHYCNT!tagno), "0000000000") & Chr(9) & _
                                  Null2String(rsPHYCNT!PARTNO) & Chr(9) & _
                                  Null2String(rsPHYCNT!PARTDESC) & Chr(9) & _
                                  Null2String(rsPHYCNT!Location) & Chr(9) & _
                                  Null2String(rsPHYCNT!ONHAND) & Chr(9) & _
                                  Null2String(rsPHYCNT!Qcount) & Chr(9) & _
                                  Null2String(rsPHYCNT!VARIANCE) & Chr(9) & _
                                  Null2String(rsPHYCNT!Mac) & Chr(9) & _
                                  Null2String(rsPHYCNT!totalmac) & Chr(9) & _
                                  Null2String(rsPHYCNT!ADate) & Chr(9) & _
                                  Null2String(rsPHYCNT!Amark) & Chr(9) & _
                                  Null2String(rsPHYCNT!Print_Stat) & Chr(9) & _
                                  Null2String(rsPHYCNT!STATUS) & Chr(9) & _
                                  Null2String(rsPHYCNT!LASTUPDATE) & Chr(9) & _
                                  Null2String(rsPHYCNT![Time]) & Chr(9) & _
                                  Null2String(rsPHYCNT!USERCODE) & Chr(9) & Null2String(rsPHYCNT!ID)
           
                         
                rsPHYCNT.MoveNext
            Loop
            
            If kcnt <> 0 Then grdPhyCnt.RemoveItem 1
            Screen.MousePointer = 0
        End If
    End If
End Sub

Private Sub txtTagNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub




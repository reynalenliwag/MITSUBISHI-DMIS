VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmAMISMASTERFILEVendor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Master List"
   ClientHeight    =   7815
   ClientLeft      =   720
   ClientTop       =   435
   ClientWidth     =   6975
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Vendor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   6975
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   630
      ScaleHeight     =   885
      ScaleWidth      =   5805
      TabIndex        =   37
      Top             =   6885
      Width           =   5805
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
         Left            =   5040
         MouseIcon       =   "Vendor.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Exit Window"
         Top             =   60
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
         Left            =   4320
         MouseIcon       =   "Vendor.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Print this Record"
         Top             =   60
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
         Left            =   3600
         MouseIcon       =   "Vendor.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   735
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
         Left            =   2880
         MouseIcon       =   "Vendor.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
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
         Left            =   2160
         MouseIcon       =   "Vendor.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
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
         Left            =   1440
         MouseIcon       =   "Vendor.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   735
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
         MouseIcon       =   "Vendor.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   735
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
         Left            =   0
         MouseIcon       =   "Vendor.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport rptVendor 
      Left            =   6420
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   45
      TabIndex        =   24
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox cboTerms 
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
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2310
         Width           =   945
      End
      Begin VB.TextBox txtRATE 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox cboATC 
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
         Left            =   960
         TabIndex        =   10
         Top             =   2280
         Width           =   1395
      End
      Begin VB.TextBox txtCity 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   1785
      End
      Begin VB.CheckBox Check1 
         Caption         =   "NON-VAT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4710
         TabIndex        =   7
         Top             =   1590
         Width           =   1275
      End
      Begin VB.TextBox txtAddress 
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
         ForeColor       =   &H00701E2A&
         Height          =   510
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "Vendor.frx":2D71
         Top             =   960
         Width           =   3585
      End
      Begin VB.TextBox txtTIN 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   4710
         MaxLength       =   18
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1890
         Width           =   2085
      End
      Begin VB.TextBox txtFax 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   960
         MaxLength       =   17
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1890
         Width           =   2565
      End
      Begin VB.TextBox txtPhone 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   960
         MaxLength       =   17
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1500
         Width           =   2565
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   570
         Width           =   1785
      End
      Begin VB.TextBox txtContact 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   960
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   570
         Width           =   3165
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   960
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   1155
      End
      Begin VB.TextBox txtNameofVendor 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2790
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   4005
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   2
         Left            =   4080
         TabIndex        =   46
         Top             =   1920
         Width           =   150
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   -30
         TabIndex        =   45
         Top             =   900
         Width           =   150
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   2070
         TabIndex        =   44
         Top             =   120
         Width           =   150
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Height          =   210
         Left            =   3600
         TabIndex        =   42
         Top             =   2310
         Width           =   225
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Left            =   2400
         TabIndex        =   41
         Top             =   2340
         Width           =   435
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ATC"
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
         Left            =   495
         TabIndex        =   40
         Top             =   2310
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
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
         Left            =   4575
         TabIndex        =   39
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   2235
         TabIndex        =   26
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TIN"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   1950
         Width           =   645
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1950
         Width           =   645
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
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
         Height          =   255
         Left            =   3990
         TabIndex        =   33
         Top             =   2310
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
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
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Height          =   255
         Left            =   -570
         TabIndex        =   31
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         Height          =   255
         Left            =   3480
         TabIndex        =   30
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   255
         Left            =   -570
         TabIndex        =   25
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4710
         TabIndex        =   28
         Top             =   240
         Width           =   225
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         TabIndex        =   27
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
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
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   4215
      Left            =   45
      TabIndex        =   36
      Top             =   2685
      Width           =   6855
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
         Left            =   60
         MaxLength       =   19
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   150
         Width           =   6735
      End
      Begin MSComctlLib.ListView lstVendor 
         Height          =   3615
         Left            =   30
         TabIndex        =   13
         Top             =   540
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   6376
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
         MouseIcon       =   "Vendor.frx":2D77
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "VENDOR NAME"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4935
      ScaleHeight     =   885
      ScaleWidth      =   1485
      TabIndex        =   38
      Top             =   6915
      Width           =   1485
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
         MouseIcon       =   "Vendor.frx":2ED9
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":302B
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
         MouseIcon       =   "Vendor.frx":3369
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":34BB
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAMISMASTERFILEVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsVENDOR                                                As ADODB.Recordset
Dim rsATC                                                   As ADODB.Recordset
Dim rsTERMS                                                 As ADODB.Recordset
Dim AddorEdit                                               As String

Private Sub cboATC_Change()
    Set rsATC = New ADODB.Recordset
    rsATC.Open "SELECT * FROM AMIS_ATC WHERE ATC = '" & cboATC.Text & "'", gconDMIS, adOpenForwardOnly
    If Not rsATC.EOF And Not rsATC.BOF Then
        txtRATE.Text = rsATC!Rate
    Else
        txtRATE.Text = ""
    End If
    Set rsATC = Nothing
End Sub

Private Sub cboATC_Click()
    Set rsATC = New ADODB.Recordset
    rsATC.Open "SELECT * FROM AMIS_ATC WHERE ATC = '" & cboATC.Text & "'", gconDMIS, adOpenForwardOnly
    If Not rsATC.EOF And Not rsATC.BOF Then
        txtRATE.Text = rsATC!Rate
    Else
        MsgBox "Please ATC from the list..", vbInformation, "ATC"
        cboATC.SetFocus
    End If
    Set rsATC = Nothing
End Sub

Private Sub cboATC_LostFocus()
    Set rsATC = New ADODB.Recordset
    If cboATC.Text <> "" Then
        rsATC.Open "SELECT * FROM AMIS_ATC WHERE ATC = '" & cboATC.Text & "'", gconDMIS, adOpenForwardOnly
        If Not rsATC.EOF And Not rsATC.BOF Then
            txtRATE.Text = rsATC!Rate
        Else
            MsgBox "Please ATC from the list..", vbInformation, "ATC"
            cboATC.SetFocus
        End If
        Set rsATC = Nothing
    End If
End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "VENDORS") = False Then Exit Sub

    AddorEdit = "ADD"
    initMemvars
    txtNameofVendor.SetFocus
    Picture1.Visible = False
    Picture2.Visible = True
    lstVendor.Enabled = False
    txtSearch.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstVendor.Enabled = True
    fraDetails.Enabled = True
    txtSearch.Enabled = True
    StoreMemVars
End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "VENDORS") = False Then Exit Sub
    On Error GoTo ErrorCode:
    Dim lngCount                                            As Long


    lngCount = gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_vw_PO_HISTORY WHERE SUPCODE=" & N2Str2Null(txtCode)).Fields(0).Value
    If lngCount > 0 Then
        MsgBox "Vendor Record Exists in Purchase Order." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If


    lngCount = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_JOURNAL_HD  WHERE VendorCode=" & N2Str2Null(txtCode)).Fields(0).Value
    If lngCount > 0 Then
        MsgBox "Vendor Record Exists in Journal Books." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If
    lngCount = gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_vw_RR_Trans WHERE RECVD_CODE=" & N2Str2Null(txtCode)).Fields(0).Value
    If lngCount > 0 Then
        MsgBox "Vendor Record Exists in Parts Transaction." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If


    lngCount = gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_Po_hd WHERE CONTRACTOR_CODE=" & N2Str2Null(txtCode)).Fields(0).Value
    If lngCount > 0 Then
        MsgBox "Vendor Record Exists in Sublet Purchase Order." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If



    lngCount = gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_PO_RC_HD WHERE CONTRACTOR_CODE=" & N2Str2Null(txtCode)).Fields(0).Value
    If lngCount > 0 Then
        MsgBox "Vendor Record Exists in Sublet Receiving." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If




    If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
        gconDMIS.Execute "delete from ALL_Vendor_TABLE where Code = '" & lstVendor.SelectedItem & "'"
        LogAudit "X", "VENDOR MASTER FILE", lstVendor.SelectedItem.SubItems(1) & txtNameofVendor
        NEW_LogAudit "X", "VENDOR MASTER FILE", SQL_STATEMENT, "", "", txtCode, "", ""
    End If
    rsRefresh
    StoreMemVars
    FillGrid

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "VENDORS") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    StoreEntry labID.Caption
    '(lstVendor.SelectedItem.SubItems(2))
    lstVendor.Enabled = False
    lstVendor.Enabled = False
    txtSearch.Enabled = False
    txtSearch.Enabled = False
    lstVendor.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdFind_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsVENDOR.MoveNext
    If rsVENDOR.EOF Then
        rsVENDOR.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:59
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsVENDOR.MovePrevious
    If rsVENDOR.BOF Then
        rsVENDOR.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "VENDORS") = False Then Exit Sub
    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    Screen.MousePointer = 11
    rptVendor.Reset

    rptVendor.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptVendor.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptVendor, AMIS_REPORT_PATH & "\Files\Suppliers.rpt", "", DMIS_REPORT_Connection, 1

    Screen.MousePointer = 0
    LogAudit "V", "VENDOR MASTER FILE", txtCode
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0

End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrorCode

    Dim VtxtCode, VtxtNameofVendor, VtxtContact             As String
    Dim VtxtPosition, VtxtAddress, VtxtPhone, vtxtCity      As String
    Dim Vcboterms, VtxtFax, VtxtTIN                         As String
    Dim VchkVAT                                             As String
    Dim vcboATC                                             As String
    Dim rsVendorDup                                         As ADODB.Recordset
    VtxtCode = N2Str2Null(RTrim(LTrim(txtCode.Text)))
    VtxtNameofVendor = N2Str2Null(txtNameofVendor.Text)
    VtxtContact = N2Str2Null(txtContact.Text)
    VtxtPosition = N2Str2Null(txtPosition.Text)
    VtxtAddress = N2Str2Null(txtAddress.Text)
    vtxtCity = N2Str2Null(txtCity.Text)
    VtxtPhone = N2Str2Null(txtPhone.Text)
    Vcboterms = N2Str2Null(cboTerms.Text)
    VtxtFax = N2Str2Null(txtFax.Text)
    VtxtTIN = N2Str2Null(txtTin.Text)
    vcboATC = N2Str2Null(cboATC.Text)
    If Check1.Value = 1 Then
        VchkVAT = "'Y'"
    Else
        VchkVAT = "'N'"
    End If

    If AddorEdit = "ADD" Then

        Set rsVendorDup = New ADODB.Recordset
        rsVendorDup.Open "select code from ALL_Vendor_TABLE where code = " & VtxtCode, gconDMIS
        If Not rsVendorDup.EOF And Not rsVendorDup.BOF Then
            MsgBox "Vendor Code Already Exist!", vbCritical, "Duplicate Code Not Allowed"
            Exit Sub
        End If
        If txtNameofVendor.Text = "" Then
            ShowIsRequiredMsg "Name of Vendor"
            On Error Resume Next
            txtNameofVendor.SetFocus
            Exit Sub
        ElseIf txtAddress.Text = "" Then
            ShowIsRequiredMsg "Address"
            On Error Resume Next
            txtAddress.SetFocus
            Exit Sub
        ElseIf txtCity.Text = "" Then
            ShowIsRequiredMsg "City"
            On Error Resume Next
            txtCity.SetFocus
            Exit Sub
        ElseIf txtTin.Text = "" Then
            ShowIsRequiredMsg "TIN"
            On Error Resume Next
            txtTin.SetFocus
            Exit Sub
        End If
        SQL_STATEMENT = "Insert into ALL_Vendor_TABLE " & _
                        "(Code,NameofVendor,contact,[position],address,Address2,phone,[Terms],fax,tin,NONVAT,ATC) " & _
                        " values (" & VtxtCode & _
                        ", " & VtxtNameofVendor & ", " & VtxtContact & _
                        ", " & VtxtPosition & ", " & VtxtAddress & "," & vtxtCity & _
                        ", " & VtxtPhone & ", " & Vcboterms & _
                        ", " & VtxtFax & ", " & VtxtTIN & "," & VchkVAT & "," & vcboATC & ")"
        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(txtCode), "Code", "ALL_Vendor", ""))
        NEW_LogAudit "A", "VENDOR MASTER FILE", SQL_STATEMENT, TransactionID, "", txtCode, "", ""
    Else

        If txtCode <> Null2String(rsVENDOR!CODE) Then
            Set rsVendorDup = New ADODB.Recordset
            rsVendorDup.Open "select code from ALL_Vendor_TABLE where code = " & VtxtCode, gconDMIS
            If Not rsVendorDup.EOF And Not rsVendorDup.BOF Then
                MsgBox "Vendor Code Already Exist!", vbCritical, "Duplicate Code Not Allowed"
                Exit Sub
            End If
        End If
        If txtAddress.Text = "" Then
            ShowIsRequiredMsg "Address"
            On Error Resume Next
            txtAddress.SetFocus
            Exit Sub
        ElseIf txtTin.Text = "" Then
            ShowIsRequiredMsg "TIN"
            On Error Resume Next
            txtTin.SetFocus
            Exit Sub
        End If

        If COMPANY_CODE = "HMH" Then
            If txtCity.Text = "" Then
                ShowIsRequiredMsg "City"
                On Error Resume Next
                txtCity.SetFocus
                Exit Sub
            End If
        End If
        SQL_STATEMENT = "update ALL_Vendor_TABLE set" & _
                        " NameofVendor = " & VtxtNameofVendor & "," & _
                        " contact = " & VtxtContact & "," & _
                        " [position] = " & VtxtPosition & "," & _
                        " address = " & VtxtAddress & "," & _
                        " Address2 = " & vtxtCity & "," & _
                        " phone = " & VtxtPhone & "," & _
                        " [Terms] = " & Vcboterms & "," & _
                        " fax = " & VtxtFax & "," & _
                        " NONVAT = " & VchkVAT & "," & _
                        " ATC = " & vcboATC & "," & _
                        " tin = " & VtxtTIN & _
                        " where ID = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(txtCode), "Code", "ALL_Vendor", ""))
        NEW_LogAudit "E", "VENDOR MASTER FILE", SQL_STATEMENT, TransactionID, "", txtCode, "", ""
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsVENDOR.Find "code = " & VtxtCode
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "VENDOR MASTER FILE"
        Call frmALL_AuditInquiry.DisplayHistory(labID, "VENDOR MASTER FILE")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    InitCbo
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "select * from ALL_Vendor_TABLE where Code <> '999999' order by NameofVendor asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtCode.Text = ""
    txtNameofVendor.Text = ""
    txtContact.Text = ""
    txtPosition.Text = ""
    txtAddress.Text = ""
    txtCity.Text = ""
    txtPhone.Text = ""
    txtFax.Text = ""
    txtTin.Text = ""
    txtSearch.Text = ""
    cboATC.Text = ""
    cboTerms.ListIndex = 0
End Sub

Sub StoreMemVars()
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        Frame1.Enabled = False
        labID.Caption = rsVENDOR!ID
        txtCode.Text = Null2String(rsVENDOR!CODE)
        txtNameofVendor.Text = Null2String(rsVENDOR!nameofvendor)
        txtContact.Text = Null2String(rsVENDOR!CONTACT)
        txtPosition.Text = Null2String(rsVENDOR!Position)
        txtAddress.Text = Null2String(rsVENDOR!Address)
        txtCity.Text = Null2String(rsVENDOR!ADDRESS2)
        txtPhone.Text = Null2String(rsVENDOR!Phone)
        cboTerms.Text = NumericVal(rsVENDOR!TERMS)
        txtFax.Text = Null2String(rsVENDOR!Fax)
        txtTin.Text = Null2String(rsVENDOR!TIN)
        cboATC.Text = Null2String(rsVENDOR!ATC)
        If Null2String(rsVENDOR!NONVAT) = "Y" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    Else
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsVendor2                                           As ADODB.Recordset
    Set rsVendor2 = New ADODB.Recordset
    rsVendor2.Open "select * from ALL_Vendor_TABLE where ID = " & XXX, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVendor2.EOF And Not rsVendor2.BOF Then
        fraDetails.Enabled = False
        labID.Caption = rsVendor2!ID
        txtCode.Text = Null2String(rsVendor2!CODE)
        txtNameofVendor.Text = Null2String(rsVendor2!nameofvendor)
        txtContact.Text = Null2String(rsVendor2!CONTACT)
        txtPosition.Text = Null2String(rsVendor2!Position)
        txtAddress.Text = Null2String(rsVendor2!Address)
        txtCity.Text = Null2String(rsVendor2!ADDRESS2)
        txtPhone.Text = Null2String(rsVendor2!Phone)
        cboTerms.Text = Null2String(rsVendor2!TERMS)
        txtFax.Text = Null2String(rsVendor2!Fax)
        txtTin.Text = Null2String(rsVendor2!TIN)
        cboATC.Text = Null2String(rsVendor2!ATC)
        If Null2String(rsVENDOR!NONVAT) = "Y" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    End If
End Sub

Private Sub FillGrid()
    Dim rsVendor2                                           As ADODB.Recordset
    lstVendor.Enabled = False
    lstVendor.Sorted = False: lstVendor.ListItems.Clear
    Set rsVendor2 = New ADODB.Recordset
    Set rsVendor2 = gconDMIS.Execute("select code,nameofvendor,ID from ALL_Vendor_TABLE where code <> '999999' order by NameofVendor asc")
    If Not (rsVendor2.EOF And rsVendor2.BOF) Then
        Listview_Loadval Me.lstVendor.ListItems, rsVendor2
        lstVendor.Refresh
        lstVendor.Enabled = True
    End If

End Sub

Private Sub FillSearchGrid(XXX As String)
    Dim rsVendor2                                           As ADODB.Recordset

    lstVendor.Enabled = False: lstVendor.Sorted = False: lstVendor.ListItems.Clear
    Set rsVendor2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsVendor2 = gconDMIS.Execute("select code,nameofvendor,ID from ALL_Vendor_TABLE where nameofvendor like '" & Repleys(XXX) & "%'")
    If Not (rsVendor2.EOF And rsVendor2.BOF) Then
        Listview_Loadval Me.lstVendor.ListItems, rsVendor2
        lstVendor.Refresh
        lstVendor.Enabled = True
    End If
End Sub

Private Sub lstVendor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstVendor
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

Private Sub lstVendor_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstVendor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    rsVENDOR.Bookmark = rsFind(rsVENDOR.Clone, "code", lstVendor.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNameofVendor_Change()
    If AddorEdit = "ADD" Then
        If Len(txtNameofVendor.Text) <> 0 Then
            Dim rsVendorCodes                               As ADODB.Recordset
            Set rsVendorCodes = New ADODB.Recordset
            rsVendorCodes.Open "select code from ALL_Vendor_TABLE where LEFT(code,1) = '" & Left(txtNameofVendor.Text, 1) & "' ORDER BY CODE", gconDMIS
            If Not rsVendorCodes.EOF And Not rsVendorCodes.BOF Then
                rsVendorCodes.MoveLast
                txtCode.Text = Left(Trim(txtNameofVendor.Text), 1) & Format(N2Str2Zero(Mid(rsVendorCodes!CODE, 2, Len(rsVendorCodes!CODE) - 1)) + 1, "00000")
            Else
                txtCode.Text = Left(Trim(txtNameofVendor.Text), 1) & "00001"
            End If
        End If
    End If
End Sub

Private Sub txtNameofVendor_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNameofVendor_LostFocus()
    If AddorEdit = "ADD" Then
        If Len(txtNameofVendor.Text) <> 0 Then
            Dim rsVendorCodes                               As ADODB.Recordset
            Set rsVendorCodes = New ADODB.Recordset
            rsVendorCodes.Open "select code from ALL_Vendor_TABLE where LEFT(code,1) = '" & Left(txtNameofVendor.Text, 1) & "' ORDER BY CODE", gconDMIS
            If Not rsVendorCodes.EOF And Not rsVendorCodes.BOF Then
                rsVendorCodes.MoveLast
                txtCode.Text = Left(Trim(txtNameofVendor.Text), 1) & Format(N2Str2Zero(Mid(rsVendorCodes!CODE, 2, Len(rsVendorCodes!CODE) - 1)) + 1, "00000")
            Else
                txtCode.Text = Left(Trim(txtNameofVendor.Text), 1) & "00001"
            End If
        End If
    End If
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstVendor.ListItems.Count > 0 And lstVendor.Enabled = True Then
            lstVendor.SetFocus
        End If
    End If
End Sub

Sub InitCbo()
    Set rsATC = New ADODB.Recordset
    rsATC.Open "SELECT * FROM AMIS_ATC", gconDMIS, adOpenForwardOnly
    cboATC.Clear
    If Not rsATC.EOF And Not rsATC.BOF Then
        Do While Not rsATC.EOF
            cboATC.AddItem RTrim(LTrim(Null2String(rsATC!ATC)))
            rsATC.MoveNext
        Loop
    End If
    Set rsATC = Nothing
    
    Set rsTERMS = New ADODB.Recordset
    rsTERMS.Open "SELECT * FROM ALL_PAYTERM ORDER BY NO_DAYS ASC", gconDMIS, adOpenForwardOnly
    cboTerms.Clear
    If Not rsTERMS.EOF And Not rsTERMS.BOF Then
        Do While Not rsTERMS.EOF
            cboTerms.AddItem RTrim(LTrim(Null2String(rsTERMS!no_Days)))
            rsTERMS.MoveNext
        Loop
    End If
    Set rsTERMS = Nothing
End Sub

Private Sub cboterms_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    Else
    End If
End Sub

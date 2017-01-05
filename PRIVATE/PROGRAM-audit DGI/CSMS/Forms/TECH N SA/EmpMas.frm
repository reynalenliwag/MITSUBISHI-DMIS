VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_SA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVICE ADVISER DATA ENTRY"
   ClientHeight    =   5550
   ClientLeft      =   720
   ClientTop       =   330
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "EmpMas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   12495
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   2670
      ScaleHeight     =   4545
      ScaleWidth      =   9765
      TabIndex        =   29
      Top             =   60
      Width           =   9795
      Begin VB.TextBox txtPOSITION 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2760
         Width           =   4695
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   1080
         ScaleHeight     =   885
         ScaleWidth      =   6795
         TabIndex        =   54
         Top             =   3180
         Width           =   6825
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   6300
            Top             =   450
         End
         Begin VB.CheckBox Check1 
            Caption         =   "SA Master"
            Height          =   255
            Left            =   60
            TabIndex        =   62
            Top             =   60
            Width           =   1425
         End
         Begin VB.CheckBox Check2 
            Caption         =   "SA Certified"
            Height          =   255
            Left            =   60
            TabIndex        =   61
            Top             =   330
            Width           =   1335
         End
         Begin VB.CheckBox Check3 
            Caption         =   "SA New"
            Height          =   255
            Left            =   1530
            TabIndex        =   60
            Top             =   60
            Width           =   1005
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Foreman"
            Height          =   255
            Left            =   1530
            TabIndex        =   59
            Top             =   330
            Width           =   1245
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Warranty"
            Height          =   255
            Left            =   2790
            TabIndex        =   58
            Top             =   60
            Width           =   2025
         End
         Begin VB.CheckBox Check6 
            Caption         =   "In-House Instructor"
            Height          =   255
            Left            =   2790
            TabIndex        =   57
            Top             =   330
            Width           =   2025
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Billing Staff"
            Height          =   255
            Left            =   4830
            TabIndex        =   56
            Top             =   60
            Width           =   2025
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Others"
            Height          =   255
            Left            =   4830
            TabIndex        =   55
            Top             =   330
            Width           =   2025
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TAGGING OF POSITION IS UNDER REVISION"
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
            Height          =   225
            Left            =   30
            TabIndex        =   64
            Top             =   630
            Visible         =   0   'False
            Width           =   3585
         End
      End
      Begin VB.TextBox txtTELE 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1170
         Width           =   1845
      End
      Begin VB.TextBox txtDRES 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2340
         Width           =   1845
      End
      Begin VB.TextBox txtDHIRED 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1950
         Width           =   1845
      End
      Begin VB.TextBox txtCITI 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2340
         Width           =   4695
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1170
         Width           =   5625
      End
      Begin VB.TextBox txtReligion 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1950
         Width           =   4695
      End
      Begin VB.ComboBox cboDept 
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
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Combo1"
         Top             =   4140
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "XXX UPLOADING XXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7260
         TabIndex        =   40
         Top             =   4170
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   2
         Top             =   750
         Width           =   2475
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4620
         MaxLength       =   50
         TabIndex        =   3
         Top             =   750
         Width           =   2475
      End
      Begin VB.TextBox txtMiddleInt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   4
         Top             =   750
         Width           =   1395
      End
      Begin VB.TextBox txtEmpNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7830
         MaxLength       =   10
         TabIndex        =   1
         Top             =   330
         Width           =   1845
      End
      Begin VB.TextBox txtENO 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Top             =   360
         Width           =   1845
      End
      Begin VB.TextBox txtBDATE 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1560
         Width           =   1845
      End
      Begin VB.Label lblEMPTYPE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
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
         Left            =   8520
         TabIndex        =   63
         Top             =   3180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Resigned"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   6480
         TabIndex        =   53
         Top             =   2460
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   6810
         TabIndex        =   52
         Top             =   2070
         Width           =   885
      End
      Begin VB.Label labid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   8490
         TabIndex        =   51
         Top             =   30
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   345
         TabIndex        =   46
         Top             =   1260
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   6825
         TabIndex        =   45
         Top             =   1290
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   300
         TabIndex        =   44
         Top             =   1650
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   345
         TabIndex        =   42
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   3630
         TabIndex        =   38
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   7170
         TabIndex        =   37
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   7155
         TabIndex        =   36
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. No"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   35
         Top             =   480
         Width           =   720
      End
      Begin MSForms.CheckBox chkActive 
         Height          =   315
         Left            =   5700
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   915
         BackColor       =   -2147483633
         ForeColor       =   0
         DisplayStyle    =   4
         Size            =   "1614;556"
         Value           =   "0"
         Caption         =   "Active"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   2820
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   32
         Top             =   4230
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblDepCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
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
         Left            =   8520
         TabIndex        =   31
         Top             =   2790
         Visible         =   0   'False
         Width           =   1125
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   10425
         _Version        =   655364
         _ExtentX        =   18389
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "INFORMATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
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
   Begin VB.PictureBox fraDetails 
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
      Height          =   4545
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   2505
      TabIndex        =   47
      Top             =   60
      Width           =   2535
      Begin VB.TextBox textSearch 
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
         Left            =   30
         MaxLength       =   35
         TabIndex        =   48
         Top             =   330
         Width           =   2445
      End
      Begin MSComctlLib.ListView lstServiceAdvisor 
         Height          =   3645
         Left            =   30
         TabIndex        =   49
         Top             =   780
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   6429
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "EmpMas.frx":08CA
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   10425
         _Version        =   655364
         _ExtentX        =   18389
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "SEARCH"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
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
      Height          =   945
      Left            =   3780
      ScaleHeight     =   945
      ScaleWidth      =   8715
      TabIndex        =   24
      Top             =   4650
      Width           =   8715
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
         Left            =   7890
         MouseIcon       =   "EmpMas.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   7170
         MouseIcon       =   "EmpMas.frx":0EE4
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":1036
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   6450
         MouseIcon       =   "EmpMas.frx":139C
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":14EE
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   5730
         MouseIcon       =   "EmpMas.frx":1819
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":196B
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   5010
         MouseIcon       =   "EmpMas.frx":1CC7
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":1E19
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   4290
         MouseIcon       =   "EmpMas.frx":212C
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":227E
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   3570
         MouseIcon       =   "EmpMas.frx":2578
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":26CA
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   2850
         MouseIcon       =   "EmpMas.frx":2A22
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":2B74
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdTrainPlan 
         Caption         =   "Training/seminar Plan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   14
         Top             =   450
         Width           =   2715
      End
      Begin VB.CommandButton cmdViewTrain 
         Caption         =   "Training/seminar attended"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   13
         Top             =   60
         Width           =   2715
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
      Left            =   10905
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   25
      Top             =   4635
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
         Left            =   780
         MouseIcon       =   "EmpMas.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
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
         Left            =   60
         MouseIcon       =   "EmpMas.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport rptTS 
      Left            =   60
      Top             =   4620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Service Advisor's Master List"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label lblTS 
      BackColor       =   &H000000FF&
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
      Left            =   2610
      TabIndex        =   28
      Top             =   5250
      Visible         =   0   'False
      Width           =   1920
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
      Left            =   2670
      TabIndex        =   23
      Top             =   4980
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMS_SA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpNo                                            As ADODB.Recordset
Dim AddorEdit                                          As String

Function FINDDEPTCODE() As String
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DEPTCODE FROM HRMS_DEPARTMENT WHERE DEPTNAME = " & N2Str2Null(cboDept) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FINDDEPTCODE = Null2String(RSTMP!DeptCode)
    Else
        FINDDEPTCODE = ""
    End If
    Set RSTMP = Nothing
End Function

Function EX_POSITION() As String
    Dim vPOSITION                                      As String
    Dim xPOSITION                                      As String
    vPOSITION = "0000000"

    If Check1.Value = 1 Then xPOSITION = xPOSITION & "1"
    If Check1.Value = 0 Then xPOSITION = xPOSITION & "0"
    If Check2.Value = 1 Then xPOSITION = xPOSITION & "1"
    If Check2.Value = 0 Then xPOSITION = xPOSITION & "0"
    If Check3.Value = 1 Then xPOSITION = xPOSITION & "1"
    If Check3.Value = 0 Then xPOSITION = xPOSITION & "0"
    If Check4.Value = 1 Then xPOSITION = xPOSITION & "1"
    If Check4.Value = 0 Then xPOSITION = xPOSITION & "0"
    If Check5.Value = 1 Then xPOSITION = xPOSITION & "1"
    If Check5.Value = 0 Then xPOSITION = xPOSITION & "0"
    If Check6.Value = 1 Then xPOSITION = xPOSITION & "1"
    If Check6.Value = 0 Then xPOSITION = xPOSITION & "0"
    If Check7.Value = 1 Then xPOSITION = xPOSITION & "1"
    If Check7.Value = 0 Then xPOSITION = xPOSITION & "0"
    If Check8.Value = 1 Then xPOSITION = xPOSITION & "1"
    If Check8.Value = 0 Then xPOSITION = xPOSITION & "0"

    EX_POSITION = vPOSITION & xPOSITION
End Function

Function FindDeptName(vDCODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select DeptName From HRMS_DEPARTMENT Where DeptCode = '" & vDCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindDeptName = Null2String(RSTMP!Deptname)
    Else
        FindDeptName = ""
    End If

    Set RSTMP = Nothing
End Function

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Sub FillPosition()
    '    cboPosition.AddItem "GJ Technician Master"
    '    cboPosition.AddItem "GJ Technician Expert"
    '    cboPosition.AddItem "GJ Technician Certified"
    '    cboPosition.AddItem "GJ Technician New"
    '    cboPosition.AddItem "BP Technician Paint"
    '    cboPosition.AddItem "BP Technician Tinsmith"

    '    cboPosition.AddItem "SA Master"
    '    cboPosition.AddItem "SA Certified"
    '    cboPosition.AddItem "SA New"
    '    cboPosition.AddItem "Foreman"
    '    cboPosition.AddItem "Warranty"
    '    cboPosition.AddItem "In-House Instructor"
    '    cboPosition.AddItem "Billing Staff"
    '    cboPosition.AddItem "Others"
End Sub

Sub Filldepartment()
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select DEPTNAME From HRMS_DEPARTMENT Order By DEPTNAME")
    cboDept.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboDept.AddItem Null2String(RSTMP!Deptname)

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Sub

'Sub fillListView()
'    Dim rsTmp As ADODB.Recordset
'    Dim ITEM As ListItem
'    Set rsTmp = gconDMIS.Execute("Select * from CSMS_SERVICE_ADVISER_TECHNICIAN where TECH_OR_SA = '" & frmMainMenu.lblTS.Caption & "' Order By Empno ASC")
'    lstServiceAdvisor.ListItems.Clear
'    If Not (rsTmp.BOF And rsTmp.EOF) Then
'        Do While Not rsTmp.EOF
'            Set ITEM = lstServiceAdvisor.ListItems.Add(, , Null2String(rsTmp!EmpName))
'            ITEM.SubItems(1) = Null2String(rsTmp!empno)
'
'            rsTmp.MoveNext
'        Loop
'    End If
'    Set rsTmp = Nothing
'End Sub

Sub initMemvars()
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleInt.Text = ""
    txtENO.Text = ""
    txtEmpNo.Text = ""
    txtBDATE.Text = ""
    txtAddress.Text = ""
    txtReligion.Text = ""
    txtTELE.Text = ""
    txtCITI.Text = ""
    txtDHIRED.Text = ""
    txtDRES.Text = ""
    'cboPosition.Text = ""
    cboDept.Text = ""

    'UPDATED BY: JUN---------
    'DATE UPDATED: 12-16-2008
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    Check7.Value = 0
    Check8.Value = 0
    'UPDATED BY: JUN---------

End Sub

Sub StoreMemVars()
    If Not (rsEmpNo.EOF And rsEmpNo.BOF) Then
        labid.Caption = rsEmpNo!ID
        txtENO.Text = Null2String(rsEmpNo!EMPNO)
        txtLastName.Text = Null2String(rsEmpNo!lastname)
        txtFirstName.Text = Null2String(rsEmpNo!Firstname)
        txtMiddleInt.Text = Null2String(rsEmpNo!MIDDLEINT)
        txtEmpNo.Text = Null2String(rsEmpNo!code)
        'txtPOSITION.Text = Null2String(rsEmpNo!PostionS)
        cboDept.Text = FindDeptName(Null2String(rsEmpNo!DeptCode))

        'lblDepCode.Caption = Null2String(rsEmpNo!DeptCode)
        'lblEMPTYPE.Caption = Null2String(rsEmpNo!EMPLEVEL)
        'cboPosition.Text = Null2String(rsEmpNo!Position)
        'txtEmpNo.Text = Left(Null2String(Null2String(rsEmpNo!lastname)), 1) & Left(Null2String(rsEmpNo!Firstname), 1) & Left(Null2String(rsEmpNo!MIDDLENAME), 1)
        'txtAddress.Text = Null2String(rsEmpNo!Address)
        'txtTELE.Text = Null2String(rsEmpNo!Telephone)
        'txtBDATE.Text = Null2String(rsEmpNo!BirthDate)
        'txtReligion.Text = Null2String(rsEmpNo!RELIGION)
        'txtCITI.Text = Null2String(rsEmpNo!CITIZEN)
        'txtDHIRED.Text = Null2String(rsEmpNo!DATEHIRED)
        'txtDRES.Text = Null2String(rsEmpNo!RESIGNED)

        'cboPosition.Text = Null2String(rsEmpNo!Position)
        Call DisplayPosition(Null2String(rsEmpNo!EMPNO))


        'If Null2String(rsEmpNo!Active) = "YES" Then chkActive.Value = True
        'If Not Null2String(rsEmpNo!Active) = "YES" Then chkActive.Value = False
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

''Sub DisplayPosition(vEMPNO As String)
''    On Error Resume Next
''    Dim RSTMP As New ADODB.Recordset
''
''    Set RSTMP = gconDMIS.Execute("SELECT CSMS_POSITION FROM HRMS_EMPINFO WHERE EMPNO = '" & vEMPNO & "'")
''    If Not (RSTMP.BOF And RSTMP.EOF) Then
''        If Not Null2String(RSTMP!CSMS_POSITION) = "" Then
''            Check1.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 8, 1)
''            Check2.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 9, 1)
''            Check3.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 10, 1)
''            Check4.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 11, 1)
''            Check5.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 12, 1)
''            Check6.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 13, 1)
''            Check7.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 14, 1)
''            Check8.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 15, 1)
''        Else
''            GoTo HELLO
''        End If
''    Else
''HELLO:
''        Check1.Value = 0: Check2.Value = 0: Check3.Value = 0: Check4.Value = 0
''        Check5.Value = 0: Check6.Value = 0: Check7.Value = 0: Check8.Value = 0
''    End If
''    Set RSTMP = Nothing
''End Sub

Sub DisplayPosition(vEMPNO As String)
    On Error Resume Next
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT CSMS_POSITION FROM HRMS_EMPINFO WHERE EMPNO = '" & vEMPNO & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Not Null2String(RSTMP!CSMS_POSITION) = "" Then
            Check1.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 1, 1)
            Check2.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 2, 1)
            Check3.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 3, 1)
            Check4.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 4, 1)
            Check5.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 5, 1)
            Check6.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 6, 1)
            Check7.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 7, 1)
            Check8.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 8, 1)
        Else
            Check1.Value = 0: Check2.Value = 0: Check3.Value = 0: Check4.Value = 0
            Check5.Value = 0: Check6.Value = 0: Check7.Value = 0: Check8.Value = 0
        End If
    Else
        'UPADATED BY: JUN
        'DATE UPDATED: 12-16-2008
        Dim rsCsmsPos                                  As ADODB.Recordset
        Set rsCsmsPos = gconDMIS.Execute("Select * from CSMS_EMPINFO where IS_SERVICE_ADVISER = 1 ")
        If Not rsCsmsPos.EOF And Not rsCsmsPos.BOF Then
            Check1.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 1, 1)
            Check2.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 2, 1)
            Check3.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 3, 1)
            Check4.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 4, 1)
            Check5.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 5, 1)
            Check6.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 6, 1)
            Check7.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 7, 1)
            Check8.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 8, 1)
        Else
            Check1.Value = 0: Check2.Value = 0: Check3.Value = 0: Check4.Value = 0
            Check5.Value = 0: Check6.Value = 0: Check7.Value = 0: Check8.Value = 0
        End If
        Set rsCsmsPos = Nothing
    End If
    Set RSTMP = Nothing
End Sub

Sub rsRefresh()
    Set rsEmpNo = New ADODB.Recordset
    rsEmpNo.Open "select * from CSMS_VW_EMPNO Order By LASTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'rsEmpNo.Open "select * from HRMS_EMPINFO Where IS_SERVICE_ADVISER = 1 Order By LASTNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsServiceAdvisor                               As ADODB.Recordset
    lstServiceAdvisor.Enabled = False
    lstServiceAdvisor.Sorted = False: lstServiceAdvisor.ListItems.Clear
    Set rsServiceAdvisor = New ADODB.Recordset
    'If frmMainMenu.lblTS.Caption = "TECH" Then
    '    Set rsServiceAdvisor = gconDMIS.Execute("select LASTNAME + ', ' + FIRSTNAME AS TECHNAME,id from CSMS_empINFO where is_technician = '1' Order by LASTNAME asc")
    'Else
    'Set rsServiceAdvisor = gconDMIS.Execute("select LASTNAME + ', ' + FIRSTNAME AS TECHNAME,id from HRMS_empINFO where is_serVICE_ADVISER = '1' Order by LASTNAME asc")
    'End If

    Set rsServiceAdvisor = gconDMIS.Execute("select LASTNAME + ', ' + FIRSTNAME AS TECHNAME,id from CSMS_VW_empNO Order by LASTNAME asc")
    If Not (rsServiceAdvisor.EOF And rsServiceAdvisor.BOF) Then
        Listview_Loadval Me.lstServiceAdvisor.ListItems, rsServiceAdvisor
        lstServiceAdvisor.Refresh
        lstServiceAdvisor.Enabled = True
    End If
    Set rsServiceAdvisor = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsServiceAdvisor                               As ADODB.Recordset
    lstServiceAdvisor.Sorted = False: lstServiceAdvisor.ListItems.Clear
    lstServiceAdvisor.Enabled = False
    Set rsServiceAdvisor = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))

    'If frmMainMenu.lblTS.Caption = "TECH" Then
    '    Set rsServiceAdvisor = gconDMIS.Execute("select LASTNAME + ', ' + FIRSTname AS TECHNAME,ID from CSMS_empinfo where LASTNAME + ', ' + FIRSTNAME Like '%" & XXX & "%' AND is_technician = '1' ORDER BY LASTNAME")
    'Else
    'Set rsServiceAdvisor = gconDMIS.Execute("select LASTNAME + ', ' + FIRSTname AS SANAME,ID from HRMS_empINFO where Lastname + ', ' + FirstName like '%" & XXX & "%' AND  IS_SERVICE_ADVISER = '1' ORDER BY LASTNAME")
    'End If

    Set rsServiceAdvisor = gconDMIS.Execute("select LASTNAME + ', ' + FIRSTNAME AS TECHNAME,id from CSMS_VW_empNO where Lastname + ', ' + FirstName like '%" & XXX & "%' Order by LASTNAME asc")
    If Not (rsServiceAdvisor.EOF And rsServiceAdvisor.BOF) Then
        Listview_Loadval Me.lstServiceAdvisor.ListItems, rsServiceAdvisor
        lstServiceAdvisor.Refresh
        lstServiceAdvisor.Enabled = True
    End If
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "SERVICE ADVISOR") = False Then Exit Sub

    Screen.MousePointer = 11

    'If frmMainMenu.lblTS.Caption = "SA" Then
    'rptTS.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    'rptTS.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    'rptTS.Formulas(2) = "ListOF = '" & "List Of Service Adviser" & "'"
    'rptTS.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
    'PrintSQLReport rptTS, "C:\A K 1 N\DOCUMENTS\SA AND TECHNICIAN TABLE\" & "LIST OF EMPLOYEE.rpt", "{CSMS_SERVICE_ADVISER_TECHNICIAN.TECH_OR_SA} = 'SA'", CSMS_REPORT_CONNECTION, 1
    'Else
    'rptTS.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    'rptTS.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    'rptTS.Formulas(2) = "ListOF = '" & "List Of Technician" & "'"
    'rptTS.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
    'PrintSQLReport rptTS, "C:\A K 1 N\DOCUMENTS\SA AND TECHNICIAN TABLE\" & "LIST OF EMPLOYEE.rpt", "{CSMS_SERVICE_ADVISER_TECHNICIAN.TECH_OR_SA} = 'TECH'", CSMS_REPORT_CONNECTION, 1
    'End If

    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next

    'MsgBox "To add an Employee, use the HRMS module", vbInformation, "CSMS"
    'Exit Sub

    If Function_Access(LOGID, "ACESS_ADD", "SERVICE ADVISOR") = False Then Exit Sub
    AddorEdit = "ADD"
    cmdViewTrain.Visible = False
    cmdTrainPlan.Visible = False

    Frame1.Enabled = True
    fraDetails.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True

    initMemvars
    txtENO.SetFocus
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    fraDetails.Enabled = True

    cmdTrainPlan.Visible = True
    cmdViewTrain.Visible = True

    rsRefresh
    rsEmpNo.MoveFirst
    rsEmpNo.Find "id = " & labid.Caption & ""
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "ACESS_DELETE", "SERVICE ADVISOR") = False Then Exit Sub
    'On Error GoTo Errorcode
    If Not rsEmpNo.BOF Or Not rsEmpNo.EOF Then
        If MsgBox("Delete this Information", vbQuestion + vbYesNo, "Are you sure") = vbYes Then
            Dim RSTMP                                  As New ADODB.Recordset

            Set RSTMP = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = '" & txtENO & "'")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                MsgBox "Cannot Delete Service Advisor in this module." & vbCrLf & "Ask the HR Personnel to delete the employee in the HRMS module", vbInformation, "CSMS"
                Exit Sub
            Else
                SQL_STATEMENT = "delete from CSMS_EMPINFO where id = " & labid
                gconDMIS.Execute SQL_STATEMENT
                Call NEW_LogAudit("X", "SERVICE ADVISOR", SQL_STATEMENT, labid, "", "LASTNAME: " & txtLastName, "", "")

                ShowDeletedMsg
            End If

            textSearch.Text = "a": textSearch.Text = ""
            rsRefresh
            StoreMemVars
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
    If Function_Access(LOGID, "ACESS_EDIT", "SERVICE ADVISOR") = False Then Exit Sub

    AddorEdit = "EDIT"
    'cmdViewTrain.Visible = True
    'cmdTrainPlan.Visible = True

    Frame1.Enabled = True
    fraDetails.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsEmpNo.MoveNext
    If rsEmpNo.EOF Then
        rsEmpNo.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsEmpNo.MovePrevious
    If rsEmpNo.BOF Then
        rsEmpNo.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    Dim vID                                            As Integer
    Dim RSTMP                                          As New ADODB.Recordset

    If IsNull(txtEmpNo.Text) = True Then
        ShowIsRequiredMsg ("Employee no. must not be empty")
        On Error Resume Next
        txtEmpNo.SetFocus
        Exit Sub
    Else
        Dim rsfindDup                                  As ADODB.Recordset
        Set rsfindDup = New ADODB.Recordset
        If AddorEdit = "ADD" Then
            rsfindDup.Open "select NAYM,CODE,EMPNO from CSMS_vw_EMPNO where code = '" & txtEmpNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgBox "Employee Code already exist: " & Null2String(rsfindDup!NAYM) & "", vbExclamation, "CSMS"
                On Error Resume Next
                txtEmpNo.SetFocus
                Exit Sub
            End If
            Set rsfindDup = Nothing

            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select NAYM,CODE,EMPNO from CSMS_vw_EMPNO where EMPNO = '" & txtENO.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgBox "Employee No. already exist: " & Null2String(rsfindDup!NAYM) & "", vbExclamation, "CSMS"
                On Error Resume Next
                txtENO.SetFocus
                Exit Sub
            End If
            Set rsfindDup = Nothing
        Else
            Set rsfindDup = gconDMIS.Execute("SELECT * FROM CSMS_VW_EMPNO WHERE CODE = '" & txtEmpNo & "'")
            If Not (rsfindDup.BOF And rsfindDup.EOF) Then
                If Not labid.Caption = rsfindDup!ID Then
                    MsgBox "Employee Code already exist: " & Null2String(rsfindDup!NAYM) & "", vbExclamation, "CSMS"
                    On Error Resume Next
                    txtEmpNo.SetFocus
                End If
            End If
            Set rsfindDup = Nothing

            Set rsfindDup = New ADODB.Recordset
            Set rsfindDup = gconDMIS.Execute("SELECT * FROM CSMS_VW_EMPNO WHERE EMPNO = '" & txtENO & "'")
            If Not (rsfindDup.BOF And rsfindDup.EOF) Then
                If Not labid.Caption = rsfindDup!ID Then
                    MsgBox "Employee No. already exist: " & Null2String(rsfindDup!NAYM) & "", vbExclamation, "CSMS"
                    On Error Resume Next
                    txtENO.SetFocus
                End If
            End If
            Set rsfindDup = Nothing
        End If
    End If

    If txtLastName.Text = "" Or txtFirstName.Text = "" Then
        MsgSpeechBox "Last Name and First Name is Required"
        On Error Resume Next
        txtLastName.SetFocus
        Exit Sub
    End If

    Dim VTXTCode, VTXTLASTNAME, VTXTFIRSTNAME          As String
    Dim VTXTMiddleInt, VTXTNaym, VTXTEmpNo             As String
    Dim vACTIVE                                        As String
    Dim vPOSITION                                      As String
    Dim vDEPCODE                                       As String
    Dim VENO                                           As String

    vPOSITION = EX_POSITION
    'vPOSITION = N2Str2Null(txtPOSITION)
    vDEPCODE = N2Str2Null(FINDDEPTCODE)

    VTXTLASTNAME = N2Str2Null(txtLastName.Text)
    VTXTFIRSTNAME = N2Str2Null(txtFirstName.Text)
    VTXTMiddleInt = N2Str2Null(txtMiddleInt.Text)
    VTXTEmpNo = N2Str2Null(txtEmpNo.Text)
    VENO = N2Str2Null(txtENO.Text)

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into CSMS_EMPINFO (EMPNO, EMPLEVEL,Firstname, LastName, middleNAME, CSMS_Position, IS_SERVICE_ADVISER) " & _
                      " Values (" & VENO & ",'E'," & VTXTFIRSTNAME & _
                        "," & VTXTLASTNAME & _
                        "," & VTXTMiddleInt & _
                        "," & vPOSITION & _
                        ",'1')"
        gconDMIS.Execute SQL_STATEMENT
        Set RSTMP = gconDMIS.Execute("SELECT ID FROM CSMS_EMPINFO WHERE USERCODE = " & VTXTEmpNo & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            vID = RSTMP!ID
        End If
        Call NEW_LogAudit("A", "SERVICE ADVISOR", SQL_STATEMENT, N2Str2Null(vID), "", "NAME: " & txtLastName, "", "")

        ShowSuccessFullyAdded
    Else
        Dim rsCHECKER                                  As New ADODB.Recordset

        Set rsCHECKER = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = " & VENO & "")
        If Not (rsCHECKER.BOF And rsCHECKER.EOF) Then
            SQL_STATEMENT = "update HRMS_EMPINFO set " & _
                            "EMPNO = " & VENO & _
                            ",LASTNAME = " & VTXTLASTNAME & _
                            ",FIRSTNAME = " & VTXTFIRSTNAME & _
                            ",MIDDLENAME = " & VTXTMiddleInt & _
                            ",CSMS_POSITION = " & vPOSITION & _
                          " WHERE ID = " & labid.Caption & ""
            gconDMIS.Execute SQL_STATEMENT
        Else
            SQL_STATEMENT = "update CSMS_EMPINFO set " & _
                            "EMPNO = " & VENO & _
                            ",LASTNAME = " & VTXTLASTNAME & _
                            ",FIRSTNAME = " & VTXTFIRSTNAME & _
                            ",MIDDLENAME = " & VTXTMiddleInt & _
                            ",CSMS_POSITION = " & vPOSITION & _
                          " WHERE ID = " & labid.Caption & ""
            gconDMIS.Execute SQL_STATEMENT
        End If
        vID = labid.Caption

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "SERVICE ADVISOR", SQL_STATEMENT, N2Str2Null(vID), "", "NAME: " & txtLastName, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        ShowSuccessFullyUpdated
    End If

    textSearch.Text = "a": textSearch.Text = ""
    rsRefresh

    On Error Resume Next

    rsEmpNo.Find "ID = " & vID
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdTrainPlan_Click()
    'frmCSMS_SATrainingPlan.Caption = "SERVICE ADVISER TRAIN AND SEMINAR PLAN TO ATTEND"
    If Function_Access(LOGID, "ACESS_ADD", "SERVICE ADVISOR") = False Then Exit Sub
    frmCSMS_SATrainingPlan.Show 1
End Sub


Private Sub cmdViewTrain_Click()
    'frmCSMS_SATRAIN.Caption = "SERVICE ADVISER TRAIN AND SEMINAR ATTENDED"
    If Function_Access(LOGID, "ACESS_ADD", "SERVICE ADVISOR") = False Then Exit Sub
    frmCSMS_SATRAIN.Show 1
End Sub

Private Sub Command1_Click()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim rsHRMS                                         As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    'gconDMIS.Execute ("Delete from CSMS_SERVICE_ADVISER_TECHNICIAN Where FromWhat = '" & "HRMS" & "'  And TECH_OR_SA = '" & frmMainMenu.lblTS.Caption & "' ")

    Set rsHRMS = gconDMIS.Execute("Select * from HRMS_EMPINFO Where IS_Service_Adviser = " & 1 & " and salarycode is not null ORDER BY LASTNAME")


    If Not (rsHRMS.BOF And rsHRMS.EOF) Then
        Do While Not rsHRMS.EOF
            Set RSTMP = gconDMIS.Execute("Select * From CSMS_EMPINFO Where EMpno = '" & rsHRMS!EMPNO & "' And IS_SERVICE_ADVISER = '1'")
            If (RSTMP.EOF And RSTMP.BOF) Then
                'gconDMIS.Execute ("Insert Into CSMS_SERVICE_ADVISER_TECHNICIAN (EMPNO,DeptCode, emp_position, FirstName,LastName,MiddleName,EmpName,FromWhat,TECH_OR_SA,ACTIVE) VALUES(" & N2Str2Null(rsHRMS!empno) & _
                 "," & N2Str2Null(rsHRMS!DeptCode) & "," & N2Str2Null(rsHRMS!Position) & _
                 "," & N2Str2Null(rsHRMS!Firstname) & "," & N2Str2Null(rsHRMS!lastname) & _
                 "," & N2Str2Null(Left(rsHRMS!MIDDLENAME, 1)) & _
                 "," & N2Str2Null(rsHRMS!lastname & ", " & rsHRMS!Firstname) & _
                 ",'" & "HRMS" & "','" & frmMainMenu.lblTS.Caption & "','" & "YES" & "')")
            Else
                'If MsgBox(rsHRMS!lastname & ", " & rsHRMS!Firstname & " - Information Has been changed From HRMS Data, do You Want to Update his Information", vbQuestion + vbYesNo, "Update") = vbYes Then
                '    gconDMIS.Execute ("UPDATE CSMS_SERVICE_ADVISER_TECHNICIAN Set " & _
                     "FirstName = " & N2Str2Null(rsHRMS!Firstname) & _
                     ",DeptCode = " & N2Str2Null(rsHRMS!DeptCode) & _
                     ",emp_Position = " & N2Str2Null(rsHRMS!Position) & _
                     ",lastname = " & N2Str2Null(rsHRMS!lastname) & _
                     ",MiddleName = " & N2Str2Null(Left((rsHRMS!MIDDLENAME), 1)) & _
                     ",EmpName = " & N2Str2Null(rsHRMS!lastname & ", " & rsHRMS!Firstname) & _
                     ",FromWhat = '" & "HRMS" & _
                     "',ACTIVE = '" & "YES" & _
                     "' Where EMPNO =  '" & rsHRMS!empno & "'")
                'End If
            End If
            rsHRMS.MoveNext
        Loop
    End If

    textSearch.Text = "a": textSearch.Text = ""
    rsRefresh
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SERVICE ADVISOR)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "EMPLOYEE INFO", "")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Frame1.Enabled = False
    textSearch.Text = "":                             'Picture3.ZOrder 0

    textSearch.Text = "A": textSearch.Text = ""

    Filldepartment
    FillPosition
    rsRefresh
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub lstServiceAdvisor_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsEmpNo.MoveFirst
    rsEmpNo.Find "ID = " & ITEM.ListSubItems(1) & ""
    StoreMemVars
End Sub

Private Sub lstServiceAdvisor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstServiceAdvisor
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

Private Sub lstServiceAdvisor_DblClick()
    '    cmdEdit.Value = True
End Sub

Private Sub lstServiceAdvisor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstServiceAdvisor.Enabled = True Then
            lstServiceAdvisor.SetFocus
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    If Label7.ForeColor = vbRed Then
        Label7.ForeColor = vbBlack
    Else
        Label7.ForeColor = vbRed
    End If
End Sub

Private Sub txtBDATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890/", KeyAscii)
    End If
End Sub

Private Sub txtFirstName_Change()
    'txtNaym.Text = txtLastName.Text & ", " & txtFirstName.Text & " " & Left(txtMiddleInt.Text, 1)
    txtEmpNo.Text = Left(txtLastName, 1) & Left(txtFirstName.Text, 1) & Left(txtMiddleInt.Text, 1)
End Sub

Private Sub txtLastName_Change()
    'txtNaym.Text = txtLastName.Text & ", " & txtFirstName.Text & " " & Left(txtMiddleInt.Text, 1)
    txtEmpNo.Text = Left(txtLastName, 1) & Left(txtFirstName.Text, 1) & Left(txtMiddleInt.Text, 1)
End Sub

Private Sub txtMiddleInt_Change()
    'txtNaym.Text = txtLastName.Text & ", " & txtFirstName.Text & " " & Left(txtMiddleInt.Text, 1)
    txtEmpNo.Text = Left(txtLastName, 1) & Left(txtFirstName.Text, 1) & Left(txtMiddleInt.Text, 1)
End Sub


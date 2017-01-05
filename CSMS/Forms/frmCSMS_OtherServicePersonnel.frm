VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_OtherServicePersonnel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other Service Personnel"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3720
      ScaleHeight     =   945
      ScaleWidth      =   8715
      TabIndex        =   47
      Top             =   4050
      Width           =   8715
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
         TabIndex        =   57
         Top             =   60
         Width           =   2715
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
         TabIndex        =   56
         Top             =   450
         Width           =   2715
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Move to Previous Record"
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":04B1
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":0603
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Move to Next Record"
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":095B
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":0AAD
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Find a Record"
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":0DA7
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":0EF9
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Add Record"
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":120C
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":135E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Edit Selected Record"
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":16BA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":180C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Delete Selected Record"
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":1B37
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":1C89
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   735
      End
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":1FEF
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":2141
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   735
      End
   End
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
      Height          =   4035
      Left            =   2580
      ScaleHeight     =   4005
      ScaleWidth      =   9765
      TabIndex        =   4
      Top             =   0
      Width           =   9795
      Begin VB.TextBox txtBDATE 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   27
         Top             =   1560
         Width           =   1845
      End
      Begin VB.TextBox txtENO 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   26
         Top             =   360
         Width           =   1845
      End
      Begin VB.TextBox txtEmpNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7830
         MaxLength       =   10
         TabIndex        =   25
         Top             =   330
         Width           =   1845
      End
      Begin VB.TextBox txtMiddleInt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         Top             =   750
         Width           =   1395
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         Top             =   750
         Width           =   2475
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         Top             =   750
         Width           =   2475
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
         Left            =   5850
         TabIndex        =   21
         Top             =   3570
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.ComboBox cboDept 
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
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   3510
         Width           =   3765
      End
      Begin VB.TextBox txtReligion 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1950
         Width           =   4695
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1170
         Width           =   5625
      End
      Begin VB.TextBox txtCITI 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2340
         Width           =   4695
      End
      Begin VB.TextBox txtDHIRED 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   16
         Top             =   1950
         Width           =   1845
      End
      Begin VB.TextBox txtDRES 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   15
         Top             =   2340
         Width           =   1845
      End
      Begin VB.TextBox txtTELE 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1170
         Width           =   1845
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   1080
         ScaleHeight     =   675
         ScaleWidth      =   6795
         TabIndex        =   5
         Top             =   2760
         Width           =   6825
         Begin VB.CheckBox Check8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Others"
            Height          =   255
            Left            =   4830
            TabIndex        =   13
            Top             =   330
            Width           =   2025
         End
         Begin VB.CheckBox Check7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Billing Staff"
            Height          =   255
            Left            =   4830
            TabIndex        =   12
            Top             =   60
            Width           =   2025
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "In-House Instructor"
            Height          =   255
            Left            =   2790
            TabIndex        =   11
            Top             =   330
            Width           =   2025
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Warranty"
            Height          =   255
            Left            =   2790
            TabIndex        =   10
            Top             =   60
            Width           =   2025
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Foreman"
            Height          =   255
            Left            =   1530
            TabIndex        =   9
            Top             =   330
            Width           =   1245
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SA New"
            Height          =   255
            Left            =   1530
            TabIndex        =   8
            Top             =   60
            Width           =   1005
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SA Certified"
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   330
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SA Master"
            Height          =   255
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1425
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   10425
         _Version        =   655364
         _ExtentX        =   18389
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "INFORMATION"
         ForeColor       =   14606302
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
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
         TabIndex        =   45
         Top             =   2790
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   44
         Top             =   3540
         Width           =   975
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
         TabIndex        =   43
         Top             =   2820
         Width           =   675
      End
      Begin MSForms.CheckBox chkActive 
         Height          =   315
         Left            =   5700
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   480
         Width           =   720
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
         TabIndex        =   40
         Top             =   420
         Width           =   450
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
         TabIndex        =   39
         Top             =   840
         Width           =   1095
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
         TabIndex        =   37
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   345
         TabIndex        =   36
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Birthdate"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   300
         TabIndex        =   34
         Top             =   1650
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   6825
         TabIndex        =   33
         Top             =   1290
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   345
         TabIndex        =   32
         Top             =   1260
         Width           =   690
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
         TabIndex        =   31
         Top             =   30
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   6810
         TabIndex        =   30
         Top             =   2070
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Resigned"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   6480
         TabIndex        =   29
         Top             =   2460
         Width           =   1245
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
         TabIndex        =   28
         Top             =   3180
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.PictureBox fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4545
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
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
         TabIndex        =   1
         Top             =   330
         Width           =   2445
      End
      Begin MSComctlLib.ListView lstServiceAdvisor 
         Height          =   3645
         Left            =   30
         TabIndex        =   2
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":24A7
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
         TabIndex        =   3
         Top             =   0
         Width           =   10425
         _Version        =   655364
         _ExtentX        =   18389
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "SEARCH"
         ForeColor       =   14606302
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.01
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
   Begin Crystal.CrystalReport rptTS 
      Left            =   0
      Top             =   4560
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
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   10845
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   58
      Top             =   4035
      Width           =   1800
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
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":2609
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":275B
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
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
         Left            =   780
         MouseIcon       =   "frmCSMS_OtherServicePersonnel.frx":2AAB
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_OtherServicePersonnel.frx":2BFD
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   2610
      TabIndex        =   62
      Top             =   4050
      Visible         =   0   'False
      Width           =   285
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
      Left            =   2550
      TabIndex        =   61
      Top             =   4320
      Visible         =   0   'False
      Width           =   1920
   End
End
Attribute VB_Name = "frmCSMS_OtherServicePersonnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

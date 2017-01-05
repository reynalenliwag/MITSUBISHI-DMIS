VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_TECH 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TECHNICIAN DATA ENTRY"
   ClientHeight    =   5895
   ClientLeft      =   720
   ClientTop       =   330
   ClientWidth     =   11610
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
   Icon            =   "EmpMas1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   11610
   Begin VB.PictureBox fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   30
      ScaleHeight     =   4875
      ScaleWidth      =   2535
      TabIndex        =   42
      Top             =   30
      Width           =   2565
      Begin VB.TextBox textSearch 
         Appearance      =   0  'Flat
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
         Left            =   60
         MaxLength       =   35
         TabIndex        =   43
         Top             =   330
         Width           =   2445
      End
      Begin MSComctlLib.ListView lstServiceAdvisor 
         Height          =   4125
         Left            =   60
         TabIndex        =   44
         Top             =   720
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   7276
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "EmpMas1.frx":08CA
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   4762
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
         TabIndex        =   45
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
   Begin VB.PictureBox Frame1 
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
      Height          =   4995
      Left            =   2640
      ScaleHeight     =   4965
      ScaleWidth      =   8895
      TabIndex        =   30
      Top             =   30
      Width           =   8925
      Begin VB.TextBox txtPOSITIOn 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1140
         TabIndex        =   10
         Top             =   3150
         Width           =   3105
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   1140
         ScaleHeight     =   1305
         ScaleWidth      =   6285
         TabIndex        =   54
         Top             =   3540
         Width           =   6315
         Begin VB.CheckBox Check8 
            Caption         =   "Quick Service Technician"
            Height          =   255
            Left            =   60
            TabIndex        =   64
            Top             =   900
            Width           =   2955
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   5520
            Top             =   840
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Contractor Technician"
            Height          =   255
            Left            =   4290
            TabIndex        =   61
            Top             =   60
            Width           =   2025
         End
         Begin VB.CheckBox Check6 
            Caption         =   "BP In-HouseTechnician Tinsmist"
            Height          =   255
            Left            =   2220
            TabIndex        =   60
            Top             =   630
            Width           =   2955
         End
         Begin VB.CheckBox Check5 
            Caption         =   "BP In-House Technician Paint"
            Height          =   255
            Left            =   2220
            TabIndex        =   59
            Top             =   330
            Width           =   2865
         End
         Begin VB.CheckBox Check4 
            Caption         =   "GJ Technician New"
            Height          =   255
            Left            =   2220
            TabIndex        =   58
            Top             =   60
            Width           =   1815
         End
         Begin VB.CheckBox Check3 
            Caption         =   "GJ Technician Certified"
            Height          =   255
            Left            =   60
            TabIndex        =   57
            Top             =   630
            Width           =   2055
         End
         Begin VB.CheckBox Check2 
            Caption         =   "GJ Technician Expert"
            Height          =   255
            Left            =   60
            TabIndex        =   56
            Top             =   330
            Width           =   1845
         End
         Begin VB.CheckBox Check1 
            Caption         =   "GJ Technician Master"
            Height          =   255
            Left            =   60
            TabIndex        =   55
            Top             =   60
            Width           =   1935
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
            Left            =   60
            TabIndex        =   63
            Top             =   960
            Visible         =   0   'False
            Width           =   3585
         End
      End
      Begin VB.TextBox txtBDATE 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1980
         Width           =   1845
      End
      Begin VB.TextBox txtReligion 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2370
         Width           =   3765
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1590
         Width           =   3765
      End
      Begin VB.TextBox txtCITI 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2760
         Width           =   3765
      End
      Begin VB.TextBox txtDHIRED 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   11
         Top             =   2340
         Width           =   2235
      End
      Begin VB.TextBox txtDRES 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2730
         Width           =   2265
      End
      Begin VB.TextBox txtTELE 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6540
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1560
         Width           =   2235
      End
      Begin VB.CommandButton Command1 
         Caption         =   "XXX UPLOAD TECHNICIAN XXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5430
         TabIndex        =   41
         Top             =   4980
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.TextBox txtLastName 
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
         Height          =   345
         Left            =   1140
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   780
         Width           =   3735
      End
      Begin VB.TextBox txtFirstName 
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
         Height          =   345
         Left            =   1140
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1170
         Width           =   3735
      End
      Begin VB.TextBox txtMiddleInt 
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
         Height          =   345
         Left            =   6570
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   750
         Width           =   2235
      End
      Begin VB.TextBox txtEmpNo 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1845
      End
      Begin VB.TextBox txtENO 
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
         Height          =   345
         Left            =   1140
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   1845
      End
      Begin VB.TextBox cboDept 
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
         Height          =   345
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4950
         Visible         =   0   'False
         Width           =   3105
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
         Left            =   7560
         TabIndex        =   62
         Top             =   3780
         Visible         =   0   'False
         Width           =   1125
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
         Left            =   405
         TabIndex        =   53
         Top             =   2490
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
         Left            =   180
         TabIndex        =   52
         Top             =   2850
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
         Left            =   360
         TabIndex        =   51
         Top             =   2100
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
         Left            =   5520
         TabIndex        =   50
         Top             =   1680
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
         Left            =   405
         TabIndex        =   49
         Top             =   1680
         Width           =   690
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
         Left            =   5550
         TabIndex        =   48
         Top             =   2460
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
         Left            =   5220
         TabIndex        =   47
         Top             =   2850
         Width           =   1245
      End
      Begin VB.Label labid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   210
         Left            =   8220
         TabIndex        =   46
         Top             =   30
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
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
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   900
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   39
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
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
         Height          =   225
         Index           =   0
         Left            =   5430
         TabIndex        =   38
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   225
         Index           =   0
         Left            =   6435
         TabIndex        =   37
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. No"
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
         Height          =   225
         Index           =   0
         Left            =   345
         TabIndex        =   36
         Top             =   450
         Width           =   720
      End
      Begin MSForms.CheckBox chkActive 
         Height          =   315
         Left            =   4830
         TabIndex        =   35
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
         Height          =   225
         Index           =   0
         Left            =   420
         TabIndex        =   34
         Top             =   3240
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   5040
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
         Left            =   7560
         TabIndex        =   32
         Top             =   3390
         Visible         =   0   'False
         Width           =   1125
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   8895
         _Version        =   655364
         _ExtentX        =   15690
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "INFORMATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Left            =   2820
      ScaleHeight     =   945
      ScaleWidth      =   9135
      TabIndex        =   25
      Top             =   5010
      Width           =   9135
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   8010
         MouseIcon       =   "EmpMas1.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   7290
         MouseIcon       =   "EmpMas1.frx":0EE4
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":1036
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   6570
         MouseIcon       =   "EmpMas1.frx":139C
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":14EE
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   5850
         MouseIcon       =   "EmpMas1.frx":1819
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":196B
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   5130
         MouseIcon       =   "EmpMas1.frx":1CC7
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":1E19
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   4410
         MouseIcon       =   "EmpMas1.frx":212C
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":227E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   3690
         MouseIcon       =   "EmpMas1.frx":2578
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":26CA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   2970
         MouseIcon       =   "EmpMas1.frx":2A22
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":2B74
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdViewTrain 
         Caption         =   "Training/seminar attended"
         Height          =   435
         Left            =   330
         TabIndex        =   14
         Top             =   60
         Width           =   2655
      End
      Begin VB.CommandButton cmdTrainPlan 
         Caption         =   "Training/seminar Plan"
         Height          =   375
         Left            =   330
         TabIndex        =   15
         Top             =   480
         Width           =   2655
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
      Left            =   10065
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   26
      Top             =   4995
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   780
         MouseIcon       =   "EmpMas1.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   60
         MouseIcon       =   "EmpMas1.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas1.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport rptTS 
      Left            =   360
      Top             =   5430
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
      Left            =   810
      TabIndex        =   29
      Top             =   5130
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
      Left            =   450
      TabIndex        =   24
      Top             =   5100
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMS_TECH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpNo                                            As ADODB.Recordset
Dim AddorEdit                                          As String

Function EX_POSITION() As String
    Dim vPOSITION                                      As String
    Dim xPOSITION                                      As String
    vPOSITION = "00000000"

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

    EX_POSITION = xPOSITION & vPOSITION
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

Sub Filldepartment()
    '    Dim rsTmp As New ADODB.Recordset

    '    Set rsTmp = gconDMIS.Execute("Select DEPTNAME From HRMS_DEPARTMENT Order By DEPTNAME")
    '    cboDept.Clear
    '    If Not (rsTmp.BOF And rsTmp.EOF) Then
    '        Do While Not rsTmp.EOF
    '            cboDept.AddItem Null2String(rsTmp!Deptname)
    '
    '            rsTmp.MoveNext
    '        Loop
    '    End If
    '
    '    Set rsTmp = Nothing
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
    cboDept.Text = ""
    cboDept.Text = ""

    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    Check6.Value = 0
    Check7.Value = 0
    Check8.Value = 0

End Sub

Sub StoreMemVars()
    If Not (rsEmpNo.EOF And rsEmpNo.BOF) Then
        labid.Caption = rsEmpNo!ID
        txtEmpNo.Text = Null2String(rsEmpNo!Technician)
        txtENO.Text = Null2String(rsEmpNo!EmpNO)

        Dim RSTMP                                      As New ADODB.Recordset
        Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & txtENO & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            txtLastName.Text = Null2String(RSTMP!lastname)
            txtFirstName.Text = Null2String(RSTMP!Firstname)
            txtMiddleInt.Text = Null2String(RSTMP!MIDDLENAME)
            lblEMPTYPE.Caption = N2Str2Null(RSTMP!EMPLEVEL)
            Call DisplayPosition(Null2String(rsEmpNo!EmpNO))
        Else

            Set RSTMP = New ADODB.Recordset
            Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_EMPINFO WHERE EMPNO = '" & txtENO & "'")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                txtLastName.Text = Null2String(RSTMP!lastname)
                txtFirstName.Text = Null2String(RSTMP!Firstname)
                txtMiddleInt.Text = Null2String(RSTMP!MIDDLENAME)
                lblEMPTYPE.Caption = N2Str2Null(RSTMP!EMPLEVEL)
                Call DisplayPosition(Null2String(rsEmpNo!EmpNO))
            End If
        End If
        Set RSTMP = Nothing

        'txtAddress.Text = Null2String(rsEmpNo!Address)
        'txtTELE.Text = Null2String(rsEmpNo!Telephone)
        'txtBDATE.Text = Null2String(rsEmpNo!BirthDate)
        'txtReligion.Text = Null2String(rsEmpNo!RELIGION)
        'txtCITI.Text = Null2String(rsEmpNo!CITIZEN)
        'txtDHIRED.Text = Null2String(rsEmpNo!DATEHIRED)
        'txtDRES.Text = Null2String(rsEmpNo!RESIGNED)

        'Call DisplayPosition(Null2String(rsEmpNo!EMPNO))
        'cboDept.Text = FindDeptName(Null2String(rsEmpNo!DeptCode))

        'If Null2String(rsEmpNo!Active) = "YES" Then chkActive.Value = True
        'If Not Null2String(rsEmpNo!Active) = "YES" Then chkActive.Value = False
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub DisplayPosition(vEMPNO As String)
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
            Check1.Value = 0: Check2.Value = 0
            Check3.Value = 0: Check4.Value = 0
            Check5.Value = 0: Check6.Value = 0
            Check7.Value = 0: Check8.Value = 0

        End If
    Else
        'UPDATED BY: JUN--------------------------------------------------------------------------------------
        'DATE UPDATED: 12-08-2008
        'DESCRIPTION: GET THE CSMS POSITON IN CSMS_EMINFO BECAUSE EMPLOYEE NO. IS NOT EXISTING IN THE HRMS
        Dim rsCSMS                                     As ADODB.Recordset
        Set rsCSMS = gconDMIS.Execute("Select CSMS_POSITION from CSMS_EMPINFO where EMPNO = '" & vEMPNO & "'")
        If Not rsCSMS.EOF And Not rsCSMS.BOF Then
            If Not Null2String(rsCSMS!CSMS_POSITION) = "" Then
                Check1.Value = Mid(Null2String(rsCSMS!CSMS_POSITION), 1, 1)
                Check2.Value = Mid(Null2String(rsCSMS!CSMS_POSITION), 2, 1)
                Check3.Value = Mid(Null2String(rsCSMS!CSMS_POSITION), 3, 1)
                Check4.Value = Mid(Null2String(rsCSMS!CSMS_POSITION), 4, 1)
                Check5.Value = Mid(Null2String(rsCSMS!CSMS_POSITION), 5, 1)
                Check6.Value = Mid(Null2String(rsCSMS!CSMS_POSITION), 6, 1)
                Check7.Value = Mid(Null2String(rsCSMS!CSMS_POSITION), 7, 1)
                Check8.Value = Mid(Null2String(rsCSMS!CSMS_POSITION), 8, 1)
            Else
                Check1.Value = 0: Check2.Value = 0
                Check3.Value = 0: Check4.Value = 0
                Check5.Value = 0: Check6.Value = 0
                Check7.Value = 0: Check8.Value = 0
            End If
        End If
        Set rsCSMS = Nothing
        'UPDATED BY: JUN--------------------------------------------------------------------------------------
    End If
    Set RSTMP = Nothing
End Sub

Sub rsRefresh()
    Set rsEmpNo = New ADODB.Recordset
    'rsEmpNo.Open "select * from HRMS_EMPINFO Where IS_TECHNICIAN = '1' Order By LASTNAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsEmpNo.Open "select * from CSMS_VW_TECHNICIAN Order By TECH_NAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsServiceAdvisor                               As ADODB.Recordset
    lstServiceAdvisor.Enabled = False
    lstServiceAdvisor.Sorted = False: lstServiceAdvisor.ListItems.Clear
    Set rsServiceAdvisor = New ADODB.Recordset

    'Set rsServiceAdvisor = gconDMIS.Execute("select LASTNAME + ', ' + FIRSTNAME AS TECHNAME,id from HRMS_empINFO where is_technician = '1' Order by LASTNAME asc")
    Set rsServiceAdvisor = gconDMIS.Execute("select TECH_NAME ,id from CSMS_VW_TECHNICIAN Order by TECH_NAME asc")
    If Not (rsServiceAdvisor.EOF And rsServiceAdvisor.BOF) Then
        Listview_Loadval Me.lstServiceAdvisor.ListItems, rsServiceAdvisor
        lstServiceAdvisor.Refresh
        lstServiceAdvisor.Enabled = True
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsServiceAdvisor                               As ADODB.Recordset
    lstServiceAdvisor.Sorted = False: lstServiceAdvisor.ListItems.Clear
    lstServiceAdvisor.Enabled = False
    Set rsServiceAdvisor = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Dim ITEM                                           As ListItem

    'Set rsServiceAdvisor = gconDMIS.Execute("select LASTNAME + ', ' + FIRSTname AS TECHNAME,ID from HRMS_empINFO where LASTNAME + ', ' + FIRSTNAME Like '%" & XXX & "%' AND is_technician = '1' ORDER BY LASTNAME")
    Set rsServiceAdvisor = gconDMIS.Execute("select TECH_NAME,id from CSMS_VW_TECHNICIAN WHERE TECH_NAME LIKE '%" & XXX & "%' Order by TECH_NAME asc")
    If Not (rsServiceAdvisor.EOF And rsServiceAdvisor.BOF) Then
        Do While Not rsServiceAdvisor.EOF
            Set ITEM = lstServiceAdvisor.ListItems.Add(, , Null2String(rsServiceAdvisor!TECH_NAME))
            ITEM.SubItems(1) = rsServiceAdvisor!ID

            rsServiceAdvisor.MoveNext
        Loop
        'Listview_Loadval Me.lstServiceAdvisor.ListItems, rsServiceAdvisor
        'lstServiceAdvisor.Refresh
        'lstServiceAdvisor.Enabled = True
    End If
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "TECHNICIAN") = False Then Exit Sub

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

    '    MsgBox "To add an Employee, use the HRMS module", vbInformation, "CSMS"
    '    Exit Sub

    If Function_Access(LOGID, "ACESS_ADD", "TECHNICIAN") = False Then Exit Sub

    AddorEdit = "ADD"

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



    rsRefresh
    rsEmpNo.MoveFirst
    rsEmpNo.Find "id = " & labid.Caption & ""
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "ACESS_DELETE", "TECHNICIAN") = False Then Exit Sub

    On Error GoTo ErrorCode
    If Not rsEmpNo.BOF Or Not rsEmpNo.EOF Then
        If MsgBox("Delete this Information", vbQuestion + vbYesNo, "Are you sure") = vbYes Then
            Dim RSTMP                                  As New ADODB.Recordset

            Set RSTMP = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = '" & txtENO & "'")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                MsgBox "you Cannot Delete This Technician in this module., contact HR and Ask to Delete it on the HRMS Module" & vbCrLf & "Ask the HR Personnel to delete this employee in the HRMS module", vbInformation, "CSMS"
                Exit Sub
            Else
                SQL_STATEMENT = "delete from CSMS_EMPINFO where id = " & labid
                gconDMIS.Execute SQL_STATEMENT

                'NEW LOG AUDIT----------------------------------------------------
                Call NEW_LogAudit("X", "TECHNICIAN", SQL_STATEMENT, labid, "", "EMPNO: " & txtENO, "", "")
                'NEW LOG AUDIT----------------------------------------------------

                ShowDeletedMsg
            End If
            Set RSTMP = Nothing

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
    If Function_Access(LOGID, "ACESS_EDIT", "TECHNICIAN") = False Then Exit Sub

    AddorEdit = "EDIT"



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
    'Picture3.Visible = False
    'Dim findStr As String
    'findStr = InputBoxXP("Please Input Name ...", "Find")
    'If findStr <> "" Then
    '   On Error Resume Next
    '   rsEmpNo.Bookmark = rsFind(rsEmpNo.Clone, "code", findStr).Bookmark
    '   If Err.Number = 3021 Then
    '      On Error GoTo ErrorCode
    '      rsEmpNo.Bookmark = rsFind(rsEmpNo.Clone, "lastname", findStr).Bookmark
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
        MsgSpeechBox "Employee no. must not be empty"
        On Error Resume Next
        txtEmpNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup                              As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select empno,USERCODE,TECH_NAME from CSMS_VW_TECHNICIAN where ltrim(rtrim(USERcode)) = '" & txtEmpNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgBox "Employee Code already exist: " & Null2String(rsfindDup!TECH_NAME) & "", vbExclamation, "CSMS"
                On Error Resume Next
                txtLastName.SetFocus
                Exit Sub
            End If
            Set rsfindDup = Nothing

            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select empno,USERCODE,TECH_NAME from CSMS_VW_TECHNICIAN where ltrim(rtrim(USERcode)) = '" & txtEmpNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgBox "Employee No. already exist: " & Null2String(rsfindDup!TECH_NAME) & "", vbExclamation, "CSMS"
                On Error Resume Next
                txtEmpNo.SetFocus
                Exit Sub
            End If
            Set rsfindDup = Nothing
        
        
            'UPDATED BY: JUN--------------------------------------------------------------------------------------------------------------------------------------------------
            'DATE UPDATED: 01-26-2008
            'DESCRIPTION:CHECK IF THE EMP. NO. IS EXISTING IN THE HR DEPARTMENT.
            Dim rsCheckEmpNo  As ADODB.Recordset
            Set rsCheckEmpNo = New ADODB.Recordset
            rsCheckEmpNo.Open "select EMPNO from HRMS_EMPINFO where ltrim(rtrim(EMPNO)) = '" & txtENO & "' and IS_TECHNICIAN = 1", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCheckEmpNo.EOF And Not rsCheckEmpNo.BOF Then
                MsgBox "Employee No. already exist in HR Department", vbExclamation, "CSMS"
                On Error Resume Next
                txtEmpNo.SetFocus
                Exit Sub
            End If
            Set rsCheckEmpNo = Nothing
            'UPDATED BY: JUN--------------------------------------------------------------------------------------------------------------------------------------------------
        Else

        End If
    End If
    
    'UPDATED BY: JUN-------------------------------------------------------------------------------------------------------------------
    'DATE UPDATED: 1-27-2008
    'DESCRIPTION: VALIDATE EMPNO UPON EDIT
    
    If AddorEdit = "EDIT" Then
        Dim rsEditChekEmpNo As ADODB.Recordset
        Dim rsHRCheckEmpNO As ADODB.Recordset
        Dim rsCsms_Empno_Check As ADODB.Recordset
        Dim TechName As String
        Dim EmpCount As Integer
        
        EmpCount = 0
        Set rsEditChekEmpNo = gconDMIS.Execute("Select EMPNO from CSMS_EMPINFO  where IS_technician = 1 and EMPNO = '" & LTrim(RTrim((rsEmpNo!EmpNO))) & "'")
        If Not rsEditChekEmpNo.EOF And Not rsEditChekEmpNo.BOF Then
            TechName = Null2String(RTrim(LTrim(rsEditChekEmpNo!EmpNO)))
            Do While Not rsEditChekEmpNo.EOF
                EmpCount = EmpCount + 1
                rsEditChekEmpNo.MoveNext
            Loop
            
            If EmpCount = 1 And LTrim(RTrim(txtENO)) = TechName Then
                'ALLOWED TO EDIT
            ElseIf EmpCount = 1 And LTrim(RTrim(txtENO)) <> TechName Then
                    Set rsHRCheckEmpNO = gconDMIS.Execute("Select EMPNO from HRMS_EMPINFO where IS_TECHNICIAN = 1 AND EMPNO = '" & RTrim(LTrim(txtENO)) & "'")
                    If Not rsHRCheckEmpNO.EOF And Not rsHRCheckEmpNO.BOF Then
                        MsgBox "Employee No. already exist in HR Department", vbExclamation, "CSMS"
                        On Error Resume Next
                        txtEmpNo.SetFocus
                        Exit Sub
                    Else
                        Set rsCsms_Empno_Check = gconDMIS.Execute("Select EMPNO from CSMS_EMPINFO where is_technician = 1 and EmpNo = '" & RTrim(LTrim(txtENO)) & "'")
                        If Not rsCsms_Empno_Check.BOF And Not rsCsms_Empno_Check.BOF Then
                            MsgBox "Employee No. already exist", vbExclamation, "CSMS"
                            On Error Resume Next
                            txtEmpNo.SetFocus
                            Exit Sub
                        Else
                            'ALLOWED TO EDIT
                        End If
                    End If
            ElseIf EmpCount > 1 Then
                MsgBox "Employee No." & " '" & TechName & "' " & vbCrLf & "has Duplicate Entry " & "in Service Technician Entry", vbExclamation, "CSMS"
                Exit Sub
            End If
        
        Else
            Set rsHRCheckEmpNO = gconDMIS.Execute("Select EMPNO from HRMS_EMPINFO where IS_TECHNICIAN = 1 AND EMPNO = '" & RTrim(LTrim(rsEmpNo!EmpNO)) & "'")
                    If Not rsHRCheckEmpNO.EOF And Not rsHRCheckEmpNO.BOF Then
                        TechName = Null2String(RTrim(LTrim(rsHRCheckEmpNO!EmpNO)))
                        
                        Do While Not rsHRCheckEmpNO.EOF
                            EmpCount = EmpCount + 1
                            rsHRCheckEmpNO.MoveNext
                        Loop
                    
                        If EmpCount = 1 And LTrim(RTrim(txtENO)) = TechName Then
                            'ALLOWED TO EDIT
                        ElseIf EmpCount = 1 And LTrim(RTrim(txtENO)) <> TechName Then
                                Set rsHRCheckEmpNO = gconDMIS.Execute("Select EMPNO from HRMS_EMPINFO where IS_TECHNICIAN = 1 AND EMPNO = '" & RTrim(LTrim(txtENO)) & "'")
                                If Not rsHRCheckEmpNO.EOF And Not rsHRCheckEmpNO.BOF Then
                                    MsgBox "Employee No. already exist in HR Department", vbExclamation, "CSMS"
                                    On Error Resume Next
                                    txtEmpNo.SetFocus
                                    Exit Sub
                                Else
                                    Set rsCsms_Empno_Check = gconDMIS.Execute("Select EMPNO from CSMS_EMPINFO where is_technician = 1 and EmpNo = '" & RTrim(LTrim(txtENO)) & "'")
                                    If Not rsCsms_Empno_Check.BOF And Not rsCsms_Empno_Check.BOF Then
                                        MsgBox "Employee No. already exist", vbExclamation, "CSMS"
                                        On Error Resume Next
                                        txtEmpNo.SetFocus
                                        Exit Sub
                                    Else
                                        'ALLOWED TO EDIT
                                    End If
                                End If
                        ElseIf EmpCount > 1 Then
                            MsgBox "Employee No." & " '" & TechName & "' " & vbCrLf & "has Duplicate Entry" & " in HRMS", vbExclamation, "CSMS"
                            Exit Sub
                        End If
                    End If
        End If
        
        Set rsCheckEmpNo = Nothing
        Set rsEditChekEmpNo = Nothing
        Set rsCsms_Empno_Check = Nothing
    End If
    'UPDATED BY: JUN-------------------------------------------------------------------------------------------------------------------
    
    
    

    If txtLastName.Text = "" Or txtFirstName.Text = "" Then
        MsgSpeechBox "Last Name and First Name is Required"
        On Error Resume Next
        txtLastName.SetFocus
        Exit Sub
    End If

    'UPDATED BY: JUN--------------------------------------------------------------------------------
    'DATE UPDATED: 12-05-2008
    'DESCRIPTION: VALIDATE THE TECHNICIAN POSITION TWO OR MORE THAN SELECTED POSITION IS NOT ALLOWED
    Dim xCount                                         As Integer
    xCount = 0

    If Check1.Value = 1 Then xCount = xCount + 1
    If Check2.Value = 1 Then xCount = xCount + 1
    If Check3.Value = 1 Then xCount = xCount + 1
    If Check4.Value = 1 Then xCount = xCount + 1
    If Check5.Value = 1 Then xCount = xCount + 1
    If Check6.Value = 1 Then xCount = xCount + 1
    If Check7.Value = 1 Then xCount = xCount + 1
    If Check8.Value = 1 Then xCount = xCount + 1

    If xCount > 1 Then
        MsgBox "You are only allowed to select one Position.", vbInformation, "INFORMATION"
        Exit Sub
    ElseIf xCount = 0 Then
        MsgBox "Pls. Select one Position", vbInformation, "INFORMATION"
        Exit Sub
    End If
    'UPDATED BY: JUN--------------------------------------------------------------------------------


    '    Dim rsHRMS As New ADODB.Recordset
    '    Set rsHRMS = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = '" & txtENO.Text & "'")
    '    If Not (rsHRMS.BOF And rsHRMS.EOF) Then
    '        MsgBox "The Employee No. Your Adding is Already Existing in the HRMS Database", vbInformation, "CSMS"
    '        txtENO.SetFocus
    '        Exit Sub
    '    End If
    '    Set rsHRMS = Nothing
    '
    '    Set rsHRMS = New ADODB.Recordset
    '    Set rsHRMS = gconDMIS.Execute("SELECT LASTNAME,FIRSTNAME,MIDDLENAME FROM HRMS_EMPINFO WHERE LTRIM(RTRIM(LASTNAME)) = '" & txtLastName.Text & "' AND LTRIM(RTRIM(FIRSTNAME)) = '" & txtFirstName & "' AND LTRIM(RTRIM(LEFT(MIDDLENAME,1))) = '" & Left(txtMiddleInt, 1) & "'")
    '    If Not (rsHRMS.BOF And rsHRMS.EOF) Then
    '        MsgBox "The Employee Name Your Adding is Already Existing in the HRMS Database", vbInformation, "CSMS"
    '        txtFirstName.SetFocus
    '        Exit Sub
    '    End If
    '    Set rsHRMS = Nothing


    Dim VTXTCode, VTXTLASTNAME, VTXTFIRSTNAME          As String
    Dim VTXTMiddleInt, VTXTNaym, VTXTEmpNo             As String
    Dim vACTIVE                                        As String
    Dim vPOSITION                                      As String
    Dim vDEPCODE                                       As String
    Dim VENO                                           As String

    VTXTLASTNAME = N2Str2Null(txtLastName.Text)
    VTXTFIRSTNAME = N2Str2Null(txtFirstName.Text)
    VTXTMiddleInt = N2Str2Null(txtMiddleInt.Text)
    VTXTEmpNo = N2Str2Null(txtEmpNo.Text)
    vPOSITION = N2Str2Null(EX_POSITION)
    VENO = N2Str2Null(txtENO.Text)

    'If chkActive.Value = True Then vACTIVE = "YES"
    'If Not chkActive.Value = True Then vACTIVE = "NO"

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = ("Insert into CSMS_EMPINFO" & _
                       " (EMPNO,USERCODE, firstname, LastName, middleName, salarycode, Position, IS_TECHNICIAN)" & _
                       " values (" & VENO & _
                         "," & VTXTEmpNo & _
                         "," & VTXTFIRSTNAME & _
                         "," & VTXTLASTNAME & _
                         "," & VTXTMiddleInt & _
                         ",'" & "L1S1" & _
                         "'," & vPOSITION & _
                         ",'1')")
        gconDMIS.Execute SQL_STATEMENT

        Set RSTMP = gconDMIS.Execute("SELECT ID FROM CSMS_EMPINFO WHERE USERCODE = " & VTXTEmpNo & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            vID = RSTMP!ID
        End If
        Call NEW_LogAudit("A", "SERVICE ADVISER", SQL_STATEMENT, Null2String(vID), "", "LASTNAME: " & VTXTEmpNo & " - " & VTXTLASTNAME, "", "")

        ShowSuccessFullyAdded
    Else
        Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & txtENO & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            SQL_STATEMENT = "update HRMS_EMPINFO set " & _
                            "EMPNO = " & VENO & _
                            ",CSMS_Position = " & vPOSITION & _
                            ",LASTNAME = " & VTXTLASTNAME & _
                            ",FIRSTNAME = " & VTXTFIRSTNAME & _
                            ",MIDDLENAME = " & VTXTMiddleInt & _
                          " WHERE ID = " & labid.Caption & ""
        Else
            SQL_STATEMENT = "update CSMS_EMPINFO set " & _
                            "EMPNO = " & VENO & _
                            ",CSMS_Position = " & vPOSITION & _
                            ",LASTNAME = " & VTXTLASTNAME & _
                            ",FIRSTNAME = " & VTXTFIRSTNAME & _
                            ",MIDDLENAME = " & VTXTMiddleInt & _
                          " WHERE ID = " & labid.Caption & ""
        End If
        gconDMIS.Execute SQL_STATEMENT

        vID = labid.Caption
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "EMPLOYEE INFO", SQL_STATEMENT, labid, "", "EMP NO: " & txtENO, "", "")
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
    Call ShowVBError
    Exit Sub
End Sub

Private Sub cmdTrainPlan_Click()
    If Function_Access(LOGID, "ACESS_ADD", "TECHNICIAN") = False Then Exit Sub
    'frmCSMS_TECHTrainingPlan.Caption = "TECHNICIAN TRAIN AND SEMINAR PLAN TO ATTEND"

    frmCSMS_TECHTrainingPlan.Show 1
End Sub

Private Sub cmdViewTrain_Click()
    If Function_Access(LOGID, "ACESS_ADD", "TECHNICIAN") = False Then Exit Sub
    'frmCSMS_TECHTRAIN.Caption = "TECHNICIAN TRAIN AND SEMINAR ATTENDED"

    frmCSMS_TECHTRAIN.Show 1
End Sub

Private Sub Command1_Click()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim rsHRMS                                         As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    'gconDMIS.Execute ("Delete from CSMS_SERVICE_ADVISER_TECHNICIAN Where FromWhat = '" & "HRMS" & "'  And TECH_OR_SA = '" & frmMainMenu.lblTS.Caption & "' ")

    Set rsHRMS = gconDMIS.Execute("Select * from HRMS_EMPINFO Where IS_Technician = " & 1 & " and salarycode is not null ORDER BY LASTNAME")

    If Not (rsHRMS.BOF And rsHRMS.EOF) Then
        Do While Not rsHRMS.EOF
            Set RSTMP = gconDMIS.Execute("Select * From CSMS_EMPINFO Where EMpno = '" & rsHRMS!EmpNO & "' And IS_TECHNICIAN = '1'")
            If (RSTMP.EOF And RSTMP.BOF) Then
                'gconDMIS.Execute ("Insert Into CSMS_EMPINFO (EMPNO,position, FirstName,LastName,MiddleName,TECH_OR_SA,ACTIVE) VALUES(" & N2Str2Null(rsHRMS!empno) & _
                 "," & N2Str2Null(rsHRMS!DeptCode) & "," & N2Str2Null(rsHRMS!Position) & _
                 "," & N2Str2Null(rsHRMS!Firstname) & "," & N2Str2Null(rsHRMS!lastname) & _
                 "," & N2Str2Null(Left(rsHRMS!MIDDLENAME, 1)) & _
                 "," & N2Str2Null(rsHRMS!lastname & ", " & rsHRMS!Firstname) & _
                 ",'" & "HRMS" & "','1','" & "YES" & "')")
            Else
                If MsgBox(rsHRMS!lastname & ", " & rsHRMS!Firstname & " - Information Has been changed From HRMS Data, do You Want to Update his Information", vbQuestion + vbYesNo, "Update") = vbYes Then
                    'gconDMIS.Execute ("UPDATE CSMS_SERVICE_ADVISER_TECHNICIAN Set " & _
                     "FirstName = " & N2Str2Null(rsHRMS!Firstname) & _
                     ",DeptCode = " & N2Str2Null(rsHRMS!DeptCode) & _
                     ",emp_Position = " & N2Str2Null(rsHRMS!Position) & _
                     ",lastname = " & N2Str2Null(rsHRMS!lastname) & _
                     ",MiddleName = " & N2Str2Null(Left((rsHRMS!MIDDLENAME), 1)) & _
                     ",EmpName = " & N2Str2Null(rsHRMS!lastname & ", " & rsHRMS!Firstname) & _
                     ",FromWhat = '" & "HRMS" & _
                     "',ACTIVE = '" & "YES" & _
                     "' Where EMPNO =  '" & rsHRMS!empno & "'")
                End If
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
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If Picture1.Visible = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TECHNICIAN RECORDS)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "EMPLOYEE INFO", "")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Frame1.Enabled = False
    textSearch.Text = "":                             'Picture3.ZOrder 0

    textSearch.Text = "A": textSearch.Text = ""

    rsRefresh
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub lstServiceAdvisor_DblClick()
    '    Dim ITEM As ListItem
    '    Dim INDEX As Integer

    '    If lstServiceAdvisor.ListItems.Count = 0 Then Exit Sub
    '
    '    INDEX = lstServiceAdvisor.SelectedItem.INDEX
    '
    '    rsEmpNo.MoveFirst
    '    rsEmpNo.Find "ID = " & lstServiceAdvisor.ListItems(INDEX).ListSubItems(1) & ""
    '    StoreMemVars
End Sub

Private Sub lstServiceAdvisor_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsEmpNo.MoveFirst
    rsEmpNo.Find "ID = " & ITEM.ListSubItems(1) & ""
    StoreMemVars
End Sub

Private Sub lstServiceAdvisor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstServiceAdvisor
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


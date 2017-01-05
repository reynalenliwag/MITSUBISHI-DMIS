VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCRIS_TestVehicles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Drive Vehicle Receiving/Returning  Entry"
   ClientHeight    =   6840
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   11040
   ForeColor       =   &H00FCFCFC&
   Icon            =   "EntryTestVehicles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11040
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2070
      Top             =   210
   End
   Begin VB.OptionButton optView 
      Caption         =   "By Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   33
      Top             =   630
      Width           =   2235
   End
   Begin VB.OptionButton optView 
      Caption         =   "By Model"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   32
      Top             =   330
      Width           =   2235
   End
   Begin VB.OptionButton optView 
      Caption         =   "By Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   31
      Top             =   60
      Value           =   -1  'True
      Width           =   2235
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   930
      Width           =   2385
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   5055
      Left            =   120
      TabIndex        =   29
      Top             =   1350
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   2670
      ScaleHeight     =   5565
      ScaleWidth      =   8235
      TabIndex        =   12
      Top             =   150
      Width           =   8235
      Begin VB.Frame Frame1 
         Caption         =   "Vehicle Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   30
         TabIndex        =   38
         Top             =   60
         Width           =   8235
         Begin VB.TextBox txtYeer 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   7260
            TabIndex        =   44
            Tag             =   "@R"
            Text            =   "Text1"
            Top             =   1020
            Width           =   885
         End
         Begin VB.TextBox txtModel 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   840
            TabIndex        =   43
            Tag             =   "@R"
            Text            =   "Text1"
            Top             =   1020
            Width           =   5625
         End
         Begin VB.TextBox txtMake 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   840
            TabIndex        =   42
            Tag             =   "@R"
            Text            =   "Text1"
            Top             =   630
            Width           =   5625
         End
         Begin VB.TextBox txtCode 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   840
            MaxLength       =   6
            TabIndex        =   41
            Tag             =   "@R"
            Text            =   "Text1"
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox txtDescript 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   3330
            TabIndex        =   40
            Tag             =   "@R"
            Text            =   "Text1"
            Top             =   240
            Width           =   4815
         End
         Begin VB.ComboBox cboClass 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F6F5&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   7260
            TabIndex        =   39
            Tag             =   "@R"
            Text            =   "Combo1"
            Top             =   630
            Width           =   885
         End
         Begin VB.Label Label26 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Class"
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
            Height          =   255
            Left            =   6660
            TabIndex        =   52
            Top             =   690
            Width           =   705
         End
         Begin VB.Label Label27 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
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
            Height          =   255
            Left            =   6720
            TabIndex        =   51
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label28 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
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
            Height          =   255
            Left            =   180
            TabIndex        =   50
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label29 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Make"
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
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   690
            Width           =   1245
         End
         Begin VB.Label Label30 
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
            Left            =   240
            TabIndex        =   48
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label31 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
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
            Height          =   255
            Left            =   2190
            TabIndex        =   47
            Top             =   300
            Width           =   1245
         End
         Begin VB.Label Label32 
            Caption         =   "Label9"
            Height          =   315
            Left            =   960
            TabIndex        =   46
            Top             =   1020
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label33 
            Caption         =   "Label9"
            Height          =   345
            Left            =   1050
            TabIndex        =   45
            Top             =   1020
            Visible         =   0   'False
            Width           =   285
         End
      End
      Begin MSComCtl2.DTPicker dtdaterecieved 
         Height          =   390
         Left            =   5655
         TabIndex        =   34
         Top             =   3615
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   55246849
         CurrentDate     =   39141
      End
      Begin VB.TextBox txtNotes 
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
         Height          =   2070
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   3390
         Width           =   3945
      End
      Begin VB.ComboBox cboColor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1230
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   2580
         Width           =   2925
      End
      Begin VB.TextBox txtEngineNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5550
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1830
         Width           =   2595
      End
      Begin VB.TextBox txtVINo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5610
         TabIndex        =   16
         Tag             =   "@R"
         Text            =   "Text1"
         Top             =   3150
         Width           =   2505
      End
      Begin VB.TextBox txtSerialNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5610
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2700
         Width           =   2505
      End
      Begin VB.ComboBox cboSource 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1200
         TabIndex        =   14
         Tag             =   "@R"
         Text            =   "Combo1"
         Top             =   1830
         Width           =   2925
      End
      Begin VB.TextBox txtUnit 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1230
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2190
         Width           =   2925
      End
      Begin MSMask.MaskEdBox txtTaggedPrice 
         Height          =   345
         Left            =   5580
         TabIndex        =   25
         Tag             =   "@R"
         Top             =   2220
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   609
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   7347754
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblStatus 
         Caption         =   "[**]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   4410
         TabIndex        =   36
         Top             =   4170
         Width           =   3765
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Arrived"
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
         Height          =   255
         Left            =   4350
         TabIndex        =   35
         Top             =   3750
         Width           =   1065
      End
      Begin VB.Label Label42 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3090
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tagged Price"
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
         Height          =   255
         Left            =   4260
         TabIndex        =   26
         Top             =   2250
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No"
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
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "VI No"
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
         Height          =   255
         Left            =   4470
         TabIndex        =   23
         Top             =   3180
         Width           =   885
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No"
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
         Height          =   255
         Left            =   4650
         TabIndex        =   22
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Height          =   255
         Left            =   60
         TabIndex        =   21
         Top             =   2610
         Width           =   885
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
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
         Height          =   255
         Left            =   75
         TabIndex        =   20
         Top             =   1830
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
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
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   2220
         Width           =   885
      End
   End
   Begin VB.PictureBox picSave 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9465
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   9
      Top             =   5880
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
         Left            =   30
         MouseIcon       =   "EntryTestVehicles.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   705
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
         Left            =   765
         MouseIcon       =   "EntryTestVehicles.frx":0D6C
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdd 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   4470
      ScaleHeight     =   945
      ScaleWidth      =   8655
      TabIndex        =   0
      Top             =   5880
      Width           =   8655
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return Vehicle"
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
         MouseIcon       =   "EntryTestVehicles.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   60
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
         MouseIcon       =   "EntryTestVehicles.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
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
         Left            =   750
         MouseIcon       =   "EntryTestVehicles.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
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
         Left            =   1448
         MouseIcon       =   "EntryTestVehicles.frx":200F
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":2161
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
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
         Left            =   2190
         MouseIcon       =   "EntryTestVehicles.frx":245B
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":25AD
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
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
         Left            =   2866
         MouseIcon       =   "EntryTestVehicles.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
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
         Left            =   3575
         MouseIcon       =   "EntryTestVehicles.frx":2D6E
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":2EC0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
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
         Left            =   4284
         MouseIcon       =   "EntryTestVehicles.frx":31EB
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":333D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   705
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
         Left            =   4995
         MouseIcon       =   "EntryTestVehicles.frx":36A3
         MousePointer    =   99  'Custom
         Picture         =   "EntryTestVehicles.frx":37F5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   45
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCRIS_TestVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TestDriveID                           As Long
Private rsTestInv                            As Recordset

Private Sub cmdAdd_Click()
    TestDriveID = 0
    InitVars
    picDetail.Enabled = True
    PicAdd.Visible = False
    PicSave.Visible = True
End Sub

Private Sub cmdCancel_Click()
    picDetail.Enabled = False
    PicAdd.Visible = True
    PicSave.Visible = False
    If Not (rsTestInv.EOF Or rsTestInv.BOF) Then
        rsTestInv.MoveLast
        StoreMemVars
    End If

End Sub
Private Sub cmdEdit_Click()
    picDetail.Enabled = True
    PicAdd.Visible = False
    PicSave.Visible = True

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    rsTestInv.MoveNext
    If rsTestInv.EOF Then
        rsTestInv.MoveLast
        'ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsTestInv.MovePrevious
    If rsTestInv.BOF Then
        rsTestInv.MoveFirst
        'ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    If Runvalidation("@R") = False Then: Exit Sub
    Dim code                                 As String
    Dim descript                             As String
    Dim Make                                 As String
    Dim Class                                As String
    Dim yearmodel                            As String
    Dim source                               As String
    Dim unit                                 As String
    Dim Color                                As String
    Dim enginenumber                         As String
    Dim serialno                             As String
    Dim vinnumber                            As String
    Dim taggedprice                          As String
    Dim notes                                As String
    Dim daterecieved                         As String
    Dim Model                                As String


    code = N2Str2Null(txtCode.Text)
    descript = N2Str2Null(txtDescript.Text)
    Make = N2Str2Null(txtMake.Text)
    Class = N2Str2Null(cboClass.Text)
    yearmodel = N2Str2Null(txtYeer.Text)
    source = N2Str2Null(cboSource.Text)
    unit = N2Str2Null(txtUnit.Text)
    Color = N2Str2Null(cboColor.Text)
    enginenumber = N2Str2Null(txtEngineNo.Text)
    serialno = N2Str2Null(txtSerialNo.Text)
    vinnumber = N2Str2Null(txtVINo.Text)
    taggedprice = CDbl(txtTaggedPrice.Text)
    notes = N2Str2Null(txtNotes.Text)
    daterecieved = N2Date2Null(dtdaterecieved.Value)
    Model = N2Date2Null(txtModel.Text)



    Dim temprs                               As ADODB.Recordset
    If TestDriveID = 0 Then
        SQL = "INSERT INTO  " & _
              "CRIS_MRRINV(Code, Descript, Make, Class, Model, YearModel, Source, Unit, Color, EngineNumber, SerialNo, VinNumber, TaggedPrice, Notes, DateReceived) " & _
              "VALUES(@CODE, @DESCRIPT, @MAKE, @CLASS, @MODEL, @YEARMODEL, @SOURCE, @UNIT, @COLOR, @ENGINENUMBER, @SERIALNO, @VINNUMBER, @TAGGEDPRICE, @NOTES, @DATERECIEVED ) " & vbCrLf & " SELECT @@IDENTITY "
    Else

        SQL = "UPDATE CRIS_MRRINV " & _
              "SET Code=@CODE, Descript=@DESCRIPT, Make=@MAKE, Class=@CLASS, Model=@MODEL, YearModel=@YEARMODEL, Source=@SOURCE, Unit=@UNIT, Color=@COLOR, EngineNumber=@ENGINENUMBER, SerialNo=@SERIALNO, VinNumber=@VINNUMBER, TaggedPrice=@TAGGEDPRICE, Notes=@NOTES, DateReceived=@DATERECIEVED " & _
              "WHERE id=@id "
    End If

    SQL = Replace(SQL, "@id", TestDriveID)
    SQL = Replace(SQL, "@CODE", code)
    SQL = Replace(SQL, "@DESCRIPT", descript)
    SQL = Replace(SQL, "@MAKE", Make)
    SQL = Replace(SQL, "@MODEL", Model)
    SQL = Replace(SQL, "@CLASS", Class)
    SQL = Replace(SQL, "@YEARMODEL", yearmodel)
    SQL = Replace(SQL, "@SOURCE", source)
    SQL = Replace(SQL, "@UNIT", unit)
    SQL = Replace(SQL, "@COLOR", Color)
    SQL = Replace(SQL, "@ENGINENUMBER", enginenumber)
    SQL = Replace(SQL, "@SERIALNO", serialno)
    SQL = Replace(SQL, "@VINNUMBER", vinnumber)
    SQL = Replace(SQL, "@TAGGEDPRICE", taggedprice)
    SQL = Replace(SQL, "@NOTES", notes)
    SQL = Replace(SQL, "@DATERECIEVED", daterecieved)
    Set temprs = gconDMIS.Execute(SQL)
    If LOGID <= 0 Then
        MessagePop RecSave, "Record Added ", "Profile Sucessfully Added"
    Else
        MessagePop RecSave, "RecordSaved", "Profile Sucessfully Updated"
    End If
    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        TestDriveID = temprs.Collect(0)
    End If
    Set temprs = Nothing
    rsTestInv.Requery
    PicAdd.Visible = True
    PicSave.Visible = False
    picDetail.Enabled = False

End Sub

Private Sub cmdReturn_Click()
    If cmdReturn.Tag = "R" Then
        If TestDriveID > 0 Then
            gconDMIS.Execute "update cris_mrrinv set DateReturned=getdate() where id =" & TestDriveID
        End If
    Else
        If TestDriveID > 0 Then
            gconDMIS.Execute "update cris_mrrinv set DateReturned=NULL where id =" & TestDriveID
        End If
    End If
    rsTestInv.Requery
    StoreMemVars
End Sub

Private Sub Form_Load()
    InitVars
    Set rsTestInv = GetRS("SELECT * FROM CRIS_MRRINV")
    Set rsSearch = rsTestInv.Clone(adLockReadOnly)
    lvSearch.ColumnHeaders.Add , , "ID", 0
    lvSearch.ColumnHeaders.Add , , "Description", lvSearch.Width * 0.97
    StoreMemVars
End Sub

Sub InitVars()
    Dim cntl                                 As Control
    For Each cntl In Me.Controls
        If TypeOf cntl Is TextBox Or TypeOf cntl Is ComboBox Then
            cntl.Text = vbNullString
        End If
    Next
    txtTaggedPrice.Text = 0
    FillCombo "SELECT ID, Color_Desc FROM ALL_Color", 0, 1, cboColor

End Sub


Sub StoreMemVars()
    If Not (rsTestInv.EOF And rsTestInv.BOF) Then
        TestDriveID = rsTestInv!ID
        txtCode.Text = Null2String(rsTestInv!code)
        txtDescript.Text = Null2String(rsTestInv!descript)
        txtMake.Text = Null2String(rsTestInv!Make)
        txtModel.Text = Null2String(rsTestInv!Model)
        cboClass.Text = Null2String(rsTestInv!Class)
        txtYeer.Text = Null2String(rsTestInv!yearmodel)
        cboSource.Text = Null2String(rsTestInv!source)
        txtUnit.Text = Null2String(rsTestInv!unit)
        cboColor.Text = Null2String(rsTestInv!Color)
        txtSerialNo.Text = Null2String(rsTestInv!serialno)
        txtVINo.Text = Null2String(rsTestInv!vinnumber)
        txtEngineNo.Text = Null2String(rsTestInv!enginenumber)
        txtTaggedPrice.Text = Null2String(rsTestInv!taggedprice)
        dtdaterecieved.Value = Null2String(rsTestInv!DateReceived)
        txtNotes.Text = Null2String(rsTestInv!notes)

        If IsNull(rsTestInv!DateReturned) = True Then
            lblStatus.caption = " Available"
            cmdReturn.caption = "Return" & vbCrLf & "Vehicles"
            cmdReturn.Tag = "R"
        Else
            lblStatus.caption = " Returned "
            cmdReturn.caption = "Avail"
            cmdReturn.Tag = "A"
        End If
    Else
        cmdAdd.Value = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsTestInv = Nothing

End Sub

Private Sub lvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsTestInv.Filter = "ID= " & Item.Text
    'rsTestInv.Bookmark = rsFind(rsTestInv.Clone, "ID", Item.Text).Bookmark
    StoreMemVars
End Sub

Private Sub Timer1_Timer()
    Dim cntrl                                As Control
    For Each cntrl In Me.Controls
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            If cntrl.ForeColor = vbYellow Then
                cntrl.ForeColor = vbBlack
                cntrl.BackColor = vbWhite
            End If
        End If
    Next
    Timer1.Enabled = False
End Sub



Private Sub txtCode_Validate(Cancel As Boolean)
    txtCode.Text = UCase(txtCode.Text)
End Sub
Private Sub txtDescript_Validate(Cancel As Boolean)
    txtDescript.Text = UCase(txtDescript.Text)
End Sub

Private Sub txtEngineNo_Validate(Cancel As Boolean)
    txtEngineNo.Text = UCase(txtEngineNo.Text)
End Sub

Private Sub txtMake_Validate(Cancel As Boolean)
    txtMake.Text = UCase(txtMake.Text)
End Sub
Private Sub txtModel_Validate(Cancel As Boolean)
    txtModel.Text = UCase(txtModel.Text)
End Sub
Private Sub cboClass_Validate(Cancel As Boolean)
    cboClass.Text = UCase(cboClass.Text)
End Sub
Private Sub cboSource_Validate(Cancel As Boolean)
    cboSource.Text = UCase(cboSource.Text)
End Sub
Private Sub txtSerialNo_Validate(Cancel As Boolean)
    txtSerialNo.Text = UCase(txtSerialNo.Text)
End Sub


Private Sub txtSearch_Change()
    ''Dim temprs As ADODB.Recordset
    '    If txtSearch.Text = vbNullString Then: Exit Sub
    '    On Error GoTo adder:
    '    If optView(0).Value = True Then
    ''Set temprs = gconDMIS.Execute(" SELECT Descript from CRIS_MRRINV WHERE  CODE LIKE '" & txtSearch.Text & "%'")
    '    rsSearch.Filter = " CODE Like '" & txtSearch.Text & "%'"
    '    ElseIf optView(1).Value = True Then
    '    'Set temprs = gconDMIS.Execute(" SELECT Descript from CRIS_MRRINV WHERE  Model LIKE '" & txtSearch.Text & "%'")
    '        rsSearch.Filter = " Model Like '" & txtSearch.Text & "%'"
    '    ElseIf optView(2).Value = True Then
    ' '       Set temprs = gconDMIS.Execute("SELECT Descript from CRIS_MRRINV WHERE  Descript LIKE '" & txtSearch.Text & "%'")
    '        rsSearch.Filter = " Descript LIKE '" & txtSearch.Text & "%'"
    '    End If
    '    flex_FillListView rs, lvSearch, False, False
    '    Exit Sub
    'adder:
    '    Err.Clear
End Sub

Private Function Runvalidation(strcase As String) As Boolean
    Runvalidation = False
    Dim txt                                  As Control
    For Each txt In Me.Controls
        If (TypeOf txt Is TextBox Or TypeOf txt Is ComboBox) And txt.Tag = strcase Then
            If Trim(txt.Text) = vbNullString Then
                MessagePop RecSaveError, "Required Filed Missing", txt.ToolTipText & " is Required Field", 1000
                Call ColorIt(txt, Timer1)
                txt.SetFocus
                Exit Function
            End If
        End If
    Next
    Runvalidation = True
End Function


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Log_TestVehicles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Drive Vehicle Receiving/Returning  Entry"
   ClientHeight    =   6690
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "TestVehicles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   7695
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7695
      TabIndex        =   32
      Top             =   5775
      Width           =   7695
      Begin VB.Timer Timer1 
         Left            =   1050
         Top             =   360
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   5970
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   33
         Top             =   0
         Width           =   2580
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
            Left            =   755
            MouseIcon       =   "TestVehicles.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   45
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
            Left            =   60
            MouseIcon       =   "TestVehicles.frx":0D5A
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":0EAC
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   2190
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   36
         Top             =   0
         Width           =   5490
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
            Left            =   4530
            MouseIcon       =   "TestVehicles.frx":11FC
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":134E
            Style           =   1  'Graphical
            TabIndex        =   37
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
            Left            =   3840
            MouseIcon       =   "TestVehicles.frx":16B4
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":1806
            Style           =   1  'Graphical
            TabIndex        =   38
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
            Left            =   3150
            MouseIcon       =   "TestVehicles.frx":1B31
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":1C83
            Style           =   1  'Graphical
            TabIndex        =   39
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
            Left            =   2460
            MouseIcon       =   "TestVehicles.frx":1FDF
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":2131
            Style           =   1  'Graphical
            TabIndex        =   40
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
            Left            =   1770
            MouseIcon       =   "TestVehicles.frx":2444
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":2596
            Style           =   1  'Graphical
            TabIndex        =   41
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
            Left            =   1080
            MouseIcon       =   "TestVehicles.frx":2890
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":29E2
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   45
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
            Left            =   390
            MouseIcon       =   "TestVehicles.frx":2D3A
            MousePointer    =   99  'Custom
            Picture         =   "TestVehicles.frx":2E8C
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   270
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox picSearch 
      Align           =   3  'Align Left
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
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   0
      Width           =   2625
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
         Left            =   30
         TabIndex        =   4
         Text            =   "TEXT1"
         Top             =   870
         Width           =   2415
      End
      Begin VB.OptionButton optByCode 
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
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.OptionButton optByModel 
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
         Left            =   0
         TabIndex        =   2
         Top             =   270
         Width           =   2235
      End
      Begin VB.OptionButton optByDescription 
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
         Left            =   45
         TabIndex        =   3
         Top             =   570
         Width           =   2235
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   3975
         Left            =   0
         TabIndex        =   5
         Top             =   1380
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   7011
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
   End
   Begin VB.PictureBox picDataEntry 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   5775
      Left            =   2625
      ScaleHeight     =   5775
      ScaleWidth      =   5535
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtSource 
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
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "@R"
         Top             =   1335
         Width           =   3540
      End
      Begin VB.TextBox txtIGNKeyNo 
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
         Left            =   1305
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2610
         Width           =   3555
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   2655
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   10
         Tag             =   "@R"
         Text            =   "Text1"
         Top             =   480
         Width           =   2160
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   1245
      End
      Begin VB.ComboBox cboDescript 
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
         Left            =   1305
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   900
         Width           =   3555
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
         Left            =   1305
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   1755
         Width           =   3555
      End
      Begin MSComCtl2.DTPicker dtdaterecieved 
         Height          =   390
         Left            =   1305
         TabIndex        =   26
         Top             =   3885
         Width           =   3570
         _ExtentX        =   6297
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
         Format          =   52166657
         CurrentDate     =   39141
      End
      Begin VB.TextBox txtNotes 
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
         ForeColor       =   &H00701E2A&
         Height          =   885
         Left            =   1305
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   4770
         Width           =   3555
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
         Left            =   1305
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2175
         Width           =   3555
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
         Left            =   1305
         TabIndex        =   24
         Tag             =   "@R"
         Text            =   "Text1"
         Top             =   3450
         Width           =   3555
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
         Left            =   1305
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   3030
         Width           =   3555
      End
      Begin MSComCtl2.DTPicker dtDateReturned 
         Height          =   390
         Left            =   1320
         TabIndex        =   29
         Top             =   4290
         Width           =   3570
         _ExtentX        =   6297
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
         Format          =   52166657
         CurrentDate     =   39141
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Returned Date"
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
         Height          =   255
         Left            =   90
         TabIndex        =   31
         Top             =   4380
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   105
         TabIndex        =   30
         Top             =   90
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
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
         Height          =   225
         Left            =   90
         TabIndex        =   13
         Top             =   1335
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2700
         TabIndex        =   7
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   1785
         Width           =   885
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   975
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Arrived"
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
         Height          =   255
         Left            =   90
         TabIndex        =   25
         Top             =   3960
         Width           =   1065
      End
      Begin VB.Label Label42 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Top             =   4785
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "CS NO"
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
         Height          =   255
         Left            =   90
         TabIndex        =   20
         Top             =   2640
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   2205
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   23
         Top             =   3480
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   3045
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmCRIS_Log_TestVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TestDriveID                     As Long
Private rs                             As Recordset

Private Sub cboDescript_Change()
    If cboDescript.ListIndex = -1 Then Exit Sub
    Dim TEMPRS                         As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT * FROM ALL_MODEL WHERE ID=" & cboDescript.ItemData(cboDescript.ListIndex))
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        txtCode = Null2String(TEMPRS!CODE)
        txtModel = Null2String(TEMPRS!Model)
    End If
End Sub

Private Sub cboDescript_Click()
    cboDescript_Change
End Sub

Private Sub cmdAdd_Click()
    TestDriveID = 0
    InitVars
    picDataEntry.Enabled = True
    picSearch.Enabled = False
    picAdds.Visible = False
    picSaves.Visible = True
    dtDateReturned.Enabled = False
    On Error Resume Next
    cboDescript.SetFocus
End Sub

Private Sub cmdCancel_Click()
    picDataEntry.Enabled = False
    picSearch.Enabled = True
    picAdds.Visible = True
    picSaves.Visible = False
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If gconDMIS.Execute("SELECT Count(*) From CRIS_TestDriveSchedules Where VehicleCode=" & N2Str2Null(txtCode)).Fields(0).Value > 0 Then
        MessagePop RecLocekd, "Record In Use", "Current Test Drive Information Has Been In Use .... Cannot Delete The Record"

        Exit Sub
    End If

    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_MRRINV where id = " & TestDriveID
        ShowDeletedMsg
        FillSearchGrid txtSearch
        rsRefresh
        StoreMemvars
    End If
End Sub

Private Sub cmdEdit_Click()
    picDataEntry.Enabled = True
    picSearch.Enabled = False
    picAdds.Visible = False
    picSaves.Visible = True
    dtDateReturned.Enabled = True
    On Error Resume Next
    cboDescript.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rs.MoveNext
    If rs.EOF Then
        rs.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars

End Sub

Private Sub cmdPrevious_Click()
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub




Private Sub cmdSave_Click()


    If txtCode = "" Then
        ShowIsRequiredMsg "Vehicle Model Code"
        On Error Resume Next
        txtCode.SetFocus
        Exit Sub
    End If

    If cboDescript.ListIndex = -1 Then
        ShowIsRequiredMsg "Vehicle Model "
        On Error Resume Next
        cboDescript.SetFocus
        Exit Sub
    End If


    If cboColor = "" Then
        ShowIsRequiredMsg "Model Color"
        On Error Resume Next
        cboColor.SetFocus
        Exit Sub
    End If

    If txtVINo = "" Or txtEngineNo = "" Then
        ShowIsRequiredMsg " Vin Number and Engine Number"
        On Error Resume Next
        txtVINo.SetFocus
        Exit Sub
    End If


    If txtIGNKeyNo = "" Then
        ShowIsRequiredMsg "Conduction Sticker Number"
        On Error Resume Next
        txtIGNKeyNo.SetFocus
        Exit Sub
    End If




    Dim CODE                           As String
    Dim DESCRIPT                       As String
    Dim Source                         As String
    Dim Color                          As String
    Dim ENGINENUMBER                   As String
    Dim SERIALNO                       As String
    Dim VINNUMBER                      As String
    Dim TAGGEDPRICE                    As String
    Dim Notes                          As String
    Dim DATERECIEVED                   As String
    Dim Model                          As String
    Dim IGNKEYNO                       As String


    CODE = N2Str2Null(txtCode)
    DESCRIPT = N2Str2Null(cboDescript)
    Color = N2Str2Null(cboColor)
    ENGINENUMBER = N2Str2Null(txtEngineNo)
    SERIALNO = N2Str2Null(txtSerialNo)
    VINNUMBER = N2Str2Null(txtVINo)
    Notes = N2Str2Null(txtNotes)
    DATERECIEVED = N2Date2Null(dtdaterecieved.Value)
    Source = N2Str2Null(txtSource)
    Model = N2Str2Null(txtModel)
    IGNKEYNO = N2Str2Null(txtIGNKeyNo)


    Dim TEMPRS                         As ADODB.Recordset
    If TestDriveID = 0 Then
        SQL = " INSERT INTO  " & _
            "  CRIS_MRRINV(Code, Descript, Model, Source, Color, EngineNumber, SerialNo, VinNumber, IGNKEYNO, Notes, DateReceived, HitCounter ) " & _
            "  VALUES( " & CODE & " , " & DESCRIPT & " , " & Model & " , " & Source & "," & Color & _
            " , " & ENGINENUMBER & " , " & SERIALNO & " ," & VINNUMBER & "," & IGNKEYNO & " ," & Notes & " ," & DateReceived & " , 0) " & vbCrLf & " SELECT @@IDENTITY "
    Else

        SQL = " UPDATE CRIS_MRRINV " & _
            " SET Code = " & CODE & " ," & _
            " Descript= " & DESCRIPT & "," & _
            " Model= " & Model & " ," & _
            " Source= " & Source & " ," & _
            " Color=" & Color & " ," & _
            " EngineNumber= " & ENGINENUMBER & " ," & _
            " SerialNo= " & SERIALNO & " ," & _
            " VinNumber= " & VINNUMBER & "," & _
            " IGNKEYNO=  " & IGNKEYNO & "," & _
            " Notes= " & Notes & " , " & _
            " DateReceived = " & DATERECIEVED & _
            " WHERE id= " & TestDriveID

        If IsNull(dtDateReturned.Value) = False Then
            If MsgBox("Do you Want to Return This Vehicle. ", vbOKCancel + vbExclamation, "Confirm Posting") = vbOK Then
                gconDMIS.Execute "update cris_mrrinv set DateReturned= " & N2Date2Null(dtDateReturned.Value) & " where id =" & TestDriveID
                MessagePop RecSaveOk, "Returned", "Test Drive Vehicle Returned"
            End If
        Else
            gconDMIS.Execute "update cris_mrrinv set DateReturned= NULL where id =" & TestDriveID
        End If
    End If


    Set TEMPRS = gconDMIS.Execute(SQL)
    If TestDriveID <= 0 Then
        MessagePop RecSave, "Record Added ", "New Test Vehicles Sucessfully "
    Else
        MessagePop RecSave, "Record Saved", "Test Vehicles  Sucessfully Updated"
    End If
    Set TEMPRS = TEMPRS.NextRecordset
    If Not TEMPRS Is Nothing Then
        TestDriveID = TEMPRS.Collect(0)
    End If
    Set TEMPRS = Nothing

    rs.Requery
    rs.Find ("ID=" & TestDriveID)
    cmdCancel.Value = True
    FillSearchGrid txtSearch

End Sub

Sub FillSearchGrid(xxx As String)
    Dim TEMPRS                         As ADODB.Recordset
    lvSearch.Sorted = False: lvSearch.ListItems.Clear
    Set TEMPRS = New ADODB.Recordset

    If optByCode.Value = True Then
        Set TEMPRS = gconDMIS.Execute("select  Code, Descript, ID from CRIS_MRRINV where CODE like'" & ReplaceQuote(xxx) & "%' order by 1 asc")
    ElseIf optByModel.Value = True Then
        Set TEMPRS = gconDMIS.Execute("select  Code, Descript, ID from CRIS_MRRINV where Model like'" & ReplaceQuote(xxx) & "%' order by 1 asc")
    ElseIf optByDescription.Value = True Then
        Set TEMPRS = gconDMIS.Execute("select  Code, Descript, ID from CRIS_MRRINV where Descript like'" & ReplaceQuote(xxx) & "%' order by 1 asc")

    End If

    If Not (TEMPRS.EOF And TEMPRS.BOF) Then
        Listview_Loadval lvSearch.ListItems, TEMPRS
        lvSearch.Refresh
    End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Call AddColumnHeader("Model,Description", lvSearch)
    Call ResizeColumnHeader(lvSearch, "25,70")
    InitVars
    rsRefresh

    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False

    StoreMemvars
    FillSearchGrid txtSearch

End Sub

Sub InitVars()
    Dim cntl                           As Control
    For Each cntl In Me.ControlS
        If TypeOf cntl Is TextBox Or TypeOf cntl Is ComboBox Then
            cntl = vbNullString
        End If
    Next
    dtdaterecieved.Value = Now
    dtDateReturned.Value = Null
    lblStatus = ""
    FillCombo "SELECT ID, Color_Desc FROM ALL_Color", 0, 1, cboColor
    FillCombo "Select ID,  Descript from All_Model where LEN(code)<> 0 order by descript asc", 0, 1, cboDescript
End Sub

Private Sub lvSearch_DblClick()
    If lvSearch.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rs.MoveFirst
    rs.Find ("ID= " & Item.ListSubItems(lvSearch.ColumnHeaders.Count))
    StoreMemvars
End Sub

Private Sub optByCode_Click()
    FillSearchGrid txtSearch
End Sub

Private Sub optByDescription_Click()
    FillSearchGrid txtSearch
End Sub

Private Sub optByModel_Click()
    FillSearchGrid txtSearch
End Sub

Sub rsRefresh()

    Set rs = gconDMIS.Execute("SELECT * FROM CRIS_MRRINV order by id desc")

End Sub

Sub StoreMemvars()
    If Not (rs.EOF And rs.BOF) Then
        TestDriveID = rs!id
        txtCode = Null2String(rs!CODE)
        cboDescript.ListIndex = SelectCombo(cboDescript, Null2String(rs!DESCRIPT))
        txtModel = Null2String(rs!Model)
        cboColor = Null2String(rs!Color)
        txtSerialNo = Null2String(rs!SERIALNO)
        txtVINo = Null2String(rs!VINNUMBER)
        txtEngineNo = Null2String(rs!ENGINENUMBER)
        txtIGNKeyNo = Null2String(rs!IGNKEYNO)
        dtdaterecieved.Value = Null2String(rs!DateReceived)

        dtDateReturned.Value = Null2String(rs!DateReturned)

        txtNotes = Null2String(rs!Notes)
        txtSource = Null2String(rs!Source)
        If Null2String(rs!HitCounter) <> "" Then
            cmdDelete.Enabled = False
        Else
            cmdDelete.Enabled = True
        End If
        If IsNull(rs!DateReturned) = True Then
            lblStatus.Caption = "***Available***"
        Else
            lblStatus.Caption = "***Returned***"
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If

End Sub

Private Sub Timer1_Timer()
    If lblStatus.Caption <> "" Then
        If lblStatus.Visible = True Then
            lblStatus.Visible = False
        Else
            lblStatus.Visible = True
        End If
    End If
End Sub

Private Sub txtEngineNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)

End Sub

Private Sub txtIGNKeyNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid txtSearch
End Sub

Private Sub txtSerialNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSource_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtVINo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub


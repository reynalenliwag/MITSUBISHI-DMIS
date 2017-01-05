VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmSMIS_Trans_VehiclesCheckList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VEHICLE CHECK LIST FORM"
   ClientHeight    =   8595
   ClientLeft      =   1125
   ClientTop       =   1200
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "VehiclesCheckList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   9645
   Begin VB.PictureBox Picture1 
      Height          =   1875
      Left            =   7200
      ScaleHeight     =   1815
      ScaleWidth      =   1575
      TabIndex        =   76
      Top             =   4950
      Visible         =   0   'False
      Width           =   1635
      Begin VB.TextBox txtPDI_NO 
         Height          =   405
         Left            =   420
         TabIndex        =   78
         Top             =   1170
         Width           =   615
      End
      Begin VB.TextBox txtVI_NO 
         Height          =   345
         Left            =   390
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label labid 
         Caption         =   "0"
         Height          =   315
         Left            =   390
         TabIndex        =   79
         Top             =   270
         Width           =   405
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   2925
      Left            =   60
      TabIndex        =   51
      Top             =   4410
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   5159
      Cols            =   6
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      GridColor       =   12632256
      Rows            =   2
   End
   Begin VB.PictureBox picBottoms 
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   60
      ScaleHeight     =   1185
      ScaleWidth      =   9540
      TabIndex        =   52
      Top             =   7380
      Width           =   9540
      Begin VB.Timer tmBlink 
         Interval        =   500
         Left            =   2640
         Top             =   120
      End
      Begin Crystal.CrystalReport rptVehicleCheckList 
         Left            =   2100
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Units Released"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   1980
         ScaleHeight     =   915
         ScaleWidth      =   11580
         TabIndex        =   56
         Top             =   330
         Width           =   11580
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   6780
            MouseIcon       =   "VehiclesCheckList.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Exit Window"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   6090
            MouseIcon       =   "VehiclesCheckList.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Print this Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdCancelCO 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   5400
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VehiclesCheckList.frx":123A
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":138C
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Cancel this Transaction"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdUnPost 
            Caption         =   "Unpost"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   4710
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VehiclesCheckList.frx":16C6
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":1818
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Unpost this Transaction"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "Post"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   4020
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VehiclesCheckList.frx":1B5D
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":1CAF
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Post this Transaction"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   3330
            MouseIcon       =   "VehiclesCheckList.frx":1FD4
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":2126
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Add Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   2640
            MouseIcon       =   "VehiclesCheckList.frx":2439
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":258B
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Edit Selected Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Fin&d"
            Height          =   795
            Left            =   1950
            MouseIcon       =   "VehiclesCheckList.frx":28E7
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":2A39
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Find a Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1260
            MouseIcon       =   "VehiclesCheckList.frx":2D33
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":2E85
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Move to Next Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   570
            MouseIcon       =   "VehiclesCheckList.frx":31DD
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":332F
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Move to Previous Record"
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
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
         Left            =   7950
         ScaleHeight     =   885
         ScaleWidth      =   1560
         TabIndex        =   53
         Top             =   330
         Width           =   1560
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   780
            MouseIcon       =   "VehiclesCheckList.frx":368E
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":37E0
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Cancel"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   90
            MouseIcon       =   "VehiclesCheckList.frx":3B1E
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesCheckList.frx":3C70
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Save this Record"
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.Label LABSATUS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   270
         TabIndex        =   73
         Top             =   630
         Width           =   1965
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4170
         TabIndex        =   72
         Top             =   30
         Width           =   1185
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "POOR:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5940
         TabIndex        =   71
         Top             =   30
         Width           =   1185
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOT CHECKED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7710
         TabIndex        =   70
         Top             =   30
         Width           =   1185
      End
      Begin VB.Label labOK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5370
         TabIndex        =   69
         Top             =   30
         Width           =   555
      End
      Begin VB.Label labPoor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7140
         TabIndex        =   68
         Top             =   30
         Width           =   555
      End
      Begin VB.Label labNotChecked 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8910
         TabIndex        =   67
         Top             =   30
         Width           =   555
      End
   End
   Begin VB.ComboBox cboPDI_Category 
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
      Left            =   7020
      TabIndex        =   49
      Top             =   3975
      Width           =   2505
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2355
      TabIndex        =   50
      Top             =   4005
      Width           =   1875
   End
   Begin VB.PictureBox picMaster 
      BorderStyle     =   0  'None
      Height          =   7125
      Left            =   -60
      ScaleHeight     =   7125
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H00701E2A&
         Height          =   840
         Left            =   360
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Tag             =   "@R"
         ToolTipText     =   "Customer Name "
         Top             =   2985
         Width           =   8880
      End
      Begin VB.TextBox txtPDI_ODOMETER 
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
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   5730
         MaxLength       =   5
         TabIndex        =   26
         ToolTipText     =   "Customer Name "
         Top             =   1980
         Width           =   825
      End
      Begin VB.ComboBox cboPDI_Transmission 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "VehiclesCheckList.frx":3FC0
         Left            =   6600
         List            =   "VehiclesCheckList.frx":3FC2
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   1980
         Width           =   1425
      End
      Begin VB.ComboBox cboPDI_FuelGauge 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "VehiclesCheckList.frx":3FC4
         Left            =   8100
         List            =   "VehiclesCheckList.frx":3FC6
         TabIndex        =   25
         Top             =   1950
         Width           =   1005
      End
      Begin VB.ComboBox cboPDI_SAE 
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
         Left            =   6900
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   570
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker txtPDI_Deyt 
         Height          =   375
         Left            =   6900
         TabIndex        =   7
         Top             =   150
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   51183619
         CurrentDate     =   38941
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   120
         TabIndex        =   1
         Top             =   -60
         Width           =   9375
         Begin VB.CommandButton Command1 
            Caption         =   "::"
            CausesValidation=   0   'False
            Height          =   345
            Left            =   8940
            TabIndex        =   5
            Top             =   240
            Width           =   345
         End
         Begin VB.TextBox txtPDI_CusName 
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
            ForeColor       =   &H00701E2A&
            Height          =   390
            Left            =   240
            TabIndex        =   3
            Tag             =   "@R"
            ToolTipText     =   "Customer Name "
            Top             =   510
            Width           =   5025
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SAE/CheckLister"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   5325
            TabIndex        =   6
            Top             =   705
            Width           =   1425
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   6330
            TabIndex        =   4
            Top             =   300
            Width           =   390
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   11
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   3030
         Left            =   120
         TabIndex        =   9
         Top             =   900
         Width           =   9375
         Begin VB.TextBox txtPDI_PlateNo 
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
            ForeColor       =   &H00701E2A&
            Height          =   360
            Left            =   3390
            MaxLength       =   6
            TabIndex        =   75
            Top             =   480
            Width           =   1725
         End
         Begin VB.TextBox txtModelDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   420
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   23
            Tag             =   "@R"
            ToolTipText     =   "Customer Name "
            Top             =   1440
            Width           =   8835
         End
         Begin VB.TextBox txtPDI_VINNO 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   10
            Tag             =   "@R"
            ToolTipText     =   "Customer Name "
            Top             =   480
            Width           =   3075
         End
         Begin VB.TextBox txtPDI_EngineNo 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   5610
            Locked          =   -1  'True
            TabIndex        =   13
            Tag             =   "@R"
            ToolTipText     =   "Customer Name "
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txtPDI_Make 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   20
            Tag             =   "@R"
            ToolTipText     =   "Customer Name "
            Top             =   1080
            Width           =   1065
         End
         Begin VB.TextBox txtPDI_Model 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   21
            Tag             =   "@R"
            ToolTipText     =   "Customer Name "
            Top             =   1080
            Width           =   2025
         End
         Begin VB.TextBox txtPDI_Color 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   22
            Tag             =   "@R"
            ToolTipText     =   "Customer Name "
            Top             =   1080
            Width           =   2115
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CS#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   3390
            TabIndex        =   74
            Top             =   210
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   24
            Top             =   1860
            Width           =   780
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tranmission"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   6480
            TabIndex        =   17
            Top             =   825
            Width           =   1065
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Odometer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   0
            Left            =   5550
            TabIndex        =   16
            Top             =   825
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuel Guage"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   3
            Left            =   7950
            TabIndex        =   18
            Top             =   825
            Width           =   930
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   2
            Left            =   1320
            TabIndex        =   15
            Top             =   825
            Width           =   510
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VIN Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   210
            Width           =   1035
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Engine Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   0
            Left            =   5610
            TabIndex        =   11
            Top             =   210
            Width           =   1290
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   17
            Left            =   3390
            TabIndex        =   19
            Top             =   825
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Make"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   825
            Width           =   465
         End
      End
      Begin VB.Label labModelCode 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4275
         TabIndex        =   30
         Top             =   4020
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label labPDI_ID 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4620
         TabIndex        =   31
         Top             =   4020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Filter View By Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   5010
         TabIndex        =   32
         Top             =   4035
         Width           =   1950
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Filter View By Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   555
         TabIndex        =   29
         Top             =   4020
         Width           =   1740
      End
   End
   Begin VB.PictureBox picAddVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   390
      ScaleHeight     =   5625
      ScaleWidth      =   9015
      TabIndex        =   41
      Top             =   1290
      Visible         =   0   'False
      Width           =   9045
      Begin XtremeReportControl.ReportControl lvAddVehicles 
         Height          =   4140
         Left            =   45
         TabIndex        =   45
         Top             =   720
         Width           =   8910
         _Version        =   655364
         _ExtentX        =   15716
         _ExtentY        =   7302
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         SkipGroupsFocus =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   8640
         TabIndex        =   48
         Top             =   0
         Width           =   330
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1755
         TabIndex        =   43
         Top             =   330
         Width           =   3915
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "&Cancel"
         Height          =   660
         Index           =   0
         Left            =   8280
         MouseIcon       =   "VehiclesCheckList.frx":3FC8
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesCheckList.frx":411A
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   4890
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Select"
         Height          =   660
         Left            =   7620
         MouseIcon       =   "VehiclesCheckList.frx":4458
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesCheckList.frx":45AA
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4890
         Width           =   675
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search Keyword"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   44
         Top             =   360
         Width           =   2505
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   -15
         TabIndex        =   42
         Top             =   0
         Width           =   9045
         _Version        =   655364
         _ExtentX        =   15954
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "ADD VEHICLE FOR PDI CHECK LIST"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   -2147483630
      End
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   390
      ScaleHeight     =   5625
      ScaleWidth      =   9015
      TabIndex        =   33
      Top             =   1290
      Visible         =   0   'False
      Width           =   9045
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   4140
         Left            =   45
         TabIndex        =   38
         Top             =   720
         Width           =   8910
         _Version        =   655364
         _ExtentX        =   15716
         _ExtentY        =   7302
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         SkipGroupsFocus =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "&Cancel"
         Height          =   660
         Index           =   2
         Left            =   8310
         MouseIcon       =   "VehiclesCheckList.frx":48E6
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesCheckList.frx":4A38
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4920
         Width           =   645
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   660
         Left            =   7620
         MouseIcon       =   "VehiclesCheckList.frx":4D76
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesCheckList.frx":4EC8
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4920
         Width           =   645
      End
      Begin VB.TextBox txtFilterViewVehicles 
         Height          =   375
         Left            =   1755
         TabIndex        =   36
         Top             =   330
         Width           =   3915
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   8640
         TabIndex        =   35
         Top             =   0
         Width           =   330
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   315
         Left            =   -15
         TabIndex        =   34
         Top             =   0
         Width           =   9045
         _Version        =   655364
         _ExtentX        =   15954
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "VEHICLE CHECK LIST"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search Keyword"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   10
         Left            =   90
         TabIndex        =   37
         Top             =   360
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_VehiclesCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsVCHK                                                            As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim ISEDITED                                                          As Boolean

Function GetFuelGauage(XXX)
    If XXX = "F" Then
        GetFuelGauage = "FULL"
    ElseIf XXX = "H" Then
        GetFuelGauage = "HALF"
    ElseIf XXX = "E" Then
        GetFuelGauage = "EMPTY"
    End If
End Function

Function GetStatus(XXX)
    If XXX = "N" Then
        GetStatus = "NOT CHECKED"
    ElseIf XXX = "O" Then
        GetStatus = "OK"
    ElseIf XXX = "P" Then
        GetStatus = "POOR"
    Else
        GetStatus = "NOT CHECKED"
    End If
End Function

Function SetFuelGauage(XXX)
    If XXX = "FULL" Then
        SetFuelGauage = "F"
    ElseIf XXX = "HALF" Then
        SetFuelGauage = "H"
    ElseIf XXX = "EMPTY" Then
        SetFuelGauage = "E"
    End If
End Function

Function SetStatus(XXX)
    If XXX = "NOT CHECKED" Then
        SetStatus = "N"
    ElseIf XXX = "OK" Then
        SetStatus = "O"
    ElseIf XXX = "POOR" Then
        SetStatus = "P"
    End If
End Function

Private Function GetCategory(XXX As String) As String
    XXX = UCase(RTrim(LTrim(XXX)))
    Select Case XXX
        Case "VE"
            GetCategory = "VEHICLE EXTERIOR"
        Case "VI"
            GetCategory = "VEHICLE INTERIOR"
        Case "EC"
            GetCategory = "ENGINE COMPARTMENT"
        Case "EE"
            GetCategory = "ELECTRICAL"
        Case "TO"
            GetCategory = "TOOLS"
    End Select
End Function

Private Function SETCATEGORY(ModelCode As String) As String
    'UDPATING CODE    :AXP-672007312
    ModelCode = UCase(RTrim(LTrim(ModelCode)))
    Select Case UCase(ModelCode)
        Case "VEHICLE EXTERIOR"
            SETCATEGORY = "VE"
        Case "VEHICLE INTERIOR"
            SETCATEGORY = "VI"
        Case "ENGINE COMPARTMENT"
            SETCATEGORY = "EC"
        Case "ELECTRICAL"
            SETCATEGORY = "EE"
        Case "TOOLS"
            SETCATEGORY = "TO"
    End Select
End Function

Sub SearchByInvoice(vVI_NO As String)
    RsVCHK.Requery
    RsVCHK.Find ("VI_NO='" & vVI_NO & "'")
    StoreMemVars
End Sub

Sub FillGrid(XXXCategory As String, XXXStatus As String)
    Grid1.Visible = False
    Dim SQL                                                           As String
    Dim temprs                                                        As ADODB.Recordset
    Dim STATUS                                                        As String
    Dim CSTATUS                                                       As String
    Dim countOK                                                       As Long
    Dim countPoor                                                     As Long
    Dim countNC                                                       As Long
    SQL = "SELECT   LOOKUP.PDINAME, LOOKUP.PDICATEGORY,DET.QTY, DET.STATUS, "
    SQL = SQL & " DET.PDILINENO FROM  SMIS_PDI_DET DET INNER JOIN "
    SQL = SQL & " SMIS_vw_PDILookUp LOOKUP ON DET.PDILINENO = LOOKUP.PDILINEID WHERE VI_NO=" & N2Str2Null(txtVI_NO)

    XXXCategory = LTrim(RTrim(XXXCategory))
    XXXStatus = LTrim(RTrim(XXXStatus))
    If XXXCategory = "" And XXXStatus = "" Then

    ElseIf XXXCategory <> "" And XXXStatus = "" Then
        SQL = SQL & "  AND PDICATEGORY='" & ReplaceQuote(XXXCategory) & "'"
    ElseIf XXXCategory = "" And XXXStatus <> "" Then
        SQL = SQL & " AND  STATUS='" & ReplaceQuote(XXXStatus) & "'"
    ElseIf XXXCategory <> "" And XXXStatus <> "" Then
        SQL = SQL & " AND PDICATEGORY='" & ReplaceQuote(XXXCategory) & "' AND STATUS='" & ReplaceQuote(XXXStatus) & "'"
    End If

    Set temprs = gconDMIS.Execute(SQL)

    Grid1.Rows = 1

    While Not temprs.EOF
        CSTATUS = Null2String(temprs!STATUS)
        STATUS = GetStatus(CSTATUS)
        If CSTATUS = "N" Then
            countNC = countNC + 1
        ElseIf CSTATUS = "O" Then
            countOK = countOK + 1
        ElseIf CSTATUS = "P" Then
            countPoor = countPoor + 1
        Else
            countNC = countNC + 1
        End If
        Grid1.AddItem _
                Null2String(temprs!PDINAME) & Chr(9) & _
                                            Null2String(temprs!PDICATEGORY) & Chr(9) & _
                                            Null2String(temprs!QTY) & Chr(9) & _
                                            STATUS & Chr(9) & _
                                            temprs!PDILINENO, False
        labOK = countOK
        labPoor = countPoor
        labNotChecked = countNC
        temprs.MoveNext
    Wend
    Grid1.Visible = True
    Grid1.Refresh
End Sub

Sub InitData()
    ReportControlAddColumnHeader lvViewVehicles, "SN,DATE,INV#,CUSTOMER NAME, PLATE#, SAE,COLOR"
    ReportControlPaintManager lvViewVehicles
    ResizeColumnHeader lvViewVehicles, ".25,.6,.7,1.5,.7,1.5,1.3"


    ReportControlAddColumnHeader lvAddVehicles, "SN,CS#, MODEL CODE, MODEL,DESCRIPT"
    ReportControlPaintManager lvAddVehicles
    ResizeColumnHeader lvAddVehicles, ".25,.6,1.6,.5,3"



    Call FillCombo("SELECT NAME  From SMIS_vw_Srep", -1, 0, cboPDI_SAE)
    With Grid1
        .Column(0).Width = 40
        .Column(1).Width = 350
        .Column(2).Width = 50
        .Column(3).Width = 50
        .Column(4).Width = 100
        .Column(5).Width = 0
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(0).Locked = True
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True

        .Column(3).Mask = cellNumeric
        .Column(3).MaxLength = 3
        .Column(3).CellType = cellTextBox
        .Column(4).CellType = cellComboBox
        .ComboBox(4).AddItem "OK"
        .ComboBox(4).AddItem "NOT CHECKED"
        .ComboBox(4).AddItem "POOR"
        .ComboBox(4).Font.Name = "ARIAL"
    End With
    With cboPDI_Category
        .AddItem "Vehicle Exterior"
        .AddItem "Vehicle Interior"
        .AddItem "Engine Compartment"
        .AddItem "Electrical"
        .AddItem "Tools"
        .AddItem "All"
    End With
    With cboPDI_FuelGauge
        .AddItem "FULL"
        .AddItem "HALF"
        .AddItem "EMPTY"
    End With

    With cboPDI_Transmission
        .AddItem "MT"
        .AddItem "AT"
    End With

    With Combo1
        .AddItem "OK"
        .AddItem "NOT CHECKED"
        .AddItem "POOR"
        .AddItem "ALL"
    End With

End Sub

Sub initMemvars()
    Grid1.Rows = 1
    txtPDI_Color = ""
    txtPDI_CusName = ""
    txtPDI_Deyt.Value = DateValue(LOGDATE)
    txtPDI_EngineNo = ""
    txtPDI_ODOMETER = ""
    txtPDI_PlateNo = ""
    txtPDI_VINNO = ""
    cboPDI_FuelGauge = ""
    cboPDI_SAE = ""
    cboPDI_Transmission = ""
    txtPDI_Make = ""
    txtPDI_Model = ""

End Sub

Sub rsRefresh()
    Set RsVCHK = New ADODB.Recordset
    RsVCHK.Open "select * from SMIS_PDI_HDR order by id desc", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not (RsVCHK.EOF Or RsVCHK.BOF) Then
        Dim TStatus                                                   As String
        TStatus = Null2String(RsVCHK!STATUS)
        labid = Null2String(RsVCHK!ID)
        '
        If Null2String(RsVCHK!STATUS) = "P" Then
            cmdPrint.Enabled = True
            LABSATUS = "**POSTED**"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = True
            cmdCancelCO.Enabled = False
        ElseIf Null2String(RsVCHK!STATUS) = "" Then
            cmdPrint.Enabled = False
            cmdEdit.Enabled = True
            LABSATUS = ""
            cmdPost.Enabled = True
            cmdUnPost.Enabled = False
            cmdCancelCO.Enabled = True

        ElseIf Null2String(RsVCHK!STATUS) = "C" Then
            cmdPrint.Enabled = False
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            LABSATUS = "**CANCELLED**"
            cmdUnPost.Enabled = False
            cmdCancelCO.Enabled = False

        End If


        txtVI_NO = Null2String(RsVCHK!VI_NO)
        txtPDI_NO = Null2String(RsVCHK!VI_NO)
        txtPDI_CusName = Null2String(RsVCHK!CustName)
        txtPDI_PlateNo = Null2String(RsVCHK!PlateNo)
        txtPDI_Color = Null2String(RsVCHK!Color)
        txtPDI_VINNO = Null2String(RsVCHK!VINO)
        cboPDI_SAE = Null2String(RsVCHK!SAE)
        txtPDI_EngineNo = Null2String(RsVCHK!EngineNo)
        cboPDI_Transmission = Null2String(RsVCHK!Tranmission)
        txtPDI_Make = Null2String(RsVCHK!Make)
        txtPDI_Model = Null2String(RsVCHK!Model)
        labPDI_ID = Null2String(RsVCHK!ID)
        txtPDI_Deyt.Value = Null2String(RsVCHK!PDIDate)
        Text1 = Null2String(RsVCHK!Notes)
        txtPDI_ODOMETER = NumericVal(RsVCHK!Odometer)
        labModelCode = Null2String(RsVCHK!ModelCode)
        cboPDI_FuelGauge = GetFuelGauage(Null2String(RsVCHK!Fuel))

        FillGrid SETCATEGORY(cboPDI_Category), SetStatus(Combo1)

    End If
End Sub

Sub UpdateDetail()

    ISEDITED = False
End Sub

Sub UpdateHeader()
    'UDPATING CODE      :   AXP-065082007 328PM
    Dim vtxtDate                                                      As String
    Dim vtxtTransmission                                              As String
    Dim vtxtSalesAE                                                   As String
    Dim vtxtOdometer                                                  As String
    Dim vtxtFuelGauge                                                 As String
    Dim SQL                                                           As String
    vtxtDate = N2Str2Null(DateValue(txtPDI_Deyt.Value))
    vtxtTransmission = N2Str2Null(cboPDI_Transmission)
    vtxtSalesAE = N2Str2Null(cboPDI_SAE)
    vtxtOdometer = NumericVal(txtPDI_ODOMETER)
    vtxtFuelGauge = N2Str2Null(SetFuelGauage(cboPDI_FuelGauge))
    SQL = " Update DMIS.dbo.SMIS_PDI_HDR SET "
    SQL = SQL & "PDIDATE=" & vtxtDate & " ,"
    SQL = SQL & "Fuel=" & vtxtFuelGauge & " ,"
    SQL = SQL & "Odometer=" & vtxtOdometer & " ,"
    SQL = SQL & "Tranmission=" & vtxtTransmission & " ,"
    SQL = SQL & "Notes=" & N2Str2Null(Text1) & " ,"
    SQL = SQL & "SAE=" & vtxtSalesAE
    SQL = SQL & " WHERE ID= " & labPDI_ID
    gconDMIS.Execute (SQL)



End Sub

Private Sub cboPDI_Category_Change()


    If AddorEdit = "EDIT" Then
        If ISEDITED = True Then
            If MsgBox("Do You Want To Save Changes.", vbYesNo Or vbQuestion Or vbDefaultButton1, App.TITLE) = vbYes Then
                UpdateDetail
                FillGrid SETCATEGORY(cboPDI_Category), SetStatus(Combo1)
                Exit Sub
            End If
        End If
    End If
    FillGrid SETCATEGORY(cboPDI_Category), SetStatus(Combo1)

End Sub

Private Sub cboPDI_Category_Click()
    cboPDI_Category_Change
End Sub

Private Sub cmdAdd_Click()
    initMemvars
    On Error GoTo Errorcode:
    txtPDI_NO = GenerateCode("SMIS_PDI_HDR", "VI_No", "000000")
    AddorEdit = "ADD"
    Screen.MousePointer = 11
    Dim SQL                                                           As String
    SQL = "SELECT IGNKEY ,MODEL,MODELCODE,DESCRIPT ,ID FROM SMIS_MRRINV_TABLE"
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute(SQL)
    flex_FillReportView temprs, lvAddVehicles, True
    ShowHidePictureBox2 picAddVehicles, True, picMaster

    picAdds.Visible = False: picSaves.Visible = True

    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    '
    '    Screen.MousePointer = 11
    '    Dim SQL                             As String
    '    SQL = "SELECT  "
    '    SQL = SQL & " SO.VI_NO,SO.CUSTNAME,SO.PLATE_NO,SO.COLOR,MRR.VINO,SO.SALESAE , "
    '    SQL = SQL & " MRR.ENGINENO, MRR.TRANSMISSION, UPPER(MRR.MAKE), SO.MODEL, SO.ID ,MRR.MODELCODE  "
    '    SQL = SQL & "FROM SMIS_SALESORDER SO "
    '    SQL = SQL & "INNER JOIN SMIS_MRRINV MRR ON MRR.PRODNO=SO.PRODNO "
    '    SQL = SQL & "WHERE SO.VI_NO NOT IN(SELECT SO.VI_NO FROM SMIS_PDI_HDR) AND  "
    '    SQL = SQL & "(SO.STATUS<>'C' OR SO.SOSTATUS<>'C') "
    '    Dim TEMPRS                          As ADODB.Recordset
    '    Set TEMPRS = gconDMIS.Execute(SQL)
    '    flex_FillReportView TEMPRS, lvViewVehicles, True
    '    ShowHidePictureBox2 picViewVehicles, True, picMaster
    '    Screen.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picMaster.Enabled = False
    picSaves.Visible = False
    picAdds.Visible = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    StoreMemVars
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "ACESS_CANCELENTRY", "PDI CHECKLIST") = False Then Exit Sub

    Dim SQL                                                           As String
    If MsgBox("Are You Sure You Want To Cancel This Transaction", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    SQL_STATEMENT = "update smis_pdi_hdr set status='C' where id = " & labid
    gconDMIS.Execute SQL_STATEMENT
    Call NEW_LogAudit("C", "PDI CHECKLIST", SQL_STATEMENT, labid, "", "", "", "")
    
    MessagePop RecSaveOk, "Transaction Cancelled", "Record Sucessfully Cancelled", 1000, 2
    
    
    rsRefresh
    RsVCHK.Find ("ID=" & labid)
    StoreMemVars
End Sub

Private Sub cmdCancelViewVehicles_Click(Index As Integer)
    ShowHidePictureBox2 picAddVehicles, False, picMaster
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "PDI CHECKLIST") = False Then Exit Sub
    On Error GoTo Errorcode:
    AddorEdit = "EDIT"
    picMaster.Enabled = True
    Grid1.Column(3).Locked = False
    Grid1.Column(4).Locked = False
    picSaves.Visible = True: picAdds.Visible = False
    txtPDI_Deyt.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    On Error GoTo Errorcode:
    Unload Me
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdFind_Click()
    On Error GoTo Errorcode:
    Screen.MousePointer = 11
    Dim SQL                                                           As String
    SQL = "SELECT CONVERT(VARCHAR, PDIDATE ,101) ,VI_NO,CUSTNAME,PLATENO,SAE,COLOR,ID FROM SMIS_PDI_HDR  "
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute(SQL)
    flex_FillReportView temprs, lvViewVehicles, True
    ShowHidePictureBox2 picViewVehicles, True, picMaster
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdNext_Click()
    On Error GoTo Errorcode:
    RsVCHK.MoveNext
    If RsVCHK.EOF Then
        RsVCHK.MoveLast
        ShowLastRecordMsg
    Else
        StoreMemVars
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "PDI CHECKLIST") = False Then Exit Sub
    On Error GoTo Errorcode:
    Dim SQL                                                           As String
    If MsgBox("Are You Sure You Want To Post This Transaction", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    SQL_STATEMENT = "update smis_pdi_hdr set status='P' where id = " & labid
    gconDMIS.Execute SQL_STATEMENT
    Call NEW_LogAudit("P", "PDI CHECKLIST", SQL_STATEMENT, labid, "", txtPDI_VINNO, "", "")
    MessagePop RecSaveOk, "Transaction Posted", "Record Sucessfully Posted", 1000, 2
    
    rsRefresh
    RsVCHK.Find ("ID=" & labid)
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    RsVCHK.MovePrevious
    If RsVCHK.BOF Then
        RsVCHK.MoveFirst
        ShowFirstRecordMsg
    Else
        StoreMemVars
    End If





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()



    On Error GoTo Errorcode:

    Screen.MousePointer = 11
    rptVehicleCheckList.WindowTitle = Me.Caption
    rptVehicleCheckList.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptVehicleCheckList.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptVehicleCheckList, SMIS_REPORT_PATH & "VehiclesCheckList.rpt", "{D.VI_NO} ='" & txtVI_NO & "'", DMIS_REPORT_Connection, 1
    
    Call NEW_LogAudit("V", "PDI CHECKLIST", "", labid, "", "", "", "")
    Screen.MousePointer = 0





    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdSave_Click()
    Dim vtxtVINo                                                      As String
    Dim vtxtCustName                                                  As String
    Dim vtxtVehiclePlateNo                                            As String
    Dim vtxtVehicleMake                                               As String
    Dim vtxtVehicleModel                                              As String
    Dim vtxtVehicleDescription                                        As String
    Dim vtxtVehicleEngineNo                                           As String
    Dim vtxtVehicleVinNo                                              As String
    Dim vtxtVehicleColor                                              As String
    Dim vtxtVehicleTransmission                                       As String
    Dim vcboSalesAE                                                   As String
    Dim vtxtVehicleModelcode
    Dim vtxtDate                                                      As String
    Dim vtxtTransmission                                              As String
    Dim vtxtSalesAE                                                   As String
    Dim vtxtOdometer                                                  As String
    Dim vtxtFuelGauge                                                 As String
    Dim SQL                                                           As String

    vtxtDate = N2Str2Null(DateValue(txtPDI_Deyt.Value))
    vtxtTransmission = N2Str2Null(cboPDI_Transmission)
    vtxtSalesAE = N2Str2Null(cboPDI_SAE)
    vtxtOdometer = NumericVal(txtPDI_ODOMETER)
    vtxtFuelGauge = N2Str2Null(SetFuelGauage(cboPDI_FuelGauge))
    vtxtVehicleModelcode = N2Str2Null(labModelCode)
    vtxtVINo = N2Str2Null(txtPDI_NO)
    vtxtCustName = N2Str2Null(txtPDI_CusName)
    vtxtVehiclePlateNo = N2Str2Null(txtPDI_PlateNo)
    vtxtVehicleMake = N2Str2Null(txtPDI_Make)
    vtxtVehicleModel = N2Str2Null(txtPDI_Model)
    vtxtVehicleDescription = N2Str2Null(txtModelDescription)
    vtxtVehicleEngineNo = N2Str2Null(txtPDI_EngineNo)
    vtxtVehicleVinNo = N2Str2Null(txtPDI_VINNO)
    vtxtVehicleColor = N2Str2Null(txtPDI_Color)
    vtxtVehicleTransmission = N2Str2Null(cboPDI_Transmission)
    vcboSalesAE = N2Str2Null(cboPDI_SAE)

    If AddorEdit = "ADD" Then
        SQL = " INSERT INTO SMIS_PDI_HDR (PDIDate , VI_NO , CustName, PlateNo, Make,Model, ModelCode,ModelDescription, EngineNo, Vino, Color, Tranmission,  SAE,Odometer , Fuel,  Notes, usercode,lastupdated) VALUES("
        SQL = SQL & N2Str2Null(LOGDATE) & ","
        SQL = SQL & vtxtVINo & ","
        SQL = SQL & vtxtCustName & ","
        SQL = SQL & vtxtVehiclePlateNo & ","
        SQL = SQL & vtxtVehicleMake & ","
        SQL = SQL & vtxtVehicleModel & ","
        SQL = SQL & vtxtVehicleModelcode & ","
        SQL = SQL & vtxtVehicleDescription & ","
        SQL = SQL & vtxtVehicleEngineNo & ","
        SQL = SQL & vtxtVehicleVinNo & ","
        SQL = SQL & vtxtVehicleColor & ","
        SQL = SQL & vtxtVehicleTransmission & ","
        SQL = SQL & vcboSalesAE & ","
        SQL = SQL & vtxtOdometer & ","
        SQL = SQL & vtxtFuelGauge & ","
        SQL = SQL & N2Str2Null(Text1) & ","
        SQL = SQL & N2Str2Null(LOGCODE) & ","
        SQL = SQL & N2Str2Null(LOGDATE) & ")"
        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        Call NEW_LogAudit("A", "PDI CHECKLIST", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtVI_NO), "VI_NO", "SMIS_PDI_HDR"), "", "VI NO: " & txtVI_NO, "", "")
        
        SQL = " INSERT INTO SMIS_PDI_DET " & vbCrLf
        SQL = SQL & "select " & vtxtVINo & " , PDILINEID, 0,'N' AS STATUS from SMIS_vw_PDILookUp where MODELDESCRIPTON=" & vtxtVehicleModel
        gconDMIS.Execute (SQL)
        
        ShowSuccessFullyAdded
    Else
        SQL = " Update SMIS_PDI_HDR SET "
        SQL = SQL & " VI_NO=" & vtxtVINo & " ,"
        SQL = SQL & " CustName=" & vtxtCustName & " ,"
        SQL = SQL & " PlateNo=" & vtxtVehiclePlateNo & " ,"
        SQL = SQL & " Make=" & vtxtVehicleMake & " ,"
        SQL = SQL & " Model=" & vtxtVehicleModel & " ,"
        SQL = SQL & " ModelCode=" & vtxtVehicleModelcode & " ,"
        SQL = SQL & " ModelDescription=" & vtxtVehicleDescription & " ,"
        SQL = SQL & " EngineNo=" & vtxtVehicleEngineNo & " ,"
        SQL = SQL & " PDIDATE=" & vtxtDate & " ,"
        SQL = SQL & " Fuel=" & vtxtFuelGauge & " ,"
        SQL = SQL & " Odometer=" & vtxtOdometer & " ,"
        SQL = SQL & " Notes=" & N2Str2Null(Text1) & " ,"
        SQL = SQL & " Vino=" & vtxtVehicleVinNo & " ,"
        SQL = SQL & " Color=" & vtxtVehicleColor & " ,"
        SQL = SQL & " Tranmission=" & vtxtVehicleTransmission & " ,"
        SQL = SQL & " SAE=" & vcboSalesAE
        SQL = SQL & " WHERE id= " & labid
        gconDMIS.Execute (SQL)
        
        SQL_STATEMENT = SQL
        Call NEW_LogAudit("E", "PDI CHECKLIST", SQL_STATEMENT, labid, "", "VI NO: " & txtVI_NO, "", "")
        
        ShowSuccessFullyUpdated
    End If

    rsRefresh
    If AddorEdit = "EDIT" Then
        RsVCHK.Find ("id=" & labid)
        Dim i                                                         As Long
        Dim ID                                                        As Long
        Dim STATUS                                                    As String
        Dim QTY                                                       As Long

        For i = 1 To Grid1.Rows - 1
            STATUS = SetStatus(Grid1.Cell(i, 4).Text)
            If Not (STATUS = "N" Or STATUS = "") Then
                ID = Grid1.Cell(i, 5).Text
                QTY = Grid1.Cell(i, 3).IntegerValue
                SQL = "update SMIS_PDI_DET SET "
                SQL = SQL & " QTY= " & QTY & ","
                SQL = SQL & " STATUS =" & N2Str2Null(STATUS)
                SQL = SQL & " WHERE PDILINENO=" & ID
                gconDMIS.Execute SQL
            End If
        Next

    End If

    cmdCancel.Value = True

    Exit Sub

    UpdateHeader
    If ISEDITED = True Then
        UpdateDetail
        'LogAudit "E", "PDI CHECKLIST", "CUSTOMER NAME " & txtPDI_CusName & " MODEL & txtPDI_Model " & " VIN" & txtPDI_VINNO
    End If
    picSaves.Visible = False
    picAdds.Visible = True
    AddorEdit = ""
    picMaster.Enabled = False
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    RsVCHK.Requery
    StoreMemVars
End Sub

Private Sub cmdSelect_Click()
    If lvViewVehicles.SelectedRows.Count = 0 Then: Exit Sub
    RsVCHK.MoveFirst
    RsVCHK.Find ("ID=" & lvViewVehicles.SelectedRows.Row(0).Record(7).Value)
    StoreMemVars
    ShowHidePictureBox2 picViewVehicles, False
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UNPost", "PDI CHECKLIST") = False Then Exit Sub

    Dim SQL                                                           As String
    If MsgBox("Are You Sure You Want To Un-Post This Transaction", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    gconDMIS.Execute "update smis_pdi_hdr set status=null where id = " & labid
    MessagePop RecSaveOk, "Transaction Un-Posted", "Record Sucessfully Un-Posted", 1000, 2
    LogAudit "P", "PDI", "VIN NO " & txtPDI_VINNO
    rsRefresh
    RsVCHK.Find ("ID=" & labid)
    StoreMemVars
End Sub

Private Sub Combo1_Change()
    If AddorEdit = "EDIT" Then
        If ISEDITED = True Then
            If MsgBox("Do You Want To Save Changes.", vbYesNo Or vbQuestion Or vbDefaultButton1, App.TITLE) = vbYes Then
                UpdateDetail
                FillGrid SETCATEGORY(cboPDI_Category), SetStatus(Combo1)
                Exit Sub
            End If
        End If
    End If
    FillGrid SETCATEGORY(cboPDI_Category), SetStatus(Combo1)
End Sub

Private Sub Combo1_Click()
    Combo1_Change
End Sub

Private Sub Command1_Click()
    '    If AddorEdit = "EDIT" Then
    If Function_Access(LOGID, "ACESS_SYSTEM", "PDI CHECKLIST") = False Then Exit Sub
    txtPDI_Deyt.Enabled = True: txtPDI_Deyt.SetFocus
    '   End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
    If lvAddVehicles.SelectedRows.Count = 0 Then: Exit Sub


    Dim rsmrrinfo                                                     As ADODB.Recordset
    Dim rscount                                                       As ADODB.Recordset
    Dim rsCust                                                        As ADODB.Recordset
    Set rscount = gconDMIS.Execute("SELECT count(*)  FROM SMIS_vw_PDILookUp where modeldescripton='" & lvAddVehicles.SelectedRows.Row(0).Record(2).Value & "'")
    If rscount.Fields(0).Value = 0 Then
        MsgBox "Cannot Add Perform PDI For this Unit. Please Set Up PDI List for & " & lvAddVehicles.SelectedRows.Row(0).Record(2).Value, vbInformation
        Exit Sub


    End If
    Set rsmrrinfo = gconDMIS.Execute("SELECT * FROM SMIS_MRRINV_TABLE where id=" & lvAddVehicles.SelectedRows.Row(0).Record(5).Value)
    If Not rsmrrinfo.EOF Or Not rsmrrinfo.BOF Then
        txtVI_NO = Null2String(rsmrrinfo!VI_NO)
        Set rsCust = gconDMIS.Execute("select acctname from all_customer_table where CUSCDE= '" & Null2String(rsmrrinfo!CustomerCode) & "'")
        If Not (rsCust.EOF Or rsCust.BOF) Then
            txtPDI_CusName = Null2String(rsCust!AcctName)
        End If

        txtPDI_Make = Null2String(rsmrrinfo!Make)
        txtPDI_Model = Null2String(rsmrrinfo!Model)
        txtModelDescription = Null2String(rsmrrinfo!DESCRIPT)

        labModelCode = Null2String(rsmrrinfo!ModelCode)
        txtPDI_EngineNo = Null2String(rsmrrinfo!EngineNo)
        txtPDI_VINNO = Null2String(rsmrrinfo!VINO)
        txtPDI_Color = Null2String(rsmrrinfo!Color)
        txtPDI_PlateNo = Null2String(rsmrrinfo!ignkey)
        cboPDI_Transmission = Null2String(rsmrrinfo!Transmission)




        ShowHidePictureBox2 picAddVehicles, False
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()

    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitData
    rsRefresh
    '    If RsVCHK.EOF Or RsVCHK.BOF Then
    '        MsgBox " There are No Vehicles Released . Form Will Now Unload ", vbInformation
    '        Unload Me
    '        Exit Sub
    '    End If
    StoreMemVars
    picMaster.Enabled = False
    picSaves.Visible = False
    picAdds.Visible = True
End Sub

Private Sub Grid1_EditRow(ByVal Row As Long)
    ISEDITED = True
End Sub



Private Sub lvViewVehicles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSelect_Click
    End If

End Sub

Private Sub lvViewVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then: Exit Sub
    cmdSelect_Click
End Sub

Private Sub Text2_Change()
    lvAddVehicles.FilterText = txtFilterViewVehicles
    lvAddVehicles.Populate
End Sub

Private Sub tmBlink_Timer()
    If LABSATUS.Caption <> "" Then
        If LABSATUS.Visible = True Then
            LABSATUS.Visible = False
        Else
            LABSATUS.Visible = True
        End If
    End If

End Sub

Private Sub txtFilterViewVehicles_Change()
    lvViewVehicles.FilterText = txtFilterViewVehicles
    lvViewVehicles.Populate
End Sub

Private Sub txtFilterViewVehicles_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPDI_CusName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPDI_ODOMETER_GotFocus()
    If txtPDI_ODOMETER = "0" Then txtPDI_ODOMETER = ""

End Sub

Private Sub txtPDI_ODOMETER_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtPDI_ODOMETER_LostFocus()
    If txtPDI_ODOMETER = "" Then txtPDI_ODOMETER = "0"
End Sub

Private Sub txtPDI_PlateNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

